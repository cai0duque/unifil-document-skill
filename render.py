#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "playwright>=1.40.0",
#     "pyyaml>=6.0",
#     "markdown-it-py>=3.0.0",
#     "pypdf>=4.0.0",
#     "python-docx>=1.1.0",
# ]
# ///
"""
render.py - Gerador de documentos academicos e cientificos em PDF e DOCX
seguindo o modelo oficial da UniFil e as normas da ABNT.

Uso:
    uv run render.py <arquivo.md> [--out <saida.pdf>] [--docx [<saida.docx>]]
    uv run render.py <arquivo.md> --docx-only [--out <saida.docx>]
    uv run render.py <arquivo.md> --mode artigo
    uv run render.py <arquivo.md> --check
    uv run render.py --list-unis
    uv run render.py --list-templates
"""
import argparse
import base64
import html as html_lib
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SKILL_ROOT = Path(__file__).resolve().parent
UNIS_DIR = SKILL_ROOT / "universidades"
DEFAULT_UNI = "UniFil"

# Caracteres proibidos na prosa academica (travessao, en-dash, setas Unicode, emojis)
FORBIDDEN_CHARS_RE = re.compile(
    "[\u2013\u2014\u2190-\u21ff\u2600-\u27bf\u2705\u274c\U0001f300-\U0001faff]"
)


def _strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)
    )


def _norm_key(s: str) -> str:
    s = _strip_accents(s).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


# -- Validacao ABNT e de Redacao --

def audit_academic_markdown(md_text: str, md_path: Path | None = None) -> tuple[list[str], list[str]]:
    """
    Executa validacoes de redacao e conformidade ABNT:
    - ausencia de travessoes, en-dashes, setas Unicode e emojis na prosa;
    - limite de palavras do Resumo (<= 250 para projeto, <= 500 para monografia);
    - pontuacao de Palavras-chave (<= 5, separadas por ';', terminando em '.');
    - siglas na Lista usadas no corpo do texto;
    - referencias citadas no corpo do texto.
    Retorna (errors, warnings).
    """
    errors: list[str] = []
    warnings: list[str] = []

    # 1. Caracteres proibidos e pontuacao de indicativo numerico de secao (NBR 6024:2012) fora de blocos de codigo
    in_code = False
    in_refs = False
    for idx, line in enumerate(md_text.splitlines(), start=1):
        stripped = line.strip()
        if stripped.startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        if re.match(r"^#{1,5}\s+\d+(?:\.\d+)*(?:[\.\-\u2013\u2014]|\s+[\-\u2013\u2014])\s+", stripped):
            errors.append(
                f"Linha {idx}: Pontuacao indevida apos indicativo numerico ({stripped}). "
                "Segundo a ABNT NBR 6024:2012 e o manual da UniFil, nao se usa ponto, hifen ou travessao entre o numero e o titulo da secao (correto: '# 1 INTRODUCAO', nunca '# 1. INTRODUCAO')."
            )
        if re.match(r"^#\s+REFER[ÊE]NCIAS\b", stripped, re.I):
            in_refs = True
            continue
        # Ignora linhas de comentario HTML ou separadores de listas de TOC/Quadro com hifen simples
        if stripped.startswith("<!--"):
            continue
        m = FORBIDDEN_CHARS_RE.search(line)
        if m:
            ch = m.group(0)
            codepoint = f"U+{ord(ch):04X}"
            snippet = stripped if len(stripped) <= 90 else stripped[:87] + "..."
            msg = f"Linha {idx} ({codepoint}): {snippet}"
            if in_refs and ch in ("\u2013", "\u2014"):
                warnings.append(f"[Referencias] {msg} (preservado se for titulo literal da obra)")
            else:
                errors.append(msg)

    # 2. Resumo e Palavras-chave
    resumo_m = re.search(
        r"^#\s+RESUMO\b(.*?)(?=^#\s+|\Z)", md_text, re.M | re.S | re.I
    )
    if resumo_m:
        res_block = resumo_m.group(1)
        kw_m = re.search(r"\*\*Palavras-chave:\*\*\s*(.+)", res_block)
        # Extrai apenas o corpo do resumo (remove referencia inicial, notas e palavras-chave)
        res_lines = []
        for ln in res_block.splitlines():
            s = ln.strip()
            if not s or s == "---":
                continue
            if s.startswith("*(") or s.startswith("**Palavras-chave:**"):
                continue
            if re.match(r"^[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,},\s", s):
                continue
            res_lines.append(s)
        res_words = " ".join(res_lines).split()
        wc = len(res_words)
        if wc > 500:
            errors.append(f"RESUMO com {wc} palavras (excede maximo ABNT de 500 palavras).")
        elif wc > 250:
            warnings.append(
                f"RESUMO com {wc} palavras (ok para monografia [150-500], mas excede 250 palavras para Projeto de Pesquisa UniFil)."
            )
        if kw_m:
            kw_str = kw_m.group(1).strip()
            if not kw_str.endswith("."):
                errors.append("Palavras-chave devem terminar com ponto final (.).")
            kws = [k.strip() for k in kw_str.rstrip(".").split(";") if k.strip()]
            if len(kws) > 5:
                errors.append(f"Palavras-chave possui {len(kws)} termos (maximo permitido: 5).")

    # 3. Corpo textual para checagem de siglas e referencias
    body_m = re.search(
        r"(^#\s+1\s+.*?)(?=^#\s+REFER[ÊE]NCIAS|\Z)", md_text, re.M | re.S | re.I
    )
    body = body_m.group(1) if body_m else ""

    siglas_m = re.search(
        r"^#\s+LISTA DE ABREVIATURAS E SIGLAS\b(.*?)(?=^#\s+|\Z)",
        md_text,
        re.M | re.S | re.I,
    )
    if siglas_m and body:
        for line in siglas_m.group(1).splitlines():
            m = re.match(r"^\s*-\s+\*\*([A-Za-z0-9/\-]+)\*\*", line)
            if not m:
                m = re.match(r"^\s*([A-Za-z0-9/\-]+)\s{2,}", line)
            if m:
                sigla = m.group(1)
                if not re.search(
                    r"(?<![A-Za-zÀ-ÿ])" + re.escape(sigla) + r"s?(?![A-Za-zÀ-ÿ])", body
                ):
                    warnings.append(
                        f"Sigla '{sigla}' listada em Abreviaturas e Siglas nao aparece no corpo textual."
                    )

    refs_m = re.search(r"^#\s+REFER[ÊE]NCIAS\b(.*?)(?=^#\s+AP[ÊE]NDICE|^#\s+ANEXO|\Z)", md_text, re.M | re.S | re.I)
    if refs_m and body:
        for para in refs_m.group(1).strip().split("\n\n"):
            p = para.strip()
            if not p or p == "---":
                continue
            surname = p.split(",")[0].strip()
            if surname and not re.search(re.escape(surname.split()[0]), body, re.I):
                warnings.append(
                    f"Referencia iniciada por '{surname}' nao foi localizada nas citacoes do corpo textual."
                )

    return errors, warnings


def print_audit_report(errors: list[str], warnings: list[str], md_path: Path) -> None:
    if not errors and not warnings:
        print(f" [OK] Validacao ABNT e de redacao aprovada: {md_path}")
        return
    if errors:
        print(f" [ERRO] {len(errors)} problema(s) em {md_path}:")
        for err in errors[:20]:
            print(f"   - {err}")
    if warnings:
        print(f" [AVISO] {len(warnings)} alerta(s) ABNT em {md_path}:")
        for warn in warnings[:20]:
            print(f"   - {warn}")


# -- Carregamento da instituicao e fontes --

def load_uni(slug: str) -> dict:
    uni_dir = UNIS_DIR / slug
    if not uni_dir.is_dir():
        available = sorted(p.name for p in UNIS_DIR.iterdir() if p.is_dir())
        raise SystemExit(
            f"Universidade '{slug}' nao encontrada em {UNIS_DIR}.\n"
            f"Disponiveis: {', '.join(available) or '(nenhuma)'}"
        )

    try:
        import yaml
    except ImportError:
        raise SystemExit("Instale PyYAML: pip install pyyaml (ou execute via: uv run render.py ...)")

    cfg_path = uni_dir / "config.yaml"
    cfg = yaml.safe_load(cfg_path.read_text(encoding="utf-8")) if cfg_path.exists() else {}
    cfg.setdefault("slug", slug)
    cfg.setdefault("display_name", slug)
    cfg.setdefault("short_name", slug)
    cfg.setdefault("city", "Londrina")

    tokens_path = uni_dir / "tokens.css"
    cfg["_tokens_css"] = tokens_path.read_text(encoding="utf-8") if tokens_path.exists() else ""

    style_path = uni_dir / "style.css"
    cfg["_extra_css"] = style_path.read_text(encoding="utf-8") if style_path.exists() else ""

    # Logo oficial
    cfg["_logo_uri"] = ""
    logo_candidates = []
    if cfg.get("logo_file"):
        logo_candidates.append(uni_dir / cfg["logo_file"])
    logo_candidates.extend(
        [
            uni_dir / "unifil-logo-trabalho-academico.png",
            uni_dir / "logo.png",
            uni_dir / "logo.svg",
        ]
    )
    mime_map = {
        ".svg": "image/svg+xml",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
    }
    for p in logo_candidates:
        if p.exists():
            mime = mime_map.get(p.suffix.lower(), "image/png")
            cfg["_logo_uri"] = f"data:{mime};base64,{base64.b64encode(p.read_bytes()).decode()}"
            break

    # Fontes locais (Arial via local + Liberation Sans metrica-identica offline)
    cfg["_font_face_css"] = ""
    fonts_dir = uni_dir / "fontes"
    if fonts_dir.is_dir():
        rules = []
        sans_map = [
            ("LiberationSans-Regular.ttf", "normal", "400", "Arial"),
            ("LiberationSans-Bold.ttf", "normal", "700", "Arial Bold"),
            ("LiberationSans-Italic.ttf", "italic", "400", "Arial Italic"),
            ("LiberationSans-BoldItalic.ttf", "italic", "700", "Arial Bold Italic"),
        ]
        for fname, style, weight, local_name in sans_map:
            fpath = fonts_dir / fname
            if fpath.exists():
                b64 = base64.b64encode(fpath.read_bytes()).decode()
                rules.append(
                    f"@font-face {{\n"
                    f"  font-family: 'Arial';\n"
                    f"  font-style: {style};\n"
                    f"  font-weight: {weight};\n"
                    f"  src: local('{local_name}'), url(data:font/ttf;base64,{b64}) format('truetype');\n"
                    f"}}"
                )
        mono_map = [
            ("LiberationMono-Regular.ttf", "normal", "400"),
            ("LiberationMono-Bold.ttf", "normal", "700"),
        ]
        for fname, style, weight in mono_map:
            fpath = fonts_dir / fname
            if fpath.exists():
                b64 = base64.b64encode(fpath.read_bytes()).decode()
                rules.append(
                    f"@font-face {{\n"
                    f"  font-family: 'Liberation Mono';\n"
                    f"  font-style: {style};\n"
                    f"  font-weight: {weight};\n"
                    f"  src: local('Liberation Mono'), local('Courier New'), url(data:font/ttf;base64,{b64}) format('truetype');\n"
                    f"}}"
                )
        cfg["_font_face_css"] = "\n".join(rules)

    return cfg


# -- Markdown & Frontmatter --

PAGEBREAK_COMMENT_RE = re.compile(r"<!--\s*(?:unifil-|ra-)?pagebreak\s*-->", re.I)
QUADRO_COMMENT_RE = re.compile(
    r"<!--\s*abnt-quadro\s*-->(.*?)<!--\s*/abnt-quadro\s*-->", re.S | re.I
)
MERMAID_BLOCK_RE = re.compile(
    r'<pre><code class="language-mermaid">(.*?)</code></pre>', re.S
)
ALERT_BLOCKQUOTE_RE = re.compile(
    r"<blockquote>\s*<p>\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\](?:[ \t]+([^\n<]+))?\s*(?:\n|<br\s*/?>)?(.*?)</p>(.*?)</blockquote>",
    re.S | re.I,
)
TASK_ITEM_RE = re.compile(r"<li>\s*\[([ xX])\]\s+")

ALERT_DEFAULT_TITLES = {
    "note": "Nota",
    "tip": "Dica",
    "important": "Importante",
    "warning": "Atenção",
    "caution": "Cuidado",
}


def parse_frontmatter(md_text: str) -> tuple[dict, str]:
    lines = md_text.splitlines(keepends=True)
    if not lines or lines[0].rstrip("\r\n") != "---":
        return {}, md_text
    for i in range(1, min(len(lines), 100)):
        if lines[i].rstrip("\r\n") == "---":
            raw_yaml = "".join(lines[1:i])
            body = "".join(lines[i + 1 :]).lstrip("\n")
            try:
                import yaml

                meta = yaml.safe_load(raw_yaml) or {}
                if not isinstance(meta, dict):
                    meta = {}
            except Exception:
                meta = {}
            return meta, body
    return {}, md_text


def _unescape_html(s: str) -> str:
    return (
        s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&#39;", "'")
        .replace("&amp;", "&")
    )


def _preprocess_academic_markdown(md_text: str) -> str:
    """
    Pre-processa o Markdown academico antes da conversao para HTML:
    - converte notacao LaTeX simples ($H_1$, $(H_{1})$, $X^{2}$) em sub/sup HTML;
    - separa alineas consecutivas (a) ..., b) ...) em paragrafos distintos;
    - remove comentarios HTML internos residuais fora de blocos de codigo.
    """
    lines = md_text.splitlines()
    out: list[str] = []
    in_code = False

    def _convert_inline_math(segment: str) -> str:
        # Converte $(H_{1})$ ou $H_{1}$ ou $H_1$ em H<sub>1</sub>
        def _repl(m: re.Match) -> str:
            expr = m.group(1).strip()
            expr = re.sub(r"([A-Za-zÀ-ÿ0-9]+)_\{([^}]+)\}", r"\1<sub>\2</sub>", expr)
            expr = re.sub(r"([A-Za-zÀ-ÿ0-9]+)_([A-Za-z0-9]+)", r"\1<sub>\2</sub>", expr)
            expr = re.sub(r"([A-Za-zÀ-ÿ0-9]+)\^\{([^}]+)\}", r"\1<sup>\2</sup>", expr)
            expr = re.sub(r"([A-Za-zÀ-ÿ0-9]+)\^([A-Za-z0-9]+)", r"\1<sup>\2</sup>", expr)
            return expr

        # Apenas expressoes curtas inline ($...$) sem quebras de linha
        return re.sub(r"\$([^$\n]{1,40})\$", _repl, segment)

    for line in lines:
        if line.strip().startswith("```"):
            in_code = not in_code
            out.append(line)
            continue
        if in_code:
            out.append(line)
            continue

        # Preserva codigo inline (`...`) ao converter math inline
        parts = re.split(r"(`[^`]+`)", line)
        converted_line = "".join(
            p if p.startswith("`") else _convert_inline_math(p) for p in parts
        )

        # Se for uma alinea (a) ..., b) ...) imediatamente apos outra linha nao vazia, insere linha em branco
        if (
            re.match(r"^[a-z]\)\s+", converted_line.strip())
            and out
            and out[-1].strip() != ""
            and not out[-1].strip().startswith(">")
        ):
            out.append("")

        out.append(converted_line)

    return "\n".join(out)


def render_md_fragment(md_text: str) -> str:
    try:
        from markdown_it import MarkdownIt
    except ImportError:
        raise SystemExit(
            "Instale markdown-it-py: pip install markdown-it-py (ou execute via: uv run render.py ...)"
        )
    md_text = _preprocess_academic_markdown(md_text)
    md_text = PAGEBREAK_COMMENT_RE.sub('<div class="unifil-pagebreak"></div>', md_text)
    md_text = QUADRO_COMMENT_RE.sub(
        lambda m: f'\n<div class="abnt-quadro">\n\n{m.group(1).strip()}\n\n</div>\n',
        md_text,
    )
    md = MarkdownIt("commonmark", {"typographer": False, "html": True}).enable(
        ["table", "strikethrough"]
    )
    html = md.render(md_text)
    html = MERMAID_BLOCK_RE.sub(
        lambda m: f'<pre class="mermaid">{_unescape_html(m.group(1)).strip()}</pre>',
        html,
    )

    def _alert_repl(m: re.Match) -> str:
        kind = m.group(1).lower()
        custom_title = (m.group(2) or "").strip()
        first_p_rest = (m.group(3) or "").strip()
        other_p = (m.group(4) or "").strip()
        title_text = custom_title or ALERT_DEFAULT_TITLES.get(kind, kind.upper())
        inner = f"<p>{first_p_rest}</p>" if first_p_rest else ""
        if other_p:
            inner += other_p
        return (
            f'<blockquote class="ra-alert ra-alert-{kind}">'
            f'<span class="ra-alert-title">{title_text}</span>'
            f"{inner}</blockquote>"
        )

    html = ALERT_BLOCKQUOTE_RE.sub(_alert_repl, html)
    html = TASK_ITEM_RE.sub(
        lambda m: f'<li class="ra-task-item"><span class="ra-checkbox{" ra-checkbox-checked" if m.group(1).lower() == "x" else ""}"></span>',
        html,
    )
    return html


def split_top_h1_sections(md_text: str) -> list[tuple[str | None, str]]:
    """Divide o Markdown pelas secoes de nivel 1 (# TITULO), ignorando blocos de codigo."""
    sections: list[tuple[str | None, str]] = []
    current_title: str | None = None
    current_lines: list[str] = []
    in_code = False

    for line in md_text.splitlines():
        if line.strip().startswith("```"):
            in_code = not in_code
            current_lines.append(line)
            continue
        if not in_code and re.match(r"^#\s+(.+)$", line):
            if current_title is not None or "\n".join(current_lines).strip():
                sections.append((current_title, "\n".join(current_lines).strip()))
            current_title = re.match(r"^#\s+(.+)$", line).group(1).strip()
            current_lines = []
        else:
            current_lines.append(line)

    if current_title is not None or "\n".join(current_lines).strip():
        sections.append((current_title, "\n".join(current_lines).strip()))

    # Remove separadores --- soltos ao final de cada secao
    cleaned: list[tuple[str | None, str]] = []
    for title, body in sections:
        lines = body.splitlines()
        while lines and lines[-1].strip() in ("", "---"):
            lines.pop()
        while lines and lines[0].strip() in ("", "---"):
            lines.pop(0)
        cleaned.append((title, "\n".join(lines)))
    return cleaned


def _strip_md_inline(s: str) -> str:
    s = s.replace(r"\_", "_").replace(r"\*", "*")
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"__(.+?)__", r"\1", s)
    s = re.sub(r"\*(.+?)\*", r"\1", s)
    s = re.sub(r"(?<!\w)_([^_\n]+?)_(?!\w)", r"\1", s)
    s = re.sub(r"`(.+?)`", r"\1", s)
    return s.strip()


def _clean_helper_notes(md_block: str) -> str:
    out = []
    for line in md_block.splitlines():
        s = line.strip()
        if re.match(r"^[*_]\(Nota\s+de\s+formata", s, re.I):
            continue
        if re.match(r"^<!--.*?-->$", s):
            continue
        out.append(line)
    return "\n".join(out)


def _render_capa_html(body_md: str, logo_uri: str, org: dict) -> str:
    """Converte o bloco # CAPA no layout oficial da Capa UniFil com logo (somente titulo em portugues)."""
    body_md = _clean_helper_notes(body_md)
    paras = [p.strip() for p in re.split(r"\n\s*\n", body_md) if p.strip()]
    author = ""
    title = ""
    bottom_lines: list[str] = []

    if paras:
        author = _strip_md_inline(paras[0])
    if len(paras) >= 2:
        title = _strip_md_inline(paras[1])
    for p in paras[2:]:
        clean = _strip_md_inline(p)
        # Ignora eventuais subtitulos em lingua estrangeira colocados entre parenteses na capa
        if clean.startswith("(") and clean.endswith(")"):
            continue
        bottom_lines.append(p.replace("  \n", "<br>").replace("\n", "<br>"))

    # Separa titulo principal : subtitulo se houver
    if ":" in title:
        main_t, sub_t = title.split(":", 1)
        title_html = (
            f'<div class="unifil-cover-title">{html_lib.escape(main_t.strip())}:</div>'
            f'<div class="unifil-cover-subtitle">{html_lib.escape(sub_t.strip())}</div>'
        )
    else:
        title_html = f'<div class="unifil-cover-title">{html_lib.escape(title)}</div>'

    logo_html = (
        f'<div class="unifil-cover-logo-wrap"><img src="{logo_uri}" alt="{org.get("short_name", "UniFil")}" class="unifil-cover-logo"></div>'
        if logo_uri
        else ""
    )
    bottom_html = "<br>".join(bottom_lines) if bottom_lines else f'{org.get("city", "Londrina")}<br>2026'

    return f"""
<section class="unifil-cover">
  <div>
    {logo_html}
    <div class="unifil-cover-author">{html_lib.escape(author)}</div>
  </div>
  <div class="unifil-cover-center">
    {title_html}
  </div>
  <div class="unifil-cover-bottom">
    {bottom_html}
  </div>
</section>
"""


def _render_folha_rosto_html(body_md: str, org: dict) -> str:
    body_md = _clean_helper_notes(body_md)
    paras = [p.strip() for p in re.split(r"\n\s*\n", body_md) if p.strip()]
    author = _strip_md_inline(paras[0]) if paras else ""
    title = _strip_md_inline(paras[1]) if len(paras) >= 2 else ""
    natureza_paras = []
    bottom_html = ""

    for p in paras[2:]:
        clean = _strip_md_inline(p)
        if clean.startswith("(") and clean.endswith(")"):
            continue
        elif p.startswith(">"):
            # Bloco de natureza / orientador
            inner = "\n".join(re.sub(r"^>\s?", "", ln) for ln in p.splitlines())
            natureza_paras.append(render_md_fragment(inner))
        else:
            bottom_html = p.replace("  \n", "<br>").replace("\n", "<br>")

    if ":" in title:
        main_t, sub_t = title.split(":", 1)
        title_html = (
            f'<div class="unifil-cover-title">{html_lib.escape(main_t.strip())}:</div>'
            f'<div class="unifil-cover-subtitle">{html_lib.escape(sub_t.strip())}</div>'
        )
    else:
        title_html = f'<div class="unifil-cover-title">{html_lib.escape(title)}</div>'

    natureza_html = "".join(natureza_paras)

    return f"""
<section class="unifil-pretext-page">
  <div class="unifil-title-page-top">
    <div class="unifil-title-page-author">{html_lib.escape(author)}</div>
  </div>
  <div class="unifil-title-page-middle">
    {title_html}
  </div>
  <div class="unifil-natureza-wrap">
    {natureza_html}
  </div>
  <div class="unifil-cover-bottom" style="text-align:center;">
    {bottom_html or (org.get("city", "Londrina") + "<br>2026")}
  </div>
</section>
"""


def _render_banca_html(body_md: str) -> str:
    paras = [p.strip() for p in re.split(r"\n\s*\n", body_md) if p.strip()]
    author = _strip_md_inline(paras[0]) if paras else ""
    title = _strip_md_inline(paras[1]) if len(paras) >= 2 else ""
    natureza_html = ""
    membros_html = []
    data_html = ""

    i = 2
    while i < len(paras):
        p = paras[i]
        if p.startswith(">"):
            inner = "\n".join(re.sub(r"^>\s?", "", ln) for ln in p.splitlines())
            natureza_html += render_md_fragment(inner)
        elif "BANCA EXAMINADORA" in p.upper():
            pass
        elif p.strip() == "---" and i + 1 < len(paras):
            lines = [ln.strip() for ln in paras[i + 1].splitlines() if ln.strip()]
            rest = "<br>".join(html_lib.escape(_strip_md_inline(ln.rstrip(" \\"))) for ln in lines)
            membros_html.append(
                f'<div class="unifil-banca-membro"><span class="unifil-banca-line"></span>{rest}</div>'
            )
            i += 1
        elif re.match(r"^(?:_{4,}|-{3,})", _strip_md_inline(p)):
            lines = [ln.strip() for ln in p.splitlines() if ln.strip()]
            rest = "<br>".join(html_lib.escape(_strip_md_inline(ln.rstrip(" \\"))) for ln in lines[1:])
            membros_html.append(
                f'<div class="unifil-banca-membro"><span class="unifil-banca-line"></span>{rest}</div>'
            )
        else:
            data_html = html_lib.escape(_strip_md_inline(p))
        i += 1

    if ":" in title:
        main_t, sub_t = title.split(":", 1)
        title_html = (
            f'<div class="unifil-cover-title">{html_lib.escape(main_t.strip())}:</div>'
            f'<div class="unifil-cover-subtitle">{html_lib.escape(sub_t.strip())}</div>'
        )
    else:
        title_html = f'<div class="unifil-cover-title">{html_lib.escape(title)}</div>'

    return f"""
<section class="unifil-pretext-page">
  <div class="unifil-title-page-top">
    <div class="unifil-title-page-author">{html_lib.escape(author)}</div>
  </div>
  <div style="text-align:center; margin-top: 16pt;">
    {title_html}
  </div>
  <div class="unifil-natureza-wrap" style="margin-top:14pt; margin-bottom:10pt;">
    {natureza_html}
  </div>
  <div>
    <div class="unifil-banca-heading">BANCA EXAMINADORA</div>
    {"".join(membros_html)}
  </div>
  <div class="unifil-banca-data">{data_html}</div>
</section>
"""


def _render_resumo_html(title: str, body_md: str) -> str:
    body_md = _clean_helper_notes(body_md)
    paras = [p.strip() for p in re.split(r"\n\s*\n", body_md) if p.strip()]
    ref_html = ""
    kw_html = ""
    body_paras = []

    for i, p in enumerate(paras):
        if i == 0 and re.match(r"^[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,},\s", p):
            ref_html = f'<div class="unifil-resumo-ref">{render_md_fragment(p)}</div>'
        elif re.match(r"^\*\*(Palavras-chave|Keywords):\*\*", p, re.I):
            kw_html = f'<div class="unifil-keywords">{render_md_fragment(p)}</div>'
        else:
            body_paras.append(render_md_fragment(p))

    return f"""
<section class="unifil-pre-section">
  {ref_html}
  <h2 class="unifil-section-center-title">{html_lib.escape(title)}</h2>
  <div class="unifil-resumo-body">
    {"".join(body_paras)}
  </div>
  {kw_html}
</section>
"""


def _render_siglas_html(title: str, body_md: str) -> str:
    rows = []
    for line in body_md.splitlines():
        s = line.strip()
        if not s:
            continue
        m = re.match(r"^-\s+\*\*(.+?)\*\*\s*[-–]\s*(.+)$", s)
        if not m:
            m = re.match(r"^([A-Za-z0-9/\-]+)\s{2,}(.+)$", s)
        if m:
            rows.append(
                f"<tr><td>{html_lib.escape(m.group(1).strip())}</td>"
                f"<td>{html_lib.escape(m.group(2).strip())}</td></tr>"
            )
    return f"""
<section class="unifil-pre-section">
  <h2 class="unifil-section-center-title">{html_lib.escape(title)}</h2>
  <table class="unifil-siglas-table">
    <tbody>{"".join(rows)}</tbody>
  </table>
</section>
"""


def _render_toc_or_list_html(title: str, body_md: str, page_map: dict[str, int] | None = None) -> str:
    rows = []
    for line in body_md.splitlines():
        s = line.strip()
        if not s:
            continue
        s = re.sub(r"^-\s+", "", s).rstrip("\\ ")
        # Separa titulo e placeholder de pagina (.......... p. X ou p. X ou 00)
        m = re.match(r"^(.*?)\s*(?:\.{2,}\s*)?(?:p\.\s*[Xx0-9]+|\b00\b)$", s)
        label = m.group(1).strip() if m else s
        page_str = "X"
        if m and not re.search(r"p\.\s*[Xx]|00$", s):
            pm = re.search(r"(\d+)$", s)
            if pm:
                page_str = pm.group(1)

        if page_map:
            nk = _norm_key(label)
            if nk in page_map:
                page_str = str(page_map[nk])

        # Nivel hierarquico para Sumario e separacao de indicativo numerico + titulo
        lvl = 1
        num_str = ""
        title_text = _strip_md_inline(label)
        num_m = re.match(r"^(\d+(?:\.\d+)*)\s+(.+)$", title_text)
        if num_m:
            num_str = num_m.group(1)
            title_text = num_m.group(2).strip()
            lvl = min(len(num_str.split(".")), 5)

        if num_str:
            rows.append(
                f'<tr class="unifil-toc-row level-{lvl}">'
                f'<td class="unifil-toc-num">{html_lib.escape(num_str)}</td>'
                f'<td class="unifil-toc-title-cell"><div class="unifil-toc-line">'
                f'<span class="unifil-toc-title">{html_lib.escape(title_text)}</span>'
                f'<span class="unifil-toc-dots"></span>'
                f'</div></td>'
                f'<td class="unifil-toc-page">{html_lib.escape(page_str)}</td>'
                f'</tr>'
            )
        else:
            rows.append(
                f'<tr class="unifil-toc-row level-{lvl}">'
                f'<td class="unifil-toc-title-cell" colspan="2"><div class="unifil-toc-line">'
                f'<span class="unifil-toc-title">{html_lib.escape(title_text)}</span>'
                f'<span class="unifil-toc-dots"></span>'
                f'</div></td>'
                f'<td class="unifil-toc-page">{html_lib.escape(page_str)}</td>'
                f'</tr>'
            )

    return f"""
<section class="unifil-pre-section">
  <h2 class="unifil-section-center-title">{html_lib.escape(title)}</h2>
  <table class="unifil-toc-table">
    <tbody>
      {"".join(rows)}
    </tbody>
  </table>
</section>
"""


def _postprocess_textual_html(html: str) -> str:
    """Ajusta legendas de Quadro/Tabela/Figura, fontes, alineas, Quadros fechados e blocos coesos."""
    # Legendas: paragrafos iniciados por Quadro N, Tabela N, Figura N
    html = re.sub(
        r"<p>(<strong>(?:Quadro|Tabela|Figura|Gráfico)\s+\d+[^<]*</strong>[^<]*)</p>",
        r'<p class="unifil-caption">\1</p>',
        html,
    )
    html = re.sub(
        r"<p>((?:Quadro|Tabela|Figura|Gráfico)\s+\d+\s*[-–:].*?)</p>",
        r'<p class="unifil-caption">\1</p>',
        html,
    )
    # Fontes
    html = re.sub(
        r"<p>((?:<strong>)?Fonte:(?:</strong>)?.*?)</p>",
        r'<p class="unifil-source">\1</p>',
        html,
    )
    # Alineas a), b), c)...
    html = re.sub(
        r"<p>([a-z]\)\s+.*?)</p>",
        r'<p class="unifil-alinea">\1</p>',
        html,
        flags=re.S,
    )
    # Marca tabelas imediatamente apos legenda de Quadro como abnt-quadro (grade fechada)
    # e agrupa legenda + tabela + fonte em bloco coeso para a legenda nao ficar orfa
    html = re.sub(
        r'(<p class="unifil-caption">.*?Quadro\s+\d+.*?</p>)\s*<table>(.*?)</table>\s*(<p class="unifil-source">.*?</p>)?',
        lambda m: (
            '<div class="abnt-float-block">'
            + m.group(1)
            + '<table class="abnt-quadro">'
            + m.group(2)
            + '</table>'
            + (m.group(3) or "")
            + '</div>'
        ),
        html,
        flags=re.S,
    )
    html = re.sub(
        r'(<p class="unifil-caption">.*?Tabela\s+\d+.*?</p>)\s*<table>(.*?)</table>\s*(<p class="unifil-source">.*?</p>)?',
        lambda m: (
            '<div class="abnt-float-block">'
            + m.group(1)
            + '<table>'
            + m.group(2)
            + '</table>'
            + (m.group(3) or "")
            + '</div>'
        ),
        html,
        flags=re.S,
    )
    return html


def build_split_html(
    md_text: str,
    org: dict,
    title: str = "Documento Acadêmico",
    mode: str | None = None,
    page_map: dict[str, int] | None = None,
) -> tuple[str | None, str]:
    """
    Retorna (capa_html_doc, miolo_html_doc).
    Separar a capa garante que o contador de paginas do miolo inicie exatamente em 1 na Folha de Rosto,
    ocultando nos pre-textuais e exibindo a partir de 1 INTRODUCAO (NBR 14724 / UniFil).
    """
    frontmatter, md_without_fm = parse_frontmatter(md_text)
    resolved_mode = (mode or str(frontmatter.get("mode", "")) or "monografia").lower()
    if resolved_mode in ("artigo", "article", "simples"):
        mode_cls = "unifil-mode-artigo"
    else:
        mode_cls = "unifil-mode-monografia"

    sections = split_top_h1_sections(md_without_fm)
    logo_uri = org.get("_logo_uri", "")

    cover_fragment: str | None = None
    pretext_fragments: list[str] = []
    textual_fragments: list[str] = []
    post_fragments: list[str] = []

    textual_h1_count = 0

    for sec_title, sec_body in sections:
        if sec_title is None:
            if sec_body.strip():
                textual_fragments.append(_postprocess_textual_html(render_md_fragment(sec_body)))
            continue

        norm = _strip_accents(sec_title).upper().strip()

        if norm == "CAPA":
            cover_fragment = _render_capa_html(sec_body, logo_uri, org)
        elif norm == "FOLHA DE ROSTO":
            pretext_fragments.append(_render_folha_rosto_html(sec_body, org))
        elif norm == "BANCA EXAMINADORA":
            pretext_fragments.append(_render_banca_html(sec_body))
        elif norm in ("DEDICATORIA",):
            pretext_fragments.append(
                f'<section class="unifil-dedicatoria"><div class="unifil-dedicatoria-inner">{render_md_fragment(sec_body)}</div></section>'
            )
        elif norm.startswith("AGRADECIMENTO"):
            pretext_fragments.append(
                f'<section class="unifil-pre-section unifil-agradecimentos"><h2 class="unifil-section-center-title">{html_lib.escape(sec_title)}</h2>{render_md_fragment(sec_body)}</section>'
            )
        elif norm in ("RESUMO", "ABSTRACT"):
            pretext_fragments.append(_render_resumo_html(sec_title, sec_body))
        elif norm in ("LISTA DE ABREVIATURAS E SIGLAS", "LISTA DE SIGLAS"):
            pretext_fragments.append(_render_siglas_html(sec_title, sec_body))
        elif norm.startswith("LISTA DE ") or norm == "SUMARIO":
            pretext_fragments.append(_render_toc_or_list_html(sec_title, sec_body, page_map))
        elif norm == "REFERENCIAS":
            post_fragments.append(
                f'<section class="unifil-referencias"><h2 class="unifil-section-center-title">{html_lib.escape(sec_title)}</h2>{render_md_fragment(sec_body)}</section>'
            )
        elif norm.startswith("APENDICE") or norm.startswith("ANEXO"):
            post_fragments.append(
                f'<section class="unifil-referencias"><h2 class="unifil-section-center-title">{html_lib.escape(sec_title)}</h2>{_postprocess_textual_html(render_md_fragment(sec_body))}</section>'
            )
        else:
            textual_h1_count += 1
            cls = ' class="unifil-h1-break"' if textual_h1_count > 1 else ""
            h1_html = f"<h1{cls}>{html_lib.escape(sec_title)}</h1>"
            body_html = _postprocess_textual_html(render_md_fragment(sec_body))
            textual_fragments.append(h1_html + "\n" + body_html)

    font_face_css = org.get("_font_face_css", "")
    tokens_css = org.get("_tokens_css", "")
    extra_css = org.get("_extra_css", "")

    def _wrap_html(body_inner: str, include_mermaid: bool = False) -> str:
        mermaid_script = ""
        if include_mermaid and 'class="mermaid"' in body_inner:
            mjs_path = SKILL_ROOT / "assets" / "mermaid.min.js"
            if mjs_path.exists():
                mermaid_script = (
                    "<script>"
                    + mjs_path.read_text(encoding="utf-8")
                    + "</script>\n"
                    "<script>\n"
                    "mermaid.initialize({startOnLoad:false, theme:'neutral', securityLevel:'loose', fontFamily:'Arial, sans-serif'});\n"
                    "window.__mermaidDone = false;\n"
                    "(async () => { try { await mermaid.run(); } catch (e) { console.error(e); } window.__mermaidDone = true; })();\n"
                    "</script>"
                )
        return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>{html_lib.escape(title)}</title>
<style>
{font_face_css}
{tokens_css}
{extra_css}
</style>
</head>
<body class="{mode_cls}">
{body_inner}
{mermaid_script}
</body>
</html>"""

    capa_doc = _wrap_html(cover_fragment) if cover_fragment else None
    miolo_inner = (
        "\n".join(pretext_fragments)
        + '\n<div class="unifil-textual">\n'
        + "\n".join(textual_fragments)
        + "\n</div>\n"
        + "\n".join(post_fragments)
    )
    miolo_doc = _wrap_html(miolo_inner, include_mermaid=True)
    return capa_doc, miolo_doc


# -- Renderizacao via Chrome/Playwright --

def find_system_browser() -> str | None:
    candidates: list[str] = []
    if sys.platform == "win32":
        local_app = os.environ.get("LOCALAPPDATA", "")
        prog_files = os.environ.get("PROGRAMFILES", r"C:\Program Files")
        prog_x86 = os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")
        candidates.extend(
            [
                os.path.join(prog_files, "Google", "Chrome", "Application", "chrome.exe"),
                os.path.join(prog_x86, "Google", "Chrome", "Application", "chrome.exe"),
                os.path.join(local_app, "Google", "Chrome", "Application", "chrome.exe"),
                os.path.join(prog_files, "Microsoft", "Edge", "Application", "msedge.exe"),
            ]
        )
    elif sys.platform == "darwin":
        candidates.extend(
            [
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Chromium.app/Contents/MacOS/Chromium",
                "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            ]
        )
    else:
        for cmd in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "microsoft-edge"):
            found = shutil.which(cmd)
            if found:
                candidates.append(found)

    for path in candidates:
        if path and Path(path).is_file():
            return path
    return None


def _render_html_to_pdf(page, html_str: str, out_pdf: Path) -> None:
    with tempfile.NamedTemporaryFile(
        suffix=".html", mode="w", encoding="utf-8", delete=False
    ) as tmp:
        tmp.write(html_str)
        tmp_path = Path(tmp.name)
    try:
        page.goto(tmp_path.resolve().as_uri(), wait_until="networkidle", timeout=20000)
        try:
            page.wait_for_function(
                "window.__mermaidDone === true || document.querySelectorAll('pre.mermaid').length === 0",
                timeout=15000,
            )
        except Exception:
            pass
        page.evaluate("() => document.fonts.ready")
        page.pdf(
            path=str(out_pdf),
            format="A4",
            margin={"top": "30mm", "right": "20mm", "bottom": "20mm", "left": "30mm"},
            print_background=True,
            prefer_css_page_size=True,
            display_header_footer=False,
        )
    finally:
        tmp_path.unlink(missing_ok=True)


def _extract_toc_page_map(miolo_pdf: Path, md_text: str) -> dict[str, int]:
    """Le o PDF do miolo (Pass 1) para descobrir a pagina ABNT exata de cada titulo, quadro e tabela."""
    try:
        from pypdf import PdfReader
    except ImportError:
        return {}

    reader = PdfReader(str(miolo_pdf))
    pages_norm = [_norm_key(page.extract_text() or "") for page in reader.pages]

    # Coleta alvos do Sumario, Lista de Quadros, Lista de Tabelas
    targets: list[str] = []
    _, md_without_fm = parse_frontmatter(md_text)
    for sec_title, sec_body in split_top_h1_sections(md_without_fm):
        if sec_title is None:
            continue
        norm = _strip_accents(sec_title).upper().strip()
        if norm == "SUMARIO" or norm.startswith("LISTA DE "):
            for line in sec_body.splitlines():
                s = re.sub(r"^-\s+", "", line.strip()).rstrip("\\ ")
                m = re.match(r"^(.*?)\s*(?:\.{2,}\s*)?(?:p\.\s*[Xx0-9]+|\b00\b)$", s)
                lbl = m.group(1).strip() if m else s
                if lbl:
                    targets.append(lbl)

    # Descobre a primeira pagina textual (onde aparece "1 introducao" ou primeiro H1 textual)
    start_idx = 0
    for idx, ptxt in enumerate(pages_norm):
        if ptxt.startswith("1 ") or " 1 introducao " in f" {ptxt} ":
            # Ignora o proprio Sumario (que tambem contem "1 introducao") procurando a ultima ocorrencia ou pagina sem "sumario"
            if "sumario" not in ptxt[:60]:
                start_idx = idx
                break

    page_map: dict[str, int] = {}
    for lbl in targets:
        nk = _norm_key(lbl)
        if not nk:
            continue
        for idx in range(start_idx, len(pages_norm)):
            if nk in pages_norm[idx]:
                page_map[nk] = idx + 1  # 1-based ABNT (Folha de Rosto = 1)
                break
    return page_map


def generate_academic_pdf(
    md_text: str,
    org: dict,
    out_pdf: Path,
    title: str = "Documento Acadêmico",
    mode: str | None = None,
    chrome_bin: str | None = None,
    sync_toc: bool = True,
) -> dict[str, int]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise SystemExit(
            "Instale Playwright: pip install playwright (ou execute via: uv run render.py ...)"
        )

    capa_html, miolo_html_p1 = build_split_html(md_text, org, title=title, mode=mode)
    page_map: dict[str, int] = {}

    with sync_playwright() as p:
        launch_args: dict = {}
        if chrome_bin:
            launch_args["executable_path"] = chrome_bin
        try:
            browser = p.chromium.launch(**launch_args)
        except Exception as exc:
            if not chrome_bin:
                detected = find_system_browser()
                if detected:
                    print(f"  Usando navegador do sistema detectado: {detected}")
                    browser = p.chromium.launch(executable_path=detected)
                else:
                    raise SystemExit(
                        f"Nao foi possivel iniciar o Chromium do Playwright ({exc}).\n"
                        "Execute 'playwright install chromium' ou informe --chrome-bin."
                    )
            else:
                raise

        page = browser.new_page()
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_dir = Path(tmpdir)
            miolo_p1_pdf = tmp_dir / "miolo_p1.pdf"
            _render_html_to_pdf(page, miolo_html_p1, miolo_p1_pdf)

            final_miolo_pdf = miolo_p1_pdf
            if sync_toc:
                page_map = _extract_toc_page_map(miolo_p1_pdf, md_text)
                if page_map:
                    _, miolo_html_p2 = build_split_html(
                        md_text, org, title=title, mode=mode, page_map=page_map
                    )
                    miolo_p2_pdf = tmp_dir / "miolo_p2.pdf"
                    _render_html_to_pdf(page, miolo_html_p2, miolo_p2_pdf)
                    final_miolo_pdf = miolo_p2_pdf

            if capa_html:
                capa_pdf = tmp_dir / "capa.pdf"
                _render_html_to_pdf(page, capa_html, capa_pdf)
                try:
                    from pypdf import PdfWriter

                    writer = PdfWriter()
                    writer.append(str(capa_pdf))
                    writer.append(str(final_miolo_pdf))
                    with open(out_pdf, "wb") as f_out:
                        writer.write(f_out)
                except ImportError:
                    shutil.copy2(final_miolo_pdf, out_pdf)
            else:
                shutil.copy2(final_miolo_pdf, out_pdf)

        browser.close()
    return page_map


def list_unis() -> list[str]:
    return sorted(p.name for p in UNIS_DIR.iterdir() if p.is_dir())


def list_templates(uni_slug: str) -> list[Path]:
    tpl_dir = UNIS_DIR / uni_slug / "templates"
    if not tpl_dir.is_dir():
        return []
    return sorted(tpl_dir.glob("*.md"))


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Gerador de documentos academicos em PDF (UniFil / ABNT)."
    )
    parser.add_argument("md", nargs="?", help="Arquivo Markdown de entrada")
    parser.add_argument(
        "--uni", "--org", default=DEFAULT_UNI, dest="uni",
        help=f"Instituicao em universidades/ (default: {DEFAULT_UNI})",
    )
    parser.add_argument("--out", "-o", default=None, help="Caminho do PDF de saida")
    parser.add_argument("--title", default=None, help="Titulo interno do PDF")
    parser.add_argument(
        "--mode", choices=["monografia", "artigo"], default=None,
        help="monografia (nova pagina por H1) ou artigo (fluxo continuo)",
    )
    parser.add_argument(
        "--chrome-bin", default=os.environ.get("CHROME_BIN"),
        help="Caminho do Chrome/Chromium/Edge",
    )
    parser.add_argument(
        "--check", action="store_true",
        help="Valida regras ABNT e redacao sem gerar PDF",
    )
    parser.add_argument(
        "--no-toc-sync", action="store_true",
        help="Nao preenche automaticamente os numeros de pagina do Sumario",
    )
    parser.add_argument(
        "--docx", nargs="?", const="__AUTO__", default=None,
        help="Exporta tambem uma versao .docx (opcionalmente informando o caminho)",
    )
    parser.add_argument(
        "--docx-only", action="store_true",
        help="Gera apenas o arquivo .docx sem gerar PDF",
    )
    parser.add_argument("--list-unis", "--list-orgs", action="store_true", dest="list_unis", help="Lista instituicoes e sai")
    parser.add_argument("--list-templates", action="store_true", help="Lista templates disponiveis e sai")

    args = parser.parse_args()

    if args.list_unis:
        for slug in list_unis():
            print(slug)
        return

    if args.list_templates:
        for tpl in list_templates(args.uni):
            print(tpl.relative_to(SKILL_ROOT))
        return

    if not args.md:
        parser.error("informe o arquivo Markdown de entrada: <arquivo.md>")

    md_path = Path(args.md).resolve()
    if not md_path.exists():
        sys.exit(f"Erro: arquivo nao encontrado: {md_path}")

    md_text = md_path.read_text(encoding="utf-8")
    errors, warnings = audit_academic_markdown(md_text, md_path)

    if args.check:
        print_audit_report(errors, warnings, md_path)
        sys.exit(1 if errors else 0)

    org = load_uni(args.uni)
    title = args.title or md_path.stem.replace("-", " ").replace("_", " ").title()
    out_pdf = Path(args.out) if args.out else (md_path.parent / f"{md_path.stem}.pdf")
    out_pdf.parent.mkdir(parents=True, exist_ok=True)

    print(f" Lendo: {md_path}")
    print(f" Instituicao: {org['display_name']} ({org['slug']})")
    if errors or warnings:
        print_audit_report(errors, warnings, md_path)

    page_map: dict[str, int] = {}
    if not args.docx_only:
        print(" Renderizando PDF ABNT/UniFil via Playwright...")
        page_map = generate_academic_pdf(
            md_text,
            org,
            out_pdf,
            title=title,
            mode=args.mode,
            chrome_bin=args.chrome_bin,
            sync_toc=not args.no_toc_sync,
        )
        print(f" PDF gerado: {out_pdf}")

    if args.docx is not None or args.docx_only:
        from docx_export import export_markdown_to_docx

        if args.docx and args.docx != "__AUTO__":
            out_docx = Path(args.docx)
        elif args.docx_only and args.out:
            out_docx = Path(args.out)
        else:
            out_docx = out_pdf.with_suffix(".docx")

        print(" Exportando DOCX ABNT/UniFil...")
        export_markdown_to_docx(
            md_text,
            org,
            out_docx,
            mode=args.mode,
            page_map=page_map or None,
            uni_dir=UNIS_DIR / args.uni,
        )
        print(f" DOCX gerado: {out_docx}")


if __name__ == "__main__":
    main()
