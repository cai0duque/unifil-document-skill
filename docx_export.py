"""
docx_export.py - Exportador de Markdown academico UniFil/ABNT para .docx oficial.
Aplica margens ABNT (3 cm sup/esq, 2 cm inf/dir), fonte Arial, logo oficial na capa,
paginacao contada a partir da Folha de Rosto e exibida a partir da Introducao,
hierarquia de secoes NBR 6024, tabelas IBGE, quadros fechados e referencias NBR 6023.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor


def _strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)
    )


def _norm_key(s: str) -> str:
    s = _strip_accents(s).lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def _set_run_font(run, name: str = "Arial", size_pt: float = 12, bold: bool = False, italic: bool = False) -> None:
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), name)
    rFonts.set(qn("w:eastAsia"), name)


def _add_inline_runs(p, text: str, size_pt: float = 12, base_bold: bool = False, base_italic: bool = False) -> None:
    """Converte negrito (**), italico (* ou _), codigo (`), sub/sup HTML e escapes em runs Word."""
    text = text.replace(r"\_", "_").replace(r"\*", "*")
    # Converte notacao LaTeX simples $H_1$ / $(H_{1})$ se houver
    text = re.sub(
        r"\$([^$\n]{1,40})\$",
        lambda m: re.sub(
            r"([A-Za-zÀ-ÿ0-9]+)\^\{?([A-Za-z0-9]+)\}?",
            r"\1<sup>\2</sup>",
            re.sub(r"([A-Za-zÀ-ÿ0-9]+)_\{?([A-Za-z0-9]+)\}?", r"\1<sub>\2</sub>", m.group(1).strip()),
        ),
        text,
    )

    token_re = re.compile(
        r"(\*\*[^*]+\*\*|__[^_]+__|<sub>[^<]+</sub>|<sup>[^<]+</sup>|`[^`]+`|\*[^*\n]+\*|(?<!\w)_[^_\n]+_(?!\w))"
    )
    pos = 0
    for m in token_re.finditer(text):
        if m.start() > pos:
            r = p.add_run(text[pos : m.start()])
            _set_run_font(r, size_pt=size_pt, bold=base_bold, italic=base_italic)
        tok = m.group(0)
        if tok.startswith("**") or tok.startswith("__"):
            inner = tok[2:-2]
            _add_inline_runs(p, inner, size_pt=size_pt, base_bold=True, base_italic=base_italic)
        elif tok.startswith("<sub>"):
            r = p.add_run(tok[5:-6])
            _set_run_font(r, size_pt=size_pt, bold=base_bold, italic=base_italic)
            r.font.subscript = True
        elif tok.startswith("<sup>"):
            r = p.add_run(tok[5:-6])
            _set_run_font(r, size_pt=size_pt, bold=base_bold, italic=base_italic)
            r.font.superscript = True
        elif tok.startswith("`"):
            r = p.add_run(tok[1:-1])
            _set_run_font(r, name="Courier New", size_pt=max(size_pt - 1, 9), bold=base_bold, italic=base_italic)
        elif tok.startswith("*") or tok.startswith("_"):
            inner = tok[1:-1]
            _add_inline_runs(p, inner, size_pt=size_pt, base_bold=base_bold, base_italic=True)
        pos = m.end()
    if pos < len(text):
        r = p.add_run(text[pos:])
        _set_run_font(r, size_pt=size_pt, bold=base_bold, italic=base_italic)


def _strip_md(s: str) -> str:
    s = s.replace(r"\_", "_").replace(r"\*", "*")
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"__(.+?)__", r"\1", s)
    s = re.sub(r"\*(.+?)\*", r"\1", s)
    s = re.sub(r"(?<!\w)_([^_\n]+?)_(?!\w)", r"\1", s)
    s = re.sub(r"`(.+?)`", r"\1", s)
    return s.strip()


def _configure_abnt_section(sect) -> None:
    sect.page_width = Cm(21.0)
    sect.page_height = Cm(29.7)
    sect.top_margin = Cm(3.0)
    sect.left_margin = Cm(3.0)
    sect.bottom_margin = Cm(2.0)
    sect.right_margin = Cm(2.0)
    sect.header_distance = Cm(2.0)
    sect.footer_distance = Cm(1.25)


def _set_section_pgnum_start(sect, start: int | None) -> None:
    sectPr = sect._sectPr
    for el in sectPr.findall(qn("w:pgNumType")):
        sectPr.remove(el)
    if start is not None:
        pg = OxmlElement("w:pgNumType")
        pg.set(qn("w:start"), str(start))
        sectPr.append(pg)


def _add_page_number_header(sect) -> None:
    sect.header.is_linked_to_previous = False
    p = sect.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = p.add_run()
    _set_run_font(r1, size_pt=10)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r1._r.append(fld_begin)
    r1._r.append(instr)
    r1._r.append(fld_sep)
    r1._r.append(fld_end)


def _clear_header(sect) -> None:
    sect.header.is_linked_to_previous = False
    p = sect.header.paragraphs[0]
    p.text = ""


def _set_cell_borders(cell, **edges) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge, spec in edges.items():
        el = tcBorders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tcBorders.append(el)
        if spec is None:
            el.set(qn("w:val"), "nil")
        else:
            val, sz, color = spec
            el.set(qn("w:val"), val)
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)


def _set_cell_margins(cell, top=80, bottom=80, left=100, right=100) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for k, v in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = tcMar.find(qn(f"w:{k}"))
        if el is None:
            el = OxmlElement(f"w:{k}")
            tcMar.append(el)
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")


def _mark_row_cant_split(row, is_header: bool = False) -> None:
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))
    if is_header and trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def _add_md_table(doc: Document, lines: list[str], is_quadro: bool) -> None:
    rows_raw: list[list[str]] = []
    aligns: list[str] = []
    for i, ln in enumerate(lines):
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if i == 1 and all(re.match(r"^:?-{2,}:?$", c) for c in cells):
            for c in cells:
                if c.startswith(":") and c.endswith(":"):
                    aligns.append("center")
                elif c.endswith(":"):
                    aligns.append("right")
                else:
                    aligns.append("left")
            continue
        rows_raw.append(cells)

    if not rows_raw:
        return
    ncols = max(len(r) for r in rows_raw)
    tbl = doc.add_table(rows=len(rows_raw), cols=ncols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    # Larguras proporcionais dentro dos 16.0 cm uteis da mancha A4 (21 - 3 - 2 = 16 cm)
    if is_quadro and ncols == 4:
        col_widths = [Cm(2.9), Cm(3.5), Cm(4.8), Cm(4.8)]
    elif not is_quadro and ncols == 4:
        col_widths = [Cm(7.6), Cm(2.8), Cm(2.8), Cm(2.8)]
    else:
        col_widths = [Cm(16.0 / ncols)] * ncols

    font_sz = 9.0 if is_quadro else 10.0

    for r_idx, r_data in enumerate(rows_raw):
        row = tbl.rows[r_idx]
        is_hdr = r_idx == 0
        _mark_row_cant_split(row, is_header=is_hdr)
        for c_idx in range(ncols):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            _set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            txt = r_data[c_idx] if c_idx < len(r_data) else ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.first_line_indent = Cm(0)

            al = aligns[c_idx] if c_idx < len(aligns) else "left"
            if al == "center":
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif al == "right":
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            _add_inline_runs(p, txt, size_pt=font_sz, base_bold=is_hdr)

            if is_quadro:
                _set_cell_borders(
                    cell,
                    top=("single", 4, "000000"),
                    bottom=("single", 4, "000000"),
                    left=("single", 4, "000000"),
                    right=("single", 4, "000000"),
                )
            else:
                # Padrao IBGE: sem bordas verticais; fios horizontais no topo, cabecalho e base
                top_b = ("single", 12, "000000") if is_hdr else None
                bot_b = (
                    ("single", 6, "000000")
                    if is_hdr
                    else (("single", 12, "000000") if r_idx == len(rows_raw) - 1 else ("single", 4, "CCCCCC"))
                )
                _set_cell_borders(cell, top=top_b, bottom=bot_b, left=None, right=None)


def export_markdown_to_docx(
    md_text: str,
    org: dict,
    out_docx: Path,
    mode: str | None = None,
    page_map: dict[str, int] | None = None,
    uni_dir: Path | None = None,
) -> None:
    from render import _clean_helper_notes, parse_frontmatter, split_top_h1_sections

    frontmatter, md_without_fm = parse_frontmatter(md_text)
    resolved_mode = (mode or str(frontmatter.get("mode", "")) or "monografia").lower()
    is_article = resolved_mode in ("artigo", "article", "simples")

    sections = split_top_h1_sections(md_without_fm)

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    rPr = normal.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rFonts.set(qn("w:cs"), "Arial")
    rPr.insert(0, rFonts)

    # Logo oficial UniFil
    logo_path: Path | None = None
    if uni_dir:
        cand = uni_dir / str(org.get("logo_file", "unifil-logo-trabalho-academico.png"))
        if cand.exists():
            logo_path = cand

    sect0 = doc.sections[0]
    _configure_abnt_section(sect0)
    _clear_header(sect0)

    has_capa = any(_strip_accents(t or "").upper().strip() == "CAPA" for t, _ in sections)
    pretext_section_started = False
    textual_section_started = False

    def _ensure_pretext_section():
        nonlocal pretext_section_started
        if not has_capa:
            if not pretext_section_started:
                _set_section_pgnum_start(sect0, 1)
                pretext_section_started = True
            return
        if not pretext_section_started:
            ns = doc.add_section(WD_SECTION.NEW_PAGE)
            _configure_abnt_section(ns)
            _clear_header(ns)
            _set_section_pgnum_start(ns, 1)
            pretext_section_started = True

    def _ensure_textual_section():
        nonlocal textual_section_started
        if textual_section_started:
            return
        if not has_capa and not pretext_section_started:
            _set_section_pgnum_start(sect0, 1)
            _add_page_number_header(sect0)
        else:
            ns = doc.add_section(WD_SECTION.NEW_PAGE)
            _configure_abnt_section(ns)
            _set_section_pgnum_start(ns, None)
            _add_page_number_header(ns)
        textual_section_started = True

    first_pretext = True

    for sec_title, sec_body in sections:
        if sec_title is None:
            continue
        norm = _strip_accents(sec_title).upper().strip()

        if norm == "CAPA":
            body_clean = _clean_helper_notes(sec_body)
            paras = [p.strip() for p in re.split(r"\n\s*\n", body_clean) if p.strip()]
            author = _strip_md(paras[0]) if paras else ""
            title = _strip_md(paras[1]) if len(paras) >= 2 else ""
            bottom = [p for p in paras[2:] if not (_strip_md(p).startswith("(") and _strip_md(p).endswith(")"))]

            if logo_path:
                lp = doc.add_paragraph()
                lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                lp.paragraph_format.space_after = Pt(12)
                lp.add_run().add_picture(str(logo_path), width=Cm(9.9))

            ap = doc.add_paragraph()
            ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ap.paragraph_format.space_after = Pt(120)
            _set_run_font(ap.add_run(author.upper()), size_pt=14)

            if ":" in title:
                mt, st = title.split(":", 1)
                tp = doc.add_paragraph()
                tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                tp.paragraph_format.line_spacing = 1.4
                tp.paragraph_format.space_after = Pt(4)
                _set_run_font(tp.add_run(mt.strip().upper() + ":"), size_pt=14, bold=True)
                sp = doc.add_paragraph()
                sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                sp.paragraph_format.line_spacing = 1.4
                sp.paragraph_format.space_after = Pt(170)
                _set_run_font(sp.add_run(st.strip().upper()), size_pt=14)
            else:
                tp = doc.add_paragraph()
                tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                tp.paragraph_format.space_after = Pt(170)
                _set_run_font(tp.add_run(title.upper()), size_pt=14, bold=True)

            bp = doc.add_paragraph()
            bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            bp.paragraph_format.line_spacing = 1.4
            loc_txt = "\n".join(_strip_md(x) for x in bottom) if bottom else f"{org.get('city', 'Londrina')}\n2026"
            _set_run_font(bp.add_run(loc_txt), size_pt=12)

        elif norm == "FOLHA DE ROSTO":
            _ensure_pretext_section()
            if not first_pretext:
                doc.add_page_break()
            first_pretext = False
            body_clean = _clean_helper_notes(sec_body)
            paras = [p.strip() for p in re.split(r"\n\s*\n", body_clean) if p.strip()]
            author = _strip_md(paras[0]) if paras else ""
            title = _strip_md(paras[1]) if len(paras) >= 2 else ""

            ap = doc.add_paragraph()
            ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ap.paragraph_format.space_after = Pt(110)
            _set_run_font(ap.add_run(author.upper()), size_pt=14)

            if ":" in title:
                mt, st = title.split(":", 1)
                tp = doc.add_paragraph()
                tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                tp.paragraph_format.space_after = Pt(4)
                _set_run_font(tp.add_run(mt.strip().upper() + ":"), size_pt=14, bold=True)
                sp = doc.add_paragraph()
                sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                sp.paragraph_format.space_after = Pt(48)
                _set_run_font(sp.add_run(st.strip().upper()), size_pt=14)
            else:
                tp = doc.add_paragraph()
                tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                tp.paragraph_format.space_after = Pt(48)
                _set_run_font(tp.add_run(title.upper()), size_pt=14, bold=True)

            bottom_lines = []
            for p_blk in paras[2:]:
                cl = _strip_md(p_blk)
                if cl.startswith("(") and cl.endswith(")"):
                    continue
                if p_blk.startswith(">"):
                    inner_lines = [re.sub(r"^>\s?", "", ln).strip() for ln in p_blk.splitlines() if ln.strip() != ">"]
                    for ln in [x for x in inner_lines if x]:
                        np = doc.add_paragraph()
                        np.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        np.paragraph_format.left_indent = Cm(8.0)
                        np.paragraph_format.line_spacing = 1.0
                        np.paragraph_format.space_after = Pt(10)
                        _add_inline_runs(np, ln, size_pt=12)
                else:
                    bottom_lines.append(cl)

            bp = doc.add_paragraph()
            bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            bp.paragraph_format.space_before = Pt(90)
            bp.paragraph_format.line_spacing = 1.4
            _set_run_font(bp.add_run("\n".join(bottom_lines) or f"{org.get('city', 'Londrina')}\n2026"), size_pt=12)

        elif norm == "BANCA EXAMINADORA":
            _ensure_pretext_section()
            doc.add_page_break()
            paras = [p.strip() for p in re.split(r"\n\s*\n", sec_body) if p.strip()]
            author = _strip_md(paras[0]) if paras else ""
            title = _strip_md(paras[1]) if len(paras) >= 2 else ""

            ap = doc.add_paragraph()
            ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ap.paragraph_format.space_after = Pt(24)
            _set_run_font(ap.add_run(author.upper()), size_pt=14)

            if ":" in title:
                mt, st = title.split(":", 1)
                tp = doc.add_paragraph()
                tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                tp.paragraph_format.space_after = Pt(2)
                _set_run_font(tp.add_run(mt.strip().upper() + ":"), size_pt=14, bold=True)
                sp = doc.add_paragraph()
                sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                sp.paragraph_format.space_after = Pt(24)
                _set_run_font(sp.add_run(st.strip().upper()), size_pt=14)

            i = 2
            while i < len(paras):
                p_blk = paras[i]
                if p_blk.startswith(">"):
                    inner_lines = [re.sub(r"^>\s?", "", ln).strip() for ln in p_blk.splitlines() if ln.strip() != ">"]
                    for ln in [x for x in inner_lines if x]:
                        np = doc.add_paragraph()
                        np.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        np.paragraph_format.left_indent = Cm(8.0)
                        np.paragraph_format.line_spacing = 1.0
                        np.paragraph_format.space_after = Pt(8)
                        _add_inline_runs(np, ln, size_pt=12)
                elif "BANCA EXAMINADORA" in p_blk.upper():
                    bh = doc.add_paragraph()
                    bh.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    bh.paragraph_format.space_before = Pt(18)
                    bh.paragraph_format.space_after = Pt(24)
                    _set_run_font(bh.add_run("BANCA EXAMINADORA"), size_pt=12, bold=True)
                elif re.match(r"^(?:_{4,}|-{3,})", _strip_md(p_blk)):
                    lines = [ln.strip() for ln in p_blk.splitlines() if ln.strip()]
                    rest = lines[1:] if len(lines) > 1 else []
                    mp = doc.add_paragraph()
                    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    mp.paragraph_format.left_indent = Cm(6.5)
                    mp.paragraph_format.line_spacing = 1.15
                    mp.paragraph_format.space_after = Pt(18)
                    _set_run_font(mp.add_run("____________________________________\n" + "\n".join(_strip_md(x) for x in rest)), size_pt=12)
                else:
                    dp = doc.add_paragraph()
                    dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    dp.paragraph_format.space_before = Pt(24)
                    _set_run_font(dp.add_run(_strip_md(p_blk)), size_pt=12)
                i += 1

        elif norm in ("RESUMO", "ABSTRACT"):
            _ensure_pretext_section()
            doc.add_page_break()
            body_clean = _clean_helper_notes(sec_body)
            paras = [p.strip() for p in re.split(r"\n\s*\n", body_clean) if p.strip()]
            for idx, p_blk in enumerate(paras):
                if idx == 0 and re.match(r"^[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,},\s", p_blk):
                    rp = doc.add_paragraph()
                    rp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    rp.paragraph_format.line_spacing = 1.0
                    rp.paragraph_format.space_after = Pt(14)
                    _add_inline_runs(rp, p_blk, size_pt=12)

                    hp = doc.add_paragraph()
                    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    hp.paragraph_format.space_after = Pt(14)
                    _set_run_font(hp.add_run(sec_title.upper()), size_pt=12, bold=True)
                else:
                    if idx == 0:
                        hp = doc.add_paragraph()
                        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        hp.paragraph_format.space_after = Pt(14)
                        _set_run_font(hp.add_run(sec_title.upper()), size_pt=12, bold=True)
                    bp = doc.add_paragraph()
                    bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    bp.paragraph_format.line_spacing = 1.0
                    bp.paragraph_format.space_after = Pt(12)
                    _add_inline_runs(bp, p_blk, size_pt=12)

        elif norm in ("LISTA DE ABREVIATURAS E SIGLAS", "LISTA DE SIGLAS"):
            _ensure_pretext_section()
            doc.add_page_break()
            hp = doc.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hp.paragraph_format.space_after = Pt(14)
            _set_run_font(hp.add_run(sec_title.upper()), size_pt=12, bold=True)

            sig_rows = []
            for ln in sec_body.splitlines():
                s = ln.strip()
                m = re.match(r"^-\s+\*\*(.+?)\*\*\s*[-–]\s*(.+)$", s) or re.match(r"^([A-Za-z0-9/\-]+)\s{2,}(.+)$", s)
                if m:
                    sig_rows.append((m.group(1).strip(), m.group(2).strip()))
            if sig_rows:
                tbl = doc.add_table(rows=len(sig_rows), cols=2)
                tbl.autofit = False
                for r_i, (k, v) in enumerate(sig_rows):
                    row = tbl.rows[r_i]
                    c0, c1 = row.cells[0], row.cells[1]
                    c0.width, c1.width = Cm(3.0), Cm(13.0)
                    _set_cell_borders(c0, top=None, bottom=None, left=None, right=None)
                    _set_cell_borders(c1, top=None, bottom=None, left=None, right=None)
                    p0, p1 = c0.paragraphs[0], c1.paragraphs[0]
                    p0.paragraph_format.space_after = Pt(2)
                    p1.paragraph_format.space_after = Pt(2)
                    _set_run_font(p0.add_run(k), size_pt=12, bold=True)
                    _set_run_font(p1.add_run(v), size_pt=12)

        elif norm.startswith("LISTA DE ") or norm == "SUMARIO":
            _ensure_pretext_section()
            doc.add_page_break()
            hp = doc.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hp.paragraph_format.space_after = Pt(14)
            _set_run_font(hp.add_run(sec_title.upper()), size_pt=12, bold=True)

            for ln in sec_body.splitlines():
                s = re.sub(r"^-\s+", "", ln.strip()).rstrip("\\ ")
                if not s:
                    continue
                m = re.match(r"^(.*?)\s*(?:\.{2,}\s*)?(?:p\.\s*[Xx0-9]+|\b00\b)$", s)
                lbl = m.group(1).strip() if m else s
                pg_str = "X"
                if m and not re.search(r"p\.\s*[Xx]|00$", s):
                    pm = re.search(r"(\d+)$", s)
                    if pm:
                        pg_str = pm.group(1)
                if page_map and _norm_key(lbl) in page_map:
                    pg_str = str(page_map[_norm_key(lbl)])

                lvl = 1
                num_m = re.match(r"^(\d+(?:\.\d+)*)\s+(.+)$", _strip_md(lbl))
                if num_m:
                    lvl = min(len(num_m.group(1).split(".")), 5)

                tp = doc.add_paragraph()
                tp.paragraph_format.line_spacing = 1.35
                tp.paragraph_format.space_before = Pt(4 if lvl == 1 else 0)
                tp.paragraph_format.space_after = Pt(2)
                tp.paragraph_format.tab_stops.add_tab_stop(
                    Cm(16.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
                )
                is_b = lvl in (1, 3)
                is_i = lvl == 5
                txt_disp = _strip_md(lbl).upper() if lvl in (1, 2) else _strip_md(lbl)
                _set_run_font(tp.add_run(f"{txt_disp}\t{pg_str}"), size_pt=12, bold=is_b, italic=is_i)

        elif norm == "REFERENCIAS":
            _ensure_textual_section()
            if not is_article:
                doc.add_page_break()
            hp = doc.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hp.paragraph_format.space_after = Pt(18)
            _set_run_font(hp.add_run(sec_title.upper()), size_pt=12, bold=True)

            for para in re.split(r"\n\s*\n", sec_body.strip()):
                p_txt = para.strip()
                if not p_txt:
                    continue
                rp = doc.add_paragraph()
                rp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                rp.paragraph_format.line_spacing = 1.0
                rp.paragraph_format.space_after = Pt(12)
                _add_inline_runs(rp, p_txt, size_pt=12)

        else:
            # Secao textual principal (# 1 INTRODUCAO, # 2 ..., etc.)
            first_textual = not textual_section_started
            _ensure_textual_section()
            if not first_textual and not is_article:
                doc.add_page_break()

            hp = doc.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            hp.paragraph_format.line_spacing = 1.5
            hp.paragraph_format.space_before = Pt(12 if is_article and not first_textual else 0)
            hp.paragraph_format.space_after = Pt(18)
            hp.paragraph_format.keep_with_next = True
            _set_run_font(hp.add_run(sec_title.upper()), size_pt=12, bold=True)

            lines = sec_body.splitlines()
            idx = 0
            pending_quadro = False
            while idx < len(lines):
                ln = lines[idx]
                s = ln.strip()
                if not s:
                    idx += 1
                    continue

                # Subsecoes (##..#####)
                hm = re.match(r"^(#{2,5})\s+(.+)$", s)
                if hm:
                    level = len(hm.group(1))
                    htext = _strip_md(hm.group(2))
                    sh = doc.add_paragraph()
                    sh.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    sh.paragraph_format.line_spacing = 1.5
                    sh.paragraph_format.space_before = Pt(14)
                    sh.paragraph_format.space_after = Pt(12)
                    sh.paragraph_format.keep_with_next = True
                    if level == 2:
                        _set_run_font(sh.add_run(htext.upper()), size_pt=12, bold=False)
                    elif level == 3:
                        _set_run_font(sh.add_run(htext), size_pt=12, bold=True)
                    elif level == 4:
                        _set_run_font(sh.add_run(htext), size_pt=12, bold=False)
                    else:
                        _set_run_font(sh.add_run(htext), size_pt=12, italic=True)
                    idx += 1
                    continue

                # Tabela Markdown
                if s.startswith("|") and idx + 1 < len(lines) and lines[idx + 1].strip().startswith("|"):
                    tbl_lines = []
                    while idx < len(lines) and lines[idx].strip().startswith("|"):
                        tbl_lines.append(lines[idx])
                        idx += 1
                    _add_md_table(doc, tbl_lines, is_quadro=pending_quadro)
                    pending_quadro = False
                    continue

                # Blockquote (citacao longa ou destaque)
                if s.startswith(">"):
                    q_lines = []
                    while idx < len(lines) and lines[idx].strip().startswith(">"):
                        q_lines.append(re.sub(r"^>\s?", "", lines[idx].strip()))
                        idx += 1
                    q_txt = " ".join(x for x in q_lines if x).strip()
                    qp = doc.add_paragraph()
                    qp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    qp.paragraph_format.left_indent = Cm(4.0)
                    qp.paragraph_format.line_spacing = 1.0
                    qp.paragraph_format.space_before = Pt(8)
                    qp.paragraph_format.space_after = Pt(8)
                    _add_inline_runs(qp, q_txt, size_pt=10)
                    continue

                # Lista com marcador (- )
                if s.startswith("- "):
                    bp = doc.add_paragraph()
                    bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    bp.paragraph_format.left_indent = Cm(2.2)
                    bp.paragraph_format.first_line_indent = Cm(-0.5)
                    bp.paragraph_format.line_spacing = 1.5
                    bp.paragraph_format.space_after = Pt(3)
                    _add_inline_runs(bp, "• " + s[2:].strip(), size_pt=12)
                    idx += 1
                    continue

                # Legenda Quadro/Tabela/Figura
                if re.match(r"^(?:\*\*)?(Quadro|Tabela|Figura|Gráfico)\s+\d+", s):
                    pending_quadro = "quadro" in s.lower()[:20]
                    cp = doc.add_paragraph()
                    cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    cp.paragraph_format.line_spacing = 1.25
                    cp.paragraph_format.space_before = Pt(12)
                    cp.paragraph_format.space_after = Pt(4)
                    cp.paragraph_format.keep_with_next = True
                    _add_inline_runs(cp, s, size_pt=12)
                    idx += 1
                    continue

                # Fonte:
                if re.match(r"^(?:\*\*)?Fonte:", s):
                    fp = doc.add_paragraph()
                    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    fp.paragraph_format.line_spacing = 1.15
                    fp.paragraph_format.space_before = Pt(4)
                    fp.paragraph_format.space_after = Pt(12)
                    _add_inline_runs(fp, s, size_pt=10)
                    idx += 1
                    continue

                # Alinea a) ...
                if re.match(r"^[a-z]\)\s+", s):
                    ap = doc.add_paragraph()
                    ap.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    ap.paragraph_format.left_indent = Cm(2.2)
                    ap.paragraph_format.first_line_indent = Cm(-0.7)
                    ap.paragraph_format.line_spacing = 1.5
                    ap.paragraph_format.space_after = Pt(0)
                    _add_inline_runs(ap, s, size_pt=12)
                    idx += 1
                    continue

                # Paragrafo comum (agrega linhas continuas ate linha em branco)
                p_lines = [s]
                idx += 1
                while idx < len(lines):
                    nxt = lines[idx].strip()
                    if not nxt or nxt.startswith(("#", "|", ">", "- ")) or re.match(r"^[a-z]\)\s+", nxt):
                        break
                    p_lines.append(nxt)
                    idx += 1
                pp = doc.add_paragraph()
                pp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                pp.paragraph_format.first_line_indent = Cm(1.5)
                pp.paragraph_format.line_spacing = 1.5
                pp.paragraph_format.space_after = Pt(0)
                _add_inline_runs(pp, " ".join(p_lines), size_pt=12)

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_docx))
