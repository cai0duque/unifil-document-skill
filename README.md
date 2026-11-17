# UniFil Academic Document Skill (`unifil-document-skill`)

Gerador de trabalhos acadêmicos e científicos da **UniFil (Centro Universitário Filadélfia)** em PDF a partir de arquivos Markdown (`.md`), seguindo o layout oficial da instituição (fonte **Arial**, logo oficial na capa, margens ABNT) e as normas **ABNT** (NBR 6023, NBR 6024, NBR 6027, NBR 6028, NBR 10520 e NBR 14724) e **IBGE**.

Inspirado na arquitetura modular de `cabr-document-skill`, adaptado para o ecossistema universitário da UniFil e compatível com **Antigravity**, **Claude Code**, **Cursor** e execução direta em terminal (**Linux**, **macOS** e **Windows**).

---

## Templates Inclusos (`universidades/UniFil/templates/`)

| Tipo de Documento | Arquivo | Modo |
|---|---|---|
| **Projeto de Pesquisa de TCC (TCC I)** | `projeto-pesquisa-tcc.md` | `monografia` |
| **Monografia / TCC II Completo** | `tcc-monografia.md` | `monografia` |
| **Artigo Científico (NBR 6022)** | `artigo-cientifico.md` | `artigo` |
| **Trabalho Acadêmico / Resenha** | `trabalho-academico-geral.md` | `monografia` |
| **Relatório Técnico / Estágio / Prática** | `relatorio-estagio-pratica.md` | `monografia` |
| **Modelo Oficial UniFil (.docx / .md)** | `modelo-de-trabalho-unifil-graduacao.docx` | Referência |

---

## Principais Recursos

- **Layout Oficial UniFil em Arial:** margens A4 de `3 cm` (superior/esquerda) e `2 cm` (inferior/direita), recuo de parágrafo de `1,5 cm`, entrelinhas `1,5`, citações longas em `10 pt` com recuo de `4 cm`, e fonte **Arial** (com **Liberation Sans** embarcada offline para fidelidade no Linux).
- **Capa e Pré-textuais ABNT:** logo oficial `unifil-logo-trabalho-academico.png` centralizado na Capa, Folha de Rosto com bloco de natureza do trabalho alinhado à direita, Banca Examinadora, Resumo/Abstract em parágrafo único e Lista de Siglas tabulada.
- **Paginação ABNT Real + Sumário Automático:** a contagem inicia na Folha de Rosto (Capa não contada), permanece oculta nos pré-textuais e passa a ser exibida no canto superior direito a partir da Introdução. Os números de página no Sumário e nas Listas são preenchidos automaticamente em duas passagens via `pypdf`.
- **Tabelas IBGE e Quadros ABNT:** tabelas sem bordas verticais no padrão IBGE e detecção automática de quadros fechados quando precedidos de `Quadro N - ...`.
- **Diagramas Mermaid e Blocos de Código Offline:** suporte nativo a ` ```mermaid ` via `assets/mermaid.min.js`.
- **Auditor ABNT Integrado (`--check`):** verifica limite de palavras do Resumo, pontuação das Palavras-chave, siglas órfãs, referências não citadas, travessões e emojis.

---

## Uso Rápido

```bash
# Validar normas ABNT e redacao do documento
uv run render.py meu-tcc.md --check

# Gerar PDF oficial UniFil (com Sumario e paginacao automatica)
uv run render.py meu-tcc.md

# Gerar em modo artigo (sem quebra de pagina por secao primaria)
uv run render.py meu-artigo.md --mode artigo --out artigo.pdf

# Listar templates e instituicoes
uv run render.py --list-templates
uv run render.py --list-unis
```

Consulte [INSTALL.md](./INSTALL.md) para instalar esta skill globalmente no Antigravity ou em outros agentes, e [SKILL.md](./SKILL.md) para as instruções completas de uso pelo agente.
