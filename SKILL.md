---
name: unifil-document-skill
description: >-
  Gera documentos acadêmicos e científicos da UniFil (Centro Universitário Filadélfia) em PDF seguindo as normas da ABNT (NBR 6023, 6024, 6027, 6028, 10520, 14724) e o layout oficial da instituição em fonte Arial, com logo oficial na capa, margens ABNT (3 cm sup/esq, 2 cm inf/dir), paginação contada da Folha de Rosto e exibida a partir da Introdução, Sumário automático, tabelas padrão IBGE, quadros fechados, citações longas com recuo de 4 cm e diagramas Mermaid offline. Use sempre que o usuário pedir para criar, estruturar, validar ou exportar qualquer trabalho acadêmico, projeto de pesquisa, TCC, artigo científico ou relatório da UniFil.
---

# UniFil Academic Document Skill (`unifil-document-skill`)

Gera **documentos acadêmicos da UniFil em PDF** a partir de Markdown (`.md`), aplicando o layout do modelo oficial [modelo-de-trabalho-unifil-graduacao.docx](./universidades/UniFil/templates/modelo-de-trabalho-unifil-graduacao.docx), fonte **Arial** (com fallback métrica-idêntico embarcado offline), logo oficial da UniFil na capa e conformidade com as normas da **ABNT** e do **IBGE**.

## Templates Disponíveis (`universidades/UniFil/templates/`)

| Tipo de Documento | Template Base | Modo (`mode`) |
|---|---|---|
| **Projeto de Pesquisa de TCC (TCC I)** | `templates/projeto-pesquisa-tcc.md` | `monografia` |
| **Monografia / TCC II Completo** | `templates/tcc-monografia.md` | `monografia` |
| **Artigo Científico (NBR 6022)** | `templates/artigo-cientifico.md` | `artigo` |
| **Trabalho Acadêmico / Resenha** | `templates/trabalho-academico-geral.md` | `monografia` |
| **Relatório Técnico / Estágio / Prática** | `templates/relatorio-estagio-pratica.md` | `monografia` |
| **Referência Oficial da Instituição** | `templates/modelo-de-trabalho-unifil-graduacao.docx` | - |

Para listar os templates via terminal:
```bash
uv run render.py --list-templates
```

## Como executar o gerador de PDF e o validador ABNT

```bash
# 1. Validar regras ABNT e de redacao (Resumo, Palavras-chave, Siglas, Referencias, travessoes/emojis)
uv run render.py meu-trabalho.md --check

# 2. Gerar PDF oficial UniFil (preenche paginas do Sumario e Listas automaticamente)
uv run render.py meu-trabalho.md

# 3. Especificar arquivo de saida ou modo artigo (fluxo continuo sem quebra por H1)
uv run render.py meu-artigo.md --mode artigo --out ./entregas/artigo.pdf
```

> **Detecção automática de navegador:** Se o Chromium do Playwright não estiver instalado, o `render.py` localiza automaticamente o Google Chrome, Chromium ou Microsoft Edge instalado no Linux, macOS ou Windows.

## Especificações de Layout Implementadas (Modelo Oficial UniFil + ABNT)

- **Papel e Margens (NBR 14724):** A4 (`210 mm x 297 mm`), margem superior e esquerda de `3 cm` (`30 mm`), inferior e direita de `2 cm` (`20 mm`).
- **Tipografia obrigatória:** **Arial** (`12 pt` no corpo do texto e títulos; `14 pt` na capa para autor e título; `10 pt` em citações diretas longas, tabelas, quadros e fontes). O pacote embarca `Liberation Sans` e `Liberation Mono` em `universidades/UniFil/fontes/` para garantir renderização offline com métricas idênticas à Arial no Linux.
- **Capa (`# CAPA`):** exibe centralizado no topo o logo oficial (`unifil-logo-trabalho-academico.png`, largura `9,9 cm`), autor em caixa alta (`14 pt`), título em negrito e subtítulo (`14 pt`), e local/ano no rodapé (`12 pt`). Não recebe número de página.
- **Folha de Rosto (`# FOLHA DE ROSTO`) e Banca (`# BANCA EXAMINADORA`):** sem logo, com bloco de natureza do trabalho em espaçamento simples recuado `50%` à direita.
- **Numeração de páginas (NBR 14724):** contada a partir da Folha de Rosto (Capa = 0, Folha de Rosto = 1), oculta nos elementos pré-textuais e exibida no canto superior direito (`10 pt`) a partir de `# 1 INTRODUÇÃO` até as `# REFERÊNCIAS` / apêndices / anexos.
- **Sumário e Listas automáticos:** marcadores `p. X` ou `00` em `# SUMÁRIO`, `# LISTA DE QUADROS`, `# LISTA DE TABELAS` e `# LISTA DE ILUSTRAÇÕES` são substituídos automaticamente em duas passagens pelos números reais de página do PDF.
- **Seções Progressivas (NBR 6024):**
  - `# 1 TÍTULO` -> caixa alta, negrito (`12 pt`), alinhado à esquerda (inicia nova página no modo `monografia`);
  - `## 2.1 TÍTULO` -> caixa alta, sem negrito (`12 pt`);
  - `### 2.1.1 Título` -> primeiras letras em maiúsculo, sem negrito (`12 pt`);
  - `#### 2.1.1.1 Título` -> somente a 1ª letra da 1ª palavra em maiúsculo (`12 pt`);
  - `##### 2.1.1.1.1 Título` -> itálico (`12 pt`).
- **Parágrafos e Citações (NBR 10520):** corpo justificado, espaçamento `1,5` entre linhas e recuo de primeira linha de `1,5 cm`. Citações diretas com mais de 3 linhas em bloco (`> ...`) recebem recuo esquerdo de `4 cm`, fonte `10 pt`, espaçamento simples e sem aspas.
- **Tabelas (IBGE) vs. Quadros (ABNT):** tabelas padrão são renderizadas sem bordas verticais (apenas fios horizontais no topo, cabeçalho e base). Tabelas precedidas por legenda `Quadro N - ...` (ou envolvidas em `<!-- abnt-quadro --> ... <!-- /abnt-quadro -->`) recebem grade fechada nos quatro lados.
- **Referências (NBR 6023):** alinhamento exclusivamente à esquerda, espaçamento simples e separação de uma linha simples entre entradas.

## Regras de Redação Acadêmica e ABNT

1. **Sem marcas de geração automática por IA:** evitar travessões (`U+2014` / `U+2013`) como conectores na prosa (preferir vírgulas ou parênteses), evitar emojis, setas Unicode e floreios retóricos.
2. **Consistência pré-textual:** título em português e em língua estrangeira idênticos entre Capa, Folha de Rosto, Banca, Resumo e Abstract.
3. **Resumo e Palavras-chave (NBR 6028 / Projeto UniFil):** parágrafo único, espaçamento simples, até 250 palavras para Projetos de Pesquisa (150 a 500 palavras para Monografia), seguido de até 5 palavras-chave separadas por ponto e vírgula (`;`) e finalizadas por ponto final (`.`).
4. **Siglas e Referências 1-para-1:** toda sigla listada em `LISTA DE ABREVIATURAS E SIGLAS` deve constar no corpo textual, e toda obra em `REFERÊNCIAS` deve estar citada no texto (e vice-versa).

Sempre valide o Markdown antes de gerar a entrega final:
```bash
uv run render.py arquivo.md --check
```
