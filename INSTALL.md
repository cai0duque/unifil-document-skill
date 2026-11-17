# Instalação da Skill - Documentos Acadêmicos UniFil (`unifil-document-skill`)

Guia de instalação e execução em **Linux**, **macOS** e **Windows**.

---

## 1. Instalar a Skill no seu Agente de IA

### No Antigravity / Gemini CLI

Copie ou crie um link simbólico desta pasta para `~/.gemini/config/skills/unifil-document-skill`:

```bash
# Linux / macOS (link simbolico mantendo o diretorio atual como fonte)
ln -s /home/duque/Documentos/TCC/unifil-document-skill ~/.gemini/config/skills/unifil-document-skill

# Ou copiando diretamente para a pasta global de skills:
cp -r /home/duque/Documentos/TCC/unifil-document-skill ~/.gemini/config/skills/unifil-document-skill
```

No Windows (PowerShell):
```powershell
Copy-Item -Recurse .\unifil-document-skill "$HOME\.gemini\config\skills\unifil-document-skill"
```

---

## 2. Ambiente Python

### Opção A - Usando `uv` (Recomendado)

Com o [uv](https://docs.astral.sh/uv/) instalado, o `render.py` resolve automaticamente as dependências (`playwright`, `pyyaml`, `markdown-it-py`, `pypdf`) declaradas no cabeçalho PEP 723:

```bash
uv run render.py --list-templates
```

### Opção B - Usando `.venv` tradicional

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

---

## 3. Navegador para Renderização do PDF

O `render.py` detecta automaticamente **Google Chrome**, **Chromium** ou **Microsoft Edge** instalados no sistema. Caso esteja em um ambiente sem navegador, instale o Chromium do Playwright:

```bash
uv run playwright install chromium
```

---

## 4. Teste Rápido

```bash
cd /home/duque/Documentos/TCC/unifil-document-skill
uv run render.py ../tcc_revisado.md --check
uv run render.py ../tcc_revisado.md --out ../tcc_revisado.pdf
```
