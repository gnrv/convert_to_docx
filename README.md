# convert_to_docx

Convert Markdown files to `.docx` with Mermaid diagrams rendered as PNG images before conversion.

The script in this folder:
- `convert_to_docx.py`

## What it does

1. Finds `pandoc` and `mmdc` (Mermaid CLI).
2. Scans Markdown for fenced Mermaid blocks:

   ```mermaid
   graph TD
     A --> B
   ```

3. Renders each Mermaid block to a PNG.
4. Replaces Mermaid blocks with Markdown image references.
5. Runs Pandoc to produce a `.docx`.

## Requirements

- Python 3.10+
- Pandoc CLI
- Mermaid CLI (`mmdc` from `@mermaid-js/mermaid-cli`)
- Node.js (required by `mmdc`)

Python package requirements are listed in `requirements.txt`.

## Install

### Quick start: everything in a project venv (no system installs)

Pandoc comes from `pypandoc_binary`, and Node.js + Mermaid CLI are installed into the venv via `nodeenv`:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
nodeenv -p                                  # Node.js inside .venv
npm install -g @mermaid-js/mermaid-cli      # mmdc inside .venv
```

After that, `source .venv/bin/activate` is all you need. To install the tools system-wide instead, follow the steps below.

### 1) Python dependencies

```bash
pip install -r requirements.txt
```

Note: there are no third-party Python packages required at this time.

### 2) Install Pandoc

**Windows 11:**
```bash
# Using conda
conda install -c conda-forge pandoc -y
# OR using winget
winget install pandoc
```

**Ubuntu 22.04 / 24.04:**
```bash
sudo apt update
sudo apt install pandoc
```
*(If you need the absolute latest, download the `.deb` from [Pandoc's releases page](https://github.com/jgm/pandoc/releases).)*

Alternative installers: https://pandoc.org/installing.html

### 3) Install Mermaid CLI (`mmdc`)

Mermaid CLI uses Headless Chrome to render images, which requires Node.js/npm.

**Windows 11:**
1. Install Node.js if you don't have it:
   ```bash
   winget install OpenJS.NodeJS.LTS
   ```
2. Install Mermaid CLI globally:
   ```bash
   npm install -g @mermaid-js/mermaid-cli
   ```
*If your shell cannot find `mmdc`, ensure your npm global bin path (often `C:\Users\<User>\AppData\Roaming\npm`) is in your `PATH`.*

**Ubuntu 22.04 / 24.04:**
1. Install Node.js, npm, and required Chromium OS dependencies:
   ```bash
   sudo apt update
   sudo apt install -y nodejs npm libnss3 libatk1.0-0 libatk-bridge2.0-0 libcups2 libdrm2 libxkbcommon0 libxcomposite1 libxdamage1 libxrandr2 libgbm1 libasound2
   ```
2. Install Mermaid CLI globally:
   ```bash
   sudo npm install -g @mermaid-js/mermaid-cli
   ```

## Usage

Basic conversion:

```bash
python convert_to_docx.py Architecture.md
```

Custom output file:

```bash
python convert_to_docx.py Architecture.md -o Architecture_styled.docx
```

Use a Word style template (`reference.docx`):

```bash
python convert_to_docx.py Architecture.md --template reference.docx
```

Any existing Word document can serve as the template; its styles, headers, footers and page setup are used, its content is ignored. Pandoc styles the template lacks (Body Text, Compact, Table, …) are added, based on the template's Normal and Table Grid styles.

### Document layout options

- **Table of contents placement:** by default a TOC is inserted before the body. Put a line containing only `[TOC]` to place it there instead (e.g. after a cover page). Word fills it in when it asks to update fields on open.
- **Custom cover page:** set `title-block: false` in the YAML front matter to suppress pandoc's title block (title, subtitle, author, date). The `title` is still written to the document properties, so template headers bound to the title show it. Write the cover in the body using template styles, e.g. `::: {custom-style="DocTitle"}`.
- **Table style per table:** add an attribute-only caption after the table, e.g. `: {custom-style="TableNormal"}` for a borderless table.
- **Images** are resolved relative to the Markdown file's folder.

## Output

- Default output: same name as input with `.docx` extension.
- Mermaid diagrams are rendered to temporary PNGs during conversion.

## Troubleshooting

- Error: missing `pandoc`
  - Install with conda command above, or add Pandoc install directory to `PATH`.

- Error: missing `mmdc`
  - Install Mermaid CLI globally with npm.

- `mmdc` fails with `"node" is not recognized`
  - Ensure Node.js is installed and available in `PATH`.

- Diagrams fail to render but docx is still produced
  - The script keeps original Mermaid code blocks if rendering fails.
