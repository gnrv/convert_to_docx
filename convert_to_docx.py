#!/usr/bin/env python3
"""Convert Markdown with Mermaid diagrams to .docx.

Requirements:
    - pandoc (https://pandoc.org/installing.html)
    - mmdc  (npm install -g @mermaid-js/mermaid-cli)

Usage:
    python convert_to_docx.py Architecture.md
    python convert_to_docx.py Architecture.md -o output.docx
    python convert_to_docx.py Architecture.md --template sigma-template.docx
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

# Home directory for building search paths
_HOME = Path.home()

# Common Windows locations to search when tools aren't on PATH
_SEARCH_PATHS = {
    "pandoc": [
        _HOME / "miniconda3" / "Library" / "bin" / "pandoc.exe",
        _HOME / "miniconda3" / "Scripts" / "pandoc.exe",
        _HOME / "anaconda3" / "Library" / "bin" / "pandoc.exe",
        _HOME / "anaconda3" / "Scripts" / "pandoc.exe",
        Path(r"C:\Program Files\Pandoc\pandoc.exe"),
        _HOME / "AppData" / "Local" / "Pandoc" / "pandoc.exe",
        _HOME / "scoop" / "shims" / "pandoc.exe",
        Path(r"C:\ProgramData\chocolatey\bin\pandoc.exe"),
    ],
    "mmdc": [
        _HOME / "AppData" / "Roaming" / "npm" / "mmdc.cmd",
        _HOME / "AppData" / "Roaming" / "npm" / "mmdc",
        Path(r"C:\Program Files\nodejs\mmdc.cmd"),
        _HOME / "scoop" / "shims" / "mmdc.exe",
    ],
}


def _find_tool(name: str) -> str:
    """Find a tool by checking PATH first, then common Windows install locations."""
    found = shutil.which(name)
    if found:
        return found
    for candidate in _SEARCH_PATHS.get(name, []):
        if candidate.exists():
            return str(candidate)
    if name == "pandoc":
        # Bundled with `pip install pypandoc_binary`
        try:
            import pypandoc
            return pypandoc.get_pandoc_path()
        except (ImportError, OSError):
            pass
    return ""


def _ensure_node_on_path():
    """Ensure Node.js is on PATH so mmdc's .cmd wrapper can find it."""
    node_dirs = [
        r"C:\Program Files\nodejs",
        str(_HOME / "AppData" / "Roaming" / "npm"),
        str(_HOME / "scoop" / "shims"),
    ]
    current = os.environ.get("PATH", "")
    additions = [d for d in node_dirs if d not in current and Path(d).exists()]
    if additions:
        os.environ["PATH"] = ";".join(additions) + ";" + current


def check_dependencies() -> dict[str, str]:
    """Locate pandoc and mmdc, return dict of {tool: path}. Exit on failure."""
    tools = {}
    missing = []
    for cmd in ("pandoc", "mmdc"):
        path = _find_tool(cmd)
        if path:
            tools[cmd] = path
            print(f"  Found {cmd}: {path}")
        else:
            missing.append(cmd)
    if missing:
        print(f"Error: missing required commands: {', '.join(missing)}", file=sys.stderr)
        if "pandoc" in missing:
            print("  Install pandoc: conda install -c conda-forge pandoc", file=sys.stderr)
        if "mmdc" in missing:
            print("  Install mmdc:   npm install -g @mermaid-js/mermaid-cli", file=sys.stderr)
        sys.exit(1)
    return tools


def render_mermaid(code: str, out_path: Path, mmdc_path: str, theme: str = "default") -> bool:
    """Render a Mermaid code block to a PNG image via mmdc."""
    with tempfile.NamedTemporaryFile(suffix=".mmd", mode="w", delete=False, encoding="utf-8") as f:
        f.write(code)
        mmd_path = Path(f.name)

    # mmdc config for high-res output
    config = {"theme": theme, "themeVariables": {"fontSize": "14px"}}
    config_path = mmd_path.with_suffix(".json")
    config_path.write_text(json.dumps(config), encoding="utf-8")

    try:
        result = subprocess.run(
            [
                mmdc_path,
                "-i", str(mmd_path),
                "-o", str(out_path),
                "-b", "white",
                "-s", "2",           # 2× scale for crisp diagrams
                "-c", str(config_path),
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            print(f"  mmdc warning: {result.stderr.strip()}", file=sys.stderr)
            return False
        return out_path.exists()
    except subprocess.TimeoutExpired:
        print("  mmdc timed out", file=sys.stderr)
        return False
    finally:
        mmd_path.unlink(missing_ok=True)
        config_path.unlink(missing_ok=True)


def process_markdown(md_text: str, work_dir: Path, tools: dict[str, str]) -> str:
    """Replace ```mermaid blocks with rendered PNG image references."""
    pattern = re.compile(r"```mermaid\s*\n(.*?)```", re.DOTALL)
    counter = 0

    def replacer(match):
        nonlocal counter
        counter += 1
        code = match.group(1).strip()
        img_name = f"mermaid_{counter:02d}.png"
        img_path = work_dir / img_name

        print(f"  Rendering diagram {counter}...", end=" ")
        if render_mermaid(code, img_path, tools["mmdc"]):
            print("OK")
            return f"![Diagram {counter}]({img_path})"
        else:
            print("FAILED — keeping as code block")
            return match.group(0)

    return pattern.sub(replacer, md_text)


# A line containing only [TOC] marks where the Word table of contents goes
TOC_PLACEHOLDER = re.compile(r"^\[TOC\][ \t]*$", re.MULTILINE)

TOC_OPENXML = """```{=openxml}
<w:sdt>
<w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr>
<w:sdtContent><w:p><w:r>
<w:fldChar w:fldCharType="begin" w:dirty="true"/>
<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>
<w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/>
</w:r></w:p></w:sdtContent>
</w:sdt>
```"""


def read_metadata(md_path: Path, tools: dict[str, str]) -> dict:
    """Return the document's metadata as pandoc JSON AST."""
    result = subprocess.run(
        [tools["pandoc"], str(md_path), "--from", "markdown", "--to", "json"],
        capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode != 0:
        print(f"pandoc error: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    return json.loads(result.stdout)["meta"]


def stringify(node) -> str:
    """Flatten a pandoc JSON AST node to plain text."""
    if isinstance(node, list):
        return "".join(stringify(n) for n in node)
    if not isinstance(node, dict):
        return ""
    t, c = node.get("t"), node.get("c")
    if t in ("Str", "MetaString"):
        return c
    if t in ("Space", "SoftBreak", "LineBreak"):
        return " "
    if t == "Quoted":
        q = '"' if c[0]["t"] == "DoubleQuote" else "'"
        return q + stringify(c[1]) + q
    return stringify(c)


# Pandoc styles that a custom template may lack, built on the template's own
# styles so they inherit its look. Without them Word silently falls back to
# Normal and LibreOffice renders them unpredictably.
_PANDOC_STYLES = {
    "BodyText": '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/>'
                '<w:basedOn w:val="Normal"/><w:qFormat/></w:style>',
    "FirstParagraph": '<w:style w:type="paragraph" w:styleId="FirstParagraph"><w:name w:val="First Paragraph"/>'
                      '<w:basedOn w:val="BodyText"/><w:next w:val="BodyText"/><w:qFormat/></w:style>',
    # Tight list items: no extra space between items, like List Paragraph
    "Compact": '<w:style w:type="paragraph" w:styleId="Compact"><w:name w:val="Compact"/>'
               '<w:basedOn w:val="BodyText"/><w:qFormat/><w:pPr><w:contextualSpacing/></w:pPr></w:style>',
    "VerbatimChar": '<w:style w:type="character" w:customStyle="1" w:styleId="VerbatimChar">'
                    '<w:name w:val="Verbatim Char"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr></w:style>',
    "Table": '<w:style w:type="table" w:styleId="Table"><w:name w:val="Table"/>'
             '<w:basedOn w:val="TableGrid"/></w:style>',
}


def _add_missing_styles(styles_xml: bytes, document_xml: bytes) -> bytes:
    """Define styles referenced by the document but absent from the template."""
    defined = set(re.findall(rb'w:styleId="([^"]+)"', styles_xml))
    used = set(re.findall(rb'<w:(?:pStyle|rStyle|tblStyle) w:val="([^"]+)"', document_xml))
    additions = []
    for style_id in sorted(used - defined):
        definition = _PANDOC_STYLES.get(style_id.decode())
        if definition is None:
            continue
        if style_id == b"Table" and b"TableGrid" not in defined:
            definition = definition.replace("TableGrid", "TableNormal")
        additions.append(definition.encode("utf-8"))
        defined.add(style_id)
    # FirstParagraph/Compact build on BodyText
    if additions and b"BodyText" not in defined:
        additions.insert(0, _PANDOC_STYLES["BodyText"].encode("utf-8"))
    return styles_xml.replace(b"</w:styles>", b"".join(additions) + b"</w:styles>")


def postprocess_docx(docx_path: Path, core_title: str | None):
    """Patch the generated .docx: fill in missing styles, set the core title."""
    with zipfile.ZipFile(docx_path) as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]
    parts = {info.filename: data for info, data in items}
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == "word/styles.xml":
                data = _add_missing_styles(data, parts["word/document.xml"])
            elif info.filename == "docProps/core.xml" and core_title:
                # Word title fields/content controls read dc:title
                data = re.sub(
                    rb"<dc:title>.*?</dc:title>|<dc:title/>",
                    b"<dc:title>" + escape(core_title).encode("utf-8") + b"</dc:title>",
                    data, count=1, flags=re.DOTALL,
                )
            zout.writestr(info, data)


def convert(
    input_path: Path,
    output_path: Path,
    tools: dict[str, str],
    template_path: Path | None = None,
):
    """Main conversion: Markdown → rendered Mermaid PNGs → pandoc → .docx."""
    md_text = input_path.read_text(encoding="utf-8")

    with tempfile.TemporaryDirectory(prefix="md2docx_") as tmp:
        work_dir = Path(tmp)

        # Render Mermaid diagrams to PNGs
        print(f"Processing {input_path.name}...")
        processed_md = process_markdown(md_text, work_dir, tools)

        # Place the TOC at a [TOC] line if present, else pandoc puts it first
        has_toc_placeholder = bool(TOC_PLACEHOLDER.search(processed_md))
        processed_md = TOC_PLACEHOLDER.sub(lambda _: TOC_OPENXML, processed_md)

        # Write processed markdown
        processed_path = work_dir / "processed.md"
        processed_path.write_text(processed_md, encoding="utf-8")

        # Build pandoc command
        cmd = [
            tools["pandoc"],
            str(processed_path),
            "-o", str(output_path),
            "--from", "markdown",
            "--to", "docx",
            # Relative image paths resolve against the input file's folder
            "--resource-path", os.pathsep.join([str(input_path.parent.resolve()), str(work_dir)]),
        ]

        if not has_toc_placeholder:
            cmd += ["--toc", "--toc-depth=3"]

        # title-block: false -> body supplies its own cover page; keep the
        # title only as a document property (used by template headers)
        meta = read_metadata(processed_path, tools)
        core_title = None
        if meta.get("title-block") == {"t": "MetaBool", "c": False}:
            core_title = stringify(meta.get("title"))
            for key in ("title", "subtitle", "author", "date"):
                if key in meta:
                    cmd += ["--metadata", f"{key}="]

        if template_path:
            cmd += ["--reference-doc", str(template_path)]

        print(f"Running pandoc...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"pandoc error: {result.stderr}", file=sys.stderr)
            sys.exit(1)

        postprocess_docx(output_path, core_title)

        print(f"Done → {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Convert Markdown + Mermaid to .docx")
    parser.add_argument("input", type=Path, help="Input Markdown file")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Output .docx path")
    parser.add_argument(
        "--template", type=Path, default=None,
        help="Word reference doc for styling (--reference-doc in pandoc)",
    )
    args = parser.parse_args()

    if not args.input.exists():
        print(f"Error: {args.input} not found", file=sys.stderr)
        sys.exit(1)

    if args.output is None:
        args.output = args.input.with_suffix(".docx")

    tools = check_dependencies()
    convert(args.input, args.output, tools, args.template)


if __name__ == "__main__":
    _ensure_node_on_path()
    main()
