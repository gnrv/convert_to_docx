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
from pathlib import Path

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
            "--toc",                    # table of contents
            "--toc-depth=3",
            "--resource-path", str(work_dir),
        ]

        if template_path:
            cmd += ["--reference-doc", str(template_path)]

        print(f"Running pandoc...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"pandoc error: {result.stderr}", file=sys.stderr)
            sys.exit(1)

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
