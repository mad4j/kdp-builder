from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .convert import convert_markdown_to_docx


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert Markdown with style attributes to DOCX format.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  %(prog)s -m input.md -s styles.yaml -l layout.yaml -o output.docx

The Markdown file supports custom style attributes:
  {text}[style_name] - Apply custom style to text
  # Heading 1 - Standard Markdown heading
""",
    )

    parser.add_argument("-m", "--markdown", required=True, help="Input Markdown file")
    parser.add_argument(
        "-s", "--styles", required=True, help="YAML file with style definitions"
    )
    parser.add_argument(
        "-l", "--layout", required=True, help="YAML file with layout definition"
    )
    parser.add_argument("-o", "--output", required=True, help="Output DOCX file")
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    for file_path, name in [
        (args.markdown, "Markdown"),
        (args.styles, "Styles"),
        (args.layout, "Layout"),
    ]:
        if not Path(file_path).exists():
            print(f"Error: {name} file not found: {file_path}", file=sys.stderr)
            raise SystemExit(1)

    try:
        convert_markdown_to_docx(args.markdown, args.styles, args.layout, args.output)
        print(f"Document saved to: {args.output}")
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
