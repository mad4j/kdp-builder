from __future__ import annotations

from .bookmarks import sanitize_bookmark_name
from .config_io import load_layout, load_styles
from .docx_builder import DocxBuilder
from .markdown_parser import MarkdownParser


def convert_markdown_to_docx(
    markdown_path: str,
    styles_path: str,
    layout_path: str,
    output_path: str,
) -> None:
    """Convert a Markdown file to DOCX using style and layout definitions."""

    styles = load_styles(styles_path)
    layout = load_layout(layout_path)

    builder = DocxBuilder(styles, layout)

    with open(markdown_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if MarkdownParser.is_pagebreak(line):
                builder.add_page_break()
            elif MarkdownParser.is_toc(line):
                builder.add_toc()
            else:
                is_idx, term = MarkdownParser.is_index(line)
                if is_idx:
                    builder.add_index_entry(term)
                    continue

                is_bkmk, name = MarkdownParser.is_bookmark(line)
                if is_bkmk:
                    builder.add_bookmark(name)
                    continue

                if line.strip():
                    segments = MarkdownParser.parse_line(line)

                    auto_bookmark = None
                    if segments and str(segments[0][1]).startswith("heading"):
                        auto_bookmark = sanitize_bookmark_name(segments[0][0])

                    builder.add_paragraph(segments, auto_bookmark)
                else:
                    builder.add_paragraph([])

    builder.apply_header_footer()
    builder.save(output_path)
