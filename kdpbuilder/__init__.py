"""kdpbuilder - Markdown to DOCX converter for KDP.

Public API:
- convert_markdown_to_docx
- main (CLI entrypoint)
"""

from .convert import convert_markdown_to_docx
from .cli import main

__all__ = ["convert_markdown_to_docx", "main"]
