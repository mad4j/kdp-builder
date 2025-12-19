from __future__ import annotations

import re
from typing import List, Tuple


class MarkdownParser:
    """Parses Markdown with custom style attributes."""

    STYLED_TEXT_PATTERN = re.compile(r"\{([^}]+)\}\[([^\]]+)\]")
    HEADER_PATTERN = re.compile(r"^(#{1,6})\s+(.+)$")
    PAGEBREAK_PATTERN = re.compile(r"^<<<pagebreak>>>$", re.IGNORECASE)
    INDEX_PATTERN = re.compile(r"^<<<index:(.+)>>>$", re.IGNORECASE)
    TOC_PATTERN = re.compile(r"^<<<toc>>>$", re.IGNORECASE)
    BOOKMARK_PATTERN = re.compile(r"^<<<bookmark:(.+)>>>$", re.IGNORECASE)
    LINK_PATTERN = re.compile(r"\[([^\]]+)\]\(#([^\)]+)\)")

    @staticmethod
    def is_pagebreak(line: str) -> bool:
        return bool(MarkdownParser.PAGEBREAK_PATTERN.match(line.strip()))

    @staticmethod
    def is_index(line: str) -> Tuple[bool, str]:
        match = MarkdownParser.INDEX_PATTERN.match(line.strip())
        if match:
            return (True, match.group(1))
        return (False, "")

    @staticmethod
    def is_toc(line: str) -> bool:
        return bool(MarkdownParser.TOC_PATTERN.match(line.strip()))

    @staticmethod
    def is_bookmark(line: str) -> Tuple[bool, str]:
        match = MarkdownParser.BOOKMARK_PATTERN.match(line.strip())
        if match:
            return (True, match.group(1))
        return (False, "")

    @staticmethod
    def parse_line(line: str) -> List[Tuple[str, str, str]]:
        """Return [(text, style_name, link_target)]."""
        segments: List[Tuple[str, str, str]] = []
        pos = 0

        header_match = MarkdownParser.HEADER_PATTERN.match(line)
        if header_match:
            level = len(header_match.group(1))
            text = header_match.group(2)
            style_name = f"heading{level}"
            return [(text, style_name, "")]

        matches = []
        for match in MarkdownParser.STYLED_TEXT_PATTERN.finditer(line):
            matches.append(("styled", match.start(), match.end(), match))
        for match in MarkdownParser.LINK_PATTERN.finditer(line):
            matches.append(("link", match.start(), match.end(), match))
        matches.sort(key=lambda x: x[1])

        for match_type, start, end, match in matches:
            if start > pos:
                normal_text = line[pos:start]
                if normal_text.strip():
                    segments.append((normal_text, "normal", ""))

            if match_type == "styled":
                segments.append((match.group(1), match.group(2), ""))
            else:
                segments.append((match.group(1), "normal", match.group(2)))

            pos = end

        if pos < len(line):
            remaining = line[pos:]
            if remaining.strip():
                segments.append((remaining, "normal", ""))

        if not segments and line.strip():
            segments.append((line, "normal", ""))

        return segments
