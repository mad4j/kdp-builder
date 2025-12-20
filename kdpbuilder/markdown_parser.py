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

    # Standard Markdown inline emphasis/strong (simple, non-nested)
    _INLINE_PATTERNS: List[Tuple[str, re.Pattern[str]]] = [
        ("strong", re.compile(r"\*\*(?=\S)(.+?)(?<=\S)\*\*")),
        ("strong", re.compile(r"__(?=\S)(.+?)(?<=\S)__")),
        ("emphasis", re.compile(r"\*(?=\S)(.+?)(?<=\S)\*")),
        ("emphasis", re.compile(r"_(?=\S)(.+?)(?<=\S)_")),
    ]

    @staticmethod
    def _split_inline_markdown_styles(text: str) -> List[Tuple[str, str, str]]:
        """Split a plain-text chunk into styled segments using Markdown markers.

        Returns [(text, style_name, link_target)] where link_target is always "".
        """
        if not text:
            return []

        segments: List[Tuple[str, str, str]] = []
        pos = 0

        while pos < len(text):
            best = None  # (style_name, match)
            for style_name, pattern in MarkdownParser._INLINE_PATTERNS:
                match = pattern.search(text, pos)
                if not match:
                    continue
                if best is None:
                    best = (style_name, match)
                    continue

                _, best_match = best
                if match.start() < best_match.start():
                    best = (style_name, match)
                elif match.start() == best_match.start() and match.end() > best_match.end():
                    # Prefer the longest match when starting at the same position
                    best = (style_name, match)

            if best is None:
                segments.append((text[pos:], "normal", ""))
                break

            style_name, match = best
            if match.start() > pos:
                segments.append((text[pos:match.start()], "normal", ""))

            inner_text = match.group(1)
            segments.append((inner_text, style_name, ""))
            pos = match.end()

        return [seg for seg in segments if seg[0] != ""]

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
                if normal_text != "":
                    segments.append((normal_text, "normal", ""))

            if match_type == "styled":
                segments.append((match.group(1), match.group(2), ""))
            else:
                segments.append((match.group(1), "normal", match.group(2)))

            pos = end

        if pos < len(line):
            remaining = line[pos:]
            if remaining != "":
                segments.append((remaining, "normal", ""))

        if not segments and line != "":
            segments.append((line, "normal", ""))

        # Post-process: apply standard Markdown emphasis/strong to plain segments
        final_segments: List[Tuple[str, str, str]] = []
        for text, style_name, link_target in segments:
            if style_name == "normal" and not link_target:
                final_segments.extend(MarkdownParser._split_inline_markdown_styles(text))
            else:
                final_segments.append((text, style_name, link_target))

        return final_segments
