import unittest

from kdpbuilder.markdown_parser import MarkdownParser


class TestMarkdownParserInlineStyles(unittest.TestCase):
    def test_emphasis_and_strong_with_asterisks(self):
        line = "This is *italic* and **bold**."
        segments = MarkdownParser.parse_line(line)
        self.assertEqual(
            segments,
            [
                ("This is ", "normal", ""),
                ("italic", "emphasis", ""),
                (" and ", "normal", ""),
                ("bold", "strong", ""),
                (".", "normal", ""),
            ],
        )

    def test_emphasis_and_strong_with_underscores(self):
        line = "__bold__ then _italic_"
        segments = MarkdownParser.parse_line(line)
        self.assertEqual(
            segments,
            [
                ("bold", "strong", ""),
                (" then ", "normal", ""),
                ("italic", "emphasis", ""),
            ],
        )

    def test_preserves_whitespace_between_custom_styled_segments(self):
        line = "A {B}[strong] C"
        segments = MarkdownParser.parse_line(line)
        self.assertEqual(
            segments,
            [
                ("A ", "normal", ""),
                ("B", "strong", ""),
                (" C", "normal", ""),
            ],
        )


class TestMarkdownParserParseInline(unittest.TestCase):
    def test_parse_inline_does_not_create_headings(self):
        text = "# Not a heading"
        segments = MarkdownParser.parse_inline(text)
        self.assertEqual(segments, [("# Not a heading", "normal", "")])

    def test_parse_inline_supports_emphasis(self):
        text = "A *b* and **c**"
        segments = MarkdownParser.parse_inline(text)
        self.assertEqual(
            segments,
            [
                ("A ", "normal", ""),
                ("b", "emphasis", ""),
                (" and ", "normal", ""),
                ("c", "strong", ""),
            ],
        )


if __name__ == "__main__":
    unittest.main()
