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


if __name__ == "__main__":
    unittest.main()
