import tempfile
import unittest
from pathlib import Path

from docx import Document

from kdpbuilder.convert import convert_markdown_to_docx


class TestNumberedListsInDocx(unittest.TestCase):
    def test_ordered_markdown_lists_become_word_numbered_lists(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)

            md = "1. first\n    2. nested\n3) third\n"
            (tmp_path / "input.md").write_text(md, encoding="utf-8")

            (tmp_path / "styles.yaml").write_text(
                "styles:\n  normal: {}\n", encoding="utf-8"
            )
            (tmp_path / "layout.yaml").write_text(
                "layout: {}\n", encoding="utf-8"
            )

            out_path = tmp_path / "out.docx"
            convert_markdown_to_docx(
                str(tmp_path / "input.md"),
                str(tmp_path / "styles.yaml"),
                str(tmp_path / "layout.yaml"),
                str(out_path),
            )

            d = Document(str(out_path))
            nonempty = [p for p in d.paragraphs if p.text.strip()]
            styles = [p.style.name for p in nonempty]

            self.assertEqual(styles[0], "List Number")
            self.assertEqual(styles[1], "List Number 2")
            self.assertEqual(styles[2], "List Number")


if __name__ == "__main__":
    unittest.main()
