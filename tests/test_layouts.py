import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from build_slides import InputError, render
from layouts import ACCENT, NAVY, table_pages
from pptx_document import emu, write_pptx


ROOT = Path(__file__).resolve().parents[1]
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def mixed_document():
    structure = {
        "type": "structure", "columns": ["入力", "出力"],
        "nodes": [
            {"id": "x", "col": 0, "row": 0, "label": "入力 x"},
            {"id": "y", "col": 1, "row": 0, "label": "結果 y"},
        ],
        "links": [{"from": "x", "to": "y"}],
    }
    return {
        "project": "検査用", "slides": [
            {"layout": "mixed", "title": "見出しの大きさ", "source": "検査用",
             "rows": [
                 {"blocks": [structure], "right_blocks": [
                     {"type": "equation", "lines": [r"y=\sqrt{x}", r"x\geq 0"]},
                     {"type": "code", "code": "def f(x):\n    return x ** 0.5"},
                 ]},
                 {"blocks": [{"type": "text", "text": "本文の文字サイズを調べます。"}]},
             ]},
        ],
    }


class LayoutTests(unittest.TestCase):
    def test_native_shapes_and_physical_a4_size(self):
        slides = render(mixed_document(), ROOT, heading_size=22)
        with TemporaryDirectory() as folder:
            file = Path(folder) / "output.pptx"
            write_pptx(file, slides, NAVY, ACCENT)
            with ZipFile(file) as archive:
                self.assertIsNone(archive.testzip())
                deck = ET.fromstring(archive.read("ppt/presentation.xml"))
                slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        size = deck.find("p:sldSz", NS)
        self.assertEqual(int(size.attrib["cx"]), emu(297 * 72 / 25.4))
        self.assertEqual(int(size.attrib["cy"]), emu(210 * 72 / 25.4))
        runs = {r.findtext("a:t", namespaces=NS): r.find("a:rPr", NS) for r in slide.findall(".//a:r", NS)}
        self.assertEqual(runs["見出しの大きさ"].attrib["sz"], "2200")
        self.assertEqual(runs["本文の文字サイズを調べます。"].attrib["sz"], "1300")
        self.assertEqual(runs["    return x ** 0.5"].find("a:latin", NS).attrib["typeface"], "Consolas")
        self.assertEqual(len(slide.findall(".//m:oMath", NS)), 2)
        self.assertIsNotNone(slide.find(".//m:rad", NS))
        self.assertIsNotNone(slide.find(".//a:tailEnd", NS))

    def test_existing_examples_and_a4_capacity(self):
        examples = json.loads((ROOT / "examples.json").read_text(encoding="utf-8"))
        for page_size in ("a4", "16:9"):
            with self.subTest(page_size=page_size):
                self.assertEqual(len(render(examples, ROOT, page_size=page_size)), 15)
        detail = {"title": "一覧", "headers": ["番号"], "rows": [[i] for i in range(35)]}
        self.assertLess(len(table_pages(detail, "a4")), len(table_pages(detail, "16:9")))

    def test_bad_inputs_point_to_the_problem(self):
        with self.assertRaisesRegex(ValueError, "20〜24"):
            render(mixed_document(), ROOT, heading_size=30)
        document = mixed_document()
        first_right = document["slides"][0]["rows"][0]["right_blocks"]
        first_right[0]["latex"] = "x=1"
        with self.assertRaisesRegex(InputError, "latex か lines"):
            render(document, ROOT)
        del first_right[0]["latex"]
        first_right[1]["code"] = "x" * 200
        with self.assertRaisesRegex(InputError, "コード行が長すぎます"):
            render(document, ROOT)


if __name__ == "__main__":
    unittest.main()
