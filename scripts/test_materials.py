"""Focused mechanical checks. Fixtures are not filing materials."""
import tempfile
import unittest
import argparse
import sys
from pathlib import Path

from docx import Document

from materials import (CodeWrapper, FILES, build, collect_source,
                       make_source, read_json, resolve_source, select_pages, write_json)


class MaterialsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.wrapper = CodeWrapper()

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.project = self.root / "project"
        self.project.mkdir()
        (self.project / "main.py").write_text("# 中文注释\n\ndef main():\n\tprint('验证缩进')\n", encoding="utf-8")
        self.data = {"software": {"name": "排版测试软件", "version": "2.3", "development_mode": "单独开发", "confirmed": True},
                     "source_files": [{"path": "main.py"}],
                     "manual": {"sections": [{"heading": "测试用途", "blocks": [{"text": "本材料仅验证排版与程序，不用于软著申请。"}]}]}}
        self.data_path = self.root / "input.json"
        write_json(self.data_path, self.data)

    def tearDown(self):
        self.tmp.cleanup()

    def test_source_path_escape_rejected(self):
        (self.root / "outside.py").write_text("print(1)")
        with self.assertRaises(ValueError):
            resolve_source(self.project.resolve(), "../outside.py")

    def test_duplicate_source_rejected(self):
        with self.assertRaises(ValueError):
            collect_source(self.project, [{"path": "main.py"}, {"path": "main.py"}], self.wrapper)

    def test_invalid_encoding_does_not_replace_characters(self):
        (self.project / "bad.py").write_bytes(b"\xff\xfe\xff")
        with self.assertRaises(UnicodeError):
            collect_source(self.project, [{"path": "bad.py"}], self.wrapper)

    def test_wrap_reconstructs_original_content(self):
        text = "\t" + "变量ABC_123 = '中文内容'; " * 80
        segments = self.wrapper.wrap(text)
        self.assertGreater(len(segments), 1)
        self.assertEqual("".join(segments), text.expandtabs(4))
        self.assertTrue(all(len(segment) > 0 for segment in segments))

    def test_page_boundaries_and_continuity(self):
        for count in [1, 59, 60, 61, 87]:
            rows = [{"text": str(i)} for i in range(count * 55)]
            pages, indices = select_pages(rows)
            expected = list(range(count)) if count <= 60 else list(range(30)) + list(range(count - 30, count))
            self.assertEqual(indices, expected)
            self.assertEqual(len(pages), count)
            self.assertEqual(len(set(indices)), len(indices))

    def test_original_and_displayed_line_counts_are_distinct(self):
        (self.project / "main.py").write_text("a = '" + "中" * 500 + "'\n\nprint(a)\n", encoding="utf-8")
        files, rows = collect_source(self.project, [{"path": "main.py"}], self.wrapper)
        self.assertEqual(files[0]["physical_lines"], 3)
        self.assertEqual(files[0]["nonblank_lines"], 2)
        self.assertGreater(len(rows), 3)

    def test_source_word_is_selected_original_rows(self):
        rows = [{"text": str(i)} for i in range(61 * 55)]
        doc, pages, indices = make_source(self.data["software"], rows)
        self.assertEqual([p.text for p in doc.paragraphs], [r["text"] for i in indices for r in pages[i]])
        self.assertEqual(sum(bool(p.paragraph_format.page_break_before) for p in doc.paragraphs), 59)

    def test_noncooperative_output_and_version(self):
        output, work = self.root / "delivery", self.root / "work"
        build(self.project, self.data_path, output, work)
        self.assertEqual({p.name for p in output.iterdir()}, {FILES[k] for k in ["source", "manual", "application"]})
        text = "\n".join(p.text for p in Document(output / FILES["application"]).paragraphs)
        self.assertIn("2.3", text)
        self.assertNotIn("1.0", text)
        job = read_json(work / "export-job.json")
        self.assertEqual(sum(Path(i["pdf"]).parent == output for i in job["documents"]), 2)

    def test_cooperative_word_only_and_no_assumed_shares(self):
        self.data["software"]["development_mode"] = "合作开发"
        self.data["cooperation"] = {"party_count": 3}
        write_json(self.data_path, self.data)
        output, work = self.root / "delivery", self.root / "work"
        build(self.project, self.data_path, output, work)
        doc = Document(output / FILES["agreement"])
        text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("合作方3", text)
        self.assertNotIn("平均", text)
        self.assertNotIn("50%", text)
        self.assertFalse((output / "合作开发协议模板.pdf").exists())
        agreement = next(i for i in read_json(work / "export-job.json")["documents"] if i["kind"] == "agreement")
        self.assertEqual(Path(agreement["pdf"]).parent, work)

    def test_existing_materials_not_overwritten(self):
        output = self.root / "delivery"
        output.mkdir()
        (output / "existing.txt").write_text("keep")
        with self.assertRaises(ValueError):
            build(self.project, self.data_path, output, self.root / "work")
        self.assertEqual((output / "existing.txt").read_text(), "keep")

    def test_unconfirmed_identity_information_cannot_be_guessed(self):
        self.data["software"]["confirmed"] = False
        write_json(self.data_path, self.data)
        with self.assertRaises(ValueError):
            build(self.project, self.data_path, self.root / "delivery", self.root / "work")

    def test_existing_work_records_not_overwritten(self):
        work = self.root / "work"
        work.mkdir()
        (work / "source-map.json").write_text("keep", encoding="utf-8")
        with self.assertRaises(ValueError):
            build(self.project, self.data_path, self.root / "delivery", work)
        self.assertEqual((work / "source-map.json").read_text(encoding="utf-8"), "keep")

    def test_missing_manual_image_fails_before_delivery(self):
        self.data["manual"]["sections"][0]["blocks"] = [{"image": "missing.png", "origin": "user_provided"}]
        write_json(self.data_path, self.data)
        output = self.root / "delivery"
        with self.assertRaises(FileNotFoundError):
            build(self.project, self.data_path, output, self.root / "work")
        self.assertFalse(output.exists())


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--temp-root")
    arguments, remaining = parser.parse_known_args()
    if arguments.temp_root:
        Path(arguments.temp_root).mkdir(parents=True, exist_ok=True)
        tempfile.tempdir = str(Path(arguments.temp_root).resolve())
    unittest.main(argv=[sys.argv[0], *remaining])
