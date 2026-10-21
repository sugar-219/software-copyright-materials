"""Feature coverage and evidence checks using isolated, non-filing fixtures."""
import argparse
import copy
import hashlib
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from PIL import Image
from pypdf import PdfWriter

from business_review import business_pending, require_business_review, review_business, write_drafts
from confirm_input import digest, input_groups, record_confirmation, verify_confirmation
from materials import build, read_json, revise, write_json
from prepare_project import document_summary, inventory


class BusinessTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.project = self.root / "project"
        self.project.mkdir()
        (self.project / "main.py").write_text("def export_result():\n    return 'result'\n", encoding="utf-8")
        (self.project / "README.md").write_text("# 合成测试\n导出结果\n", encoding="utf-8")
        Image.new("RGB", (100, 60), "white").save(self.root / "screen.png")
        self.data = {"software": {"name": "业务测试软件", "version": "2.3", "development_mode": "单独开发", "confirmed": True},
                     "source_files": [{"path": "main.py"}], "application": {"functions": "将结果导出为文件。"},
                     "business": {"positioning": "测试业务资料检查", "target_users": ["测试用户"],
                         "operation_flow": ["查看结果", "导出"], "application_feature_ids": ["export"], "features": [
                             {"id": "export", "name": "结果导出", "purpose": "保存结果", "entry": "结果页导出入口",
                              "steps": ["选择格式", "导出"], "feedback": ["生成文件"],
                              "evidence": [{"path": "main.py", "start_line": 1, "end_line": 2},
                                           {"path": "README.md", "location": "合成测试章节"}]}]},
                     "manual": {"sections": [{"heading": "结果导出", "feature_ids": ["export"], "blocks": [
                         {"text": "在结果页选择格式并导出，完成后得到文件。"},
                         {"image": "screen.png", "origin": "user_provided", "caption": "仅为测试图片"}]}]}}
        self.path = self.root / "input.json"
        write_json(self.path, self.data)

    def tearDown(self):
        self.tmp.cleanup()

    def test_valid_coverage_binds_sections_screenshots_and_evidence(self):
        report = require_business_review(self.project, self.data)
        self.assertEqual(report["warnings"], [])
        self.assertEqual(report["features"][0]["manual_sections"], ["结果导出"])
        self.assertEqual(report["features"][0]["screenshots"], ["screen.png"])
        self.assertEqual(set(report["evidence_files"]), {"main.py", "README.md"})

    def test_missing_operation_details_and_duplicate_ids_fail(self):
        self.data["business"]["features"][0].pop("entry")
        self.data["business"]["features"].append(copy.deepcopy(self.data["business"]["features"][0]))
        report = review_business(self.project, self.data)
        self.assertTrue(any("entry" in message for message in report["errors"]))
        self.assertTrue(any("重复" in message for message in report["errors"]))

    def test_unknown_application_and_manual_references_fail(self):
        self.data["business"]["application_feature_ids"] = ["unknown"]
        self.data["manual"]["sections"][0]["feature_ids"] = ["another"]
        report = review_business(self.project, self.data)
        self.assertEqual(len(report["errors"]), 2)

    def test_missing_manual_and_evidence_are_explicit_pending_items(self):
        self.data["manual"]["sections"][0].pop("feature_ids")
        self.data["business"]["features"][0]["evidence"] = []
        report = require_business_review(self.project, self.data)
        self.assertTrue(any("未关联说明书" in m for m in report["warnings"]))
        self.assertTrue(any("缺少文件证据" in m for m in report["warnings"]))
        errors, pending = business_pending(self.project, self.data, report)
        self.assertFalse(errors)
        self.assertEqual(pending, report["warnings"])

    def test_screenshots_are_checked_per_feature_and_skipping_is_not_a_structural_error(self):
        feature = copy.deepcopy(self.data["business"]["features"][0])
        feature.update(id="preview", name="结果预览")
        self.data["business"]["features"].append(feature)
        self.data["manual"]["sections"].append({"heading": "结果预览", "feature_ids": ["preview"],
            "blocks": [{"text": "查看结果。【截图预留：结果页】"}]})
        report = require_business_review(self.project, self.data)
        self.assertEqual(len(report["warnings"]), 1)
        self.assertIn("preview", report["warnings"][0])
        self.assertEqual(report["features"][1]["screenshots"], [])

    def test_missing_evidence_escape_and_invalid_lines_fail(self):
        for reference in [{"path": "missing.py"}, {"path": "../input.json"},
                          {"path": "main.py", "start_line": 1, "end_line": 90}]:
            with self.subTest(reference=reference):
                self.data["business"]["features"][0]["evidence"] = [reference]
                with self.assertRaises(ValueError):
                    require_business_review(self.project, self.data)

    def test_non_source_document_change_invalidates_confirmation(self):
        record_confirmation(self.project, self.path, "测试夹具确认")
        (self.project / "README.md").write_text("项目说明改变\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "确认已失效"):
            verify_confirmation(self.project, self.path)

    def test_business_edit_invalidates_confirmation(self):
        record_confirmation(self.project, self.path, "测试夹具确认")
        self.data["business"]["features"][0]["feedback"] = ["不同结果"]
        write_json(self.path, self.data)
        with self.assertRaisesRegex(ValueError, "确认已失效"):
            verify_confirmation(self.project, self.path)

    def test_legacy_input_keeps_previous_confirmation_digest(self):
        self.data.pop("business")
        write_json(self.path, self.data)
        content = {key: self.data.get(key) for key in ["application", "manual", "cooperation", "field_limits", "fact_evidence"]}
        content["images"] = [{"image": "screen.png", "sha256": hashlib.sha256((self.root / "screen.png").read_bytes()).hexdigest()}]
        self.assertEqual(input_groups(self.project, self.path)["content"], digest(content))
        report = require_business_review(self.project, self.data)
        self.assertFalse(report["enabled"])
        self.assertTrue(report["warnings"])

    def test_readable_drafts_preserve_authored_text_and_do_not_write_project(self):
        before = {p.name: p.read_bytes() for p in self.project.iterdir()}
        report = write_drafts(self.project, self.path, self.root / "drafts")
        self.assertFalse(report["errors"])
        text = (self.root / "drafts" / "materials-draft.md").read_text(encoding="utf-8")
        self.assertIn(self.data["application"]["functions"], text)
        self.assertIn(self.data["manual"]["sections"][0]["blocks"][0]["text"], text)
        self.assertIn("【待本人确认】", text)
        self.assertEqual(before, {p.name: p.read_bytes() for p in self.project.iterdir()})
        with self.assertRaises(ValueError):
            write_drafts(self.project, self.path, self.project)

    def test_build_saves_coverage_and_current_evidence_is_rechecked(self):
        record_confirmation(self.project, self.path, "测试夹具确认")
        build(self.project, self.path, self.root / "delivery", self.root / "run")
        report = read_json(self.root / "run" / "business-review.json")
        self.assertTrue(report["enabled"])
        (self.project / "README.md").write_text("改变文档", encoding="utf-8")
        errors, _ = business_pending(self.project, self.data, report)
        self.assertTrue(any("生成后发生变化" in m for m in errors))

    def test_invalid_business_inputs_report_errors_instead_of_passing(self):
        for business in [None, "wrong", {"features": []},
                         {"features": [{"id": "x", "steps": "wrong", "feedback": [], "pending": None}]}]:
            self.data["business"] = business
            self.assertTrue(review_business(self.project, self.data)["errors"])

    def test_revise_reuses_manual_when_only_internal_feature_links_change(self):
        # Explicit synthetic receipts test reuse only, not real PDF inspection.
        record_confirmation(self.project, self.path, "测试夹具确认")
        output, work = self.root / "baseline", self.root / "baseline-run"
        build(self.project, self.path, output, work)
        job = read_json(work / "export-job.json")
        hashes = {}
        for item in job["documents"]:
            Path(item["pdf"]).write_bytes(b"synthetic PDF fixture")
            hashes[item["kind"]] = {form: hashlib.sha256(Path(item[form]).read_bytes()).hexdigest() for form in ["docx", "pdf"]}
        write_json(work / "word-export-report.json", {"documents": job["documents"]})
        write_json(work / "inspection.json", {"technical_checks_passed": True, "artifact_hashes": hashes})
        self.data["manual"]["sections"][0].pop("feature_ids")
        write_json(self.path, self.data)
        record_confirmation(self.project, self.path, "测试内部关联变化确认")
        result = revise(work, self.path, self.root / "revision", self.root / "revision-run", ["application"])
        self.assertIn("manual", result["reused_documents"])
        self.assertEqual((output / "软件说明书.docx").read_bytes(), (self.root / "revision" / "软件说明书.docx").read_bytes())
        self.data["manual"]["sections"][0]["blocks"][0]["text"] += "测试正文发生修改。"
        write_json(self.path, self.data)
        record_confirmation(self.project, self.path, "测试实际正文变更确认")
        result = revise(work, self.path, self.root / "revision-text", self.root / "revision-text-run", ["application"])
        self.assertIn("manual", result["regenerated_documents"])

    def test_document_summaries_include_real_text_and_retain_failures(self):
        doc = Document()
        doc.add_paragraph("文档中的业务线索，仅供选读。")
        path = self.project / "设计.docx"
        doc.save(path)
        summary = document_summary(path, path.read_bytes())
        self.assertIn("业务线索", summary["excerpt"])
        report = inventory(self.project, self.root / "preparation")
        md = next(item for item in report["documents"] if item["path"] == "README.md")
        self.assertIn("合成测试", md["headings"])
        self.assertFalse(document_summary(self.project / "broken.docx", b"not a zip")["text_extracted"])
        self.assertFalse(document_summary(self.project / "broken.pdf", b"not a pdf")["text_extracted"])

    def test_empty_pdf_and_unsupported_docs_are_marked_for_reading(self):
        writer = PdfWriter()
        writer.add_blank_page(width=595, height=842)
        path = self.project / "empty.pdf"
        with path.open("wb") as stream:
            writer.write(stream)
        result = document_summary(path, path.read_bytes())
        self.assertFalse(result["text_extracted"])
        self.assertIn("需直接阅读", result["extraction_status"])
        result = document_summary(self.project / "legacy.doc", b"binary fixture")
        self.assertFalse(result["text_extracted"])
        self.assertIn("仅登记", result["extraction_status"])


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--temp-root", required=True)
    args, remaining = parser.parse_known_args()
    Path(args.temp_root).mkdir(parents=True, exist_ok=True)
    tempfile.tempdir = str(Path(args.temp_root).resolve())
    unittest.main(argv=[sys.argv[0], *remaining])
