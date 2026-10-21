"""Regression checks for preflight, source discovery and stale confirmations."""
import argparse
import copy
import importlib
import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from check_environment import check_environment
from confirm_input import confirmation_path, record_confirmation, verify_confirmation
from materials import build, read_json, revise, write_json
from prepare_project import inventory


class PreparationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.project = self.root / "project"
        self.project.mkdir()
        (self.project / "main.py").write_text("print('测试')\n", encoding="utf-8")
        self.data = {"software": {"name": "测试软件", "version": "2.3", "development_mode": "单独开发", "confirmed": True},
                     "source_files": [{"path": "main.py"}], "manual": {"sections": [
                         {"heading": "用途", "blocks": [{"text": "仅为测试夹具，不是申请材料。"}]}]}}
        self.data_path = self.root / "input.json"
        write_json(self.data_path, self.data)
        record_confirmation(self.project, self.data_path, "测试夹具已确认")

    def tearDown(self):
        self.tmp.cleanup()

    def test_missing_confirmation_stops_before_output(self):
        confirmation_path(self.data_path).unlink()
        with self.assertRaisesRegex(ValueError, "缺少内容确认"):
            build(self.project, self.data_path, self.root / "delivery", self.root / "run")
        self.assertFalse((self.root / "delivery").exists())

    def test_changed_identity_content_or_source_scope_invalidates_confirmation(self):
        for group in ["identity", "content", "source"]:
            with self.subTest(group=group):
                data = copy.deepcopy(self.data)
                if group == "identity":
                    data["software"]["version"] = "2.4"
                elif group == "content":
                    data["application"] = {"completion_date": "2026-10-01"}
                else:
                    data["source_layout"] = {"omit_blank_lines": False}
                write_json(self.data_path, data)
                with self.assertRaisesRegex(ValueError, "确认已失效"):
                    verify_confirmation(self.project, self.data_path)

    def test_source_byte_change_stops_build_and_can_be_reconfirmed(self):
        (self.project / "main.py").write_text("print('changed')\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "源码内容"):
            build(self.project, self.data_path, self.root / "delivery", self.root / "run")
        self.assertFalse((self.root / "delivery").exists())
        record_confirmation(self.project, self.data_path, "测试夹具变更后确认")
        build(self.project, self.data_path, self.root / "delivery", self.root / "run")
        self.assertEqual(read_json(self.root / "run" / "input-confirmation.json")["groups"],
                         read_json(confirmation_path(self.data_path))["groups"])

    def test_source_order_is_bound_but_json_format_is_not(self):
        (self.project / "other.py").write_text("print(2)", encoding="utf-8")
        self.data["source_files"].append({"path": "other.py"})
        write_json(self.data_path, self.data)
        record_confirmation(self.project, self.data_path, "测试顺序")
        self.data_path.write_text(json.dumps(self.data, sort_keys=True), encoding="utf-8")
        verify_confirmation(self.project, self.data_path)
        self.data["source_files"].reverse()
        write_json(self.data_path, self.data)
        with self.assertRaisesRegex(ValueError, "确认已失效"):
            verify_confirmation(self.project, self.data_path)

    def test_screenshot_bytes_are_bound(self):
        image = self.root / "screen.png"
        image.write_bytes(b"test-only bytes")
        self.data["manual"]["sections"][0]["blocks"].append({"image": "screen.png", "origin": "user_provided"})
        write_json(self.data_path, self.data)
        record_confirmation(self.project, self.data_path, "测试截图指纹")
        image.write_bytes(b"changed bytes")
        with self.assertRaisesRegex(ValueError, "截图"):
            verify_confirmation(self.project, self.data_path)

    def test_confirmation_rejects_escape_and_other_project(self):
        self.data["source_files"] = [{"path": "../input.json"}]
        write_json(self.data_path, self.data)
        with self.assertRaisesRegex(ValueError, "项目内"):
            record_confirmation(self.project, self.data_path, "测试越界")
        other = self.root / "other"
        other.mkdir()
        with self.assertRaisesRegex(ValueError, "其他项目"):
            verify_confirmation(other, self.data_path)

    def test_false_confirmation_flag_and_empty_note_are_rejected(self):
        with self.assertRaises(ValueError):
            record_confirmation(self.project, self.data_path, " ")
        self.data["software"]["confirmed"] = False
        write_json(self.data_path, self.data)
        with self.assertRaisesRegex(ValueError, "尚未确认"):
            verify_confirmation(self.project, self.data_path)

    def test_revise_accepts_legacy_baseline_but_requires_current_confirmation(self):
        # Synthetic receipts exercise reuse mechanics only; no claim of a Word/PDF audit.
        output, work = self.root / "baseline", self.root / "baseline-run"
        build(self.project, self.data_path, output, work)
        (work / "input-confirmation.json").unlink()
        job = read_json(work / "export-job.json")
        hashes = {}
        for item in job["documents"]:
            Path(item["pdf"]).write_bytes(b"synthetic PDF fixture; not a validated artifact")
            hashes[item["kind"]] = {form: hashlib.sha256(Path(item[form]).read_bytes()).hexdigest()
                                    for form in ["docx", "pdf"]}
        write_json(work / "word-export-report.json", {"documents": job["documents"]})
        write_json(work / "inspection.json", {"technical_checks_passed": True, "artifact_hashes": hashes})
        self.data["application"] = {"purpose": "测试授权范围内的修订"}
        write_json(self.data_path, self.data)
        new_output, new_work = self.root / "revision", self.root / "revision-run"
        with self.assertRaisesRegex(ValueError, "确认已失效"):
            revise(work, self.data_path, new_output, new_work, ["application"])
        self.assertFalse(new_output.exists())
        record_confirmation(self.project, self.data_path, "测试修订确认")
        result = revise(work, self.data_path, new_output, new_work, ["application"])
        self.assertEqual(result["regenerated_documents"], ["application"])
        self.assertEqual(result["reused_documents"], ["manual", "source"])
        self.assertEqual((output / "源程序鉴别材料.docx").read_bytes(),
                         (new_output / "源程序鉴别材料.docx").read_bytes())

    def test_discovery_does_not_follow_external_directory_links(self):
        outside = self.root / "outside"
        outside.mkdir()
        (outside / "hidden.py").write_text("print('outside')")
        try:
            (self.project / "linked").symlink_to(outside, target_is_directory=True)
        except OSError:
            if sys.platform != "win32":
                self.skipTest("当前账户不能创建符号链接")
            import _winapi
            _winapi.CreateJunction(str(outside), str(self.project / "linked"))
        report = inventory(self.project, self.root / "evidence")
        self.assertEqual([i["path"] for i in report["source_candidates"]], ["main.py"])
        self.assertIn("linked", {i["path"] for i in report["excluded"]})

    def test_discovery_includes_unknown_text_and_documents_without_selecting(self):
        (self.project / "logic.newlang").write_text("custom_function()\n", encoding="utf-8")
        (self.project / "设计.docx").write_bytes(b"binary fixture; not a real Word file")
        (self.project / "settings.json").write_text("{}", encoding="utf-8")
        (self.project / ".env").write_text("SECRET=do-not-copy", encoding="utf-8")
        (self.project / "node_modules").mkdir()
        (self.project / "node_modules" / "hidden.js").write_text("third_party()")
        before = {str(p): p.read_bytes() for p in self.project.rglob("*") if p.is_file()}
        report = inventory(self.project, self.root / "evidence")
        self.assertEqual({i["path"] for i in report["source_candidates"]}, {"main.py", "logic.newlang"})
        self.assertTrue(all(not i["selected"] for i in report["source_candidates"]))
        self.assertIn("设计.docx", {i["path"] for i in report["documents"]})
        self.assertEqual([i["path"] for i in report["configuration_evidence"]], ["settings.json"])
        self.assertNotIn("do-not-copy", str(report))
        self.assertEqual(before, {str(p): p.read_bytes() for p in self.project.rglob("*") if p.is_file()})

    def test_discovery_registers_big_files_and_bad_encoding_for_review(self):
        (self.project / "big.custom").write_bytes(b"a" * (2 * 1024 * 1024 + 1))
        (self.project / "bad.custom").write_bytes(b"\xff\x00\xfe")
        report = inventory(self.project, self.root / "evidence")
        self.assertEqual({i["path"] for i in report["needs_review"]}, {"big.custom", "bad.custom"})

    def test_discovery_skips_own_output_and_links(self):
        out = self.project / "evidence"
        out.mkdir()
        (out / "old.custom").write_text("old()")
        report = inventory(self.project, out)
        self.assertEqual([i["path"] for i in report["source_candidates"]], ["main.py"])
        with self.assertRaises(ValueError):
            inventory(self.project, self.project)

    def test_missing_dependency_reports_failure_without_installing(self):
        original = importlib.import_module
        def fake_import(name):
            if name == "pypdf":
                raise ImportError("test")
            return original(name)
        with patch("check_environment.importlib.import_module", side_effect=fake_import):
            report = check_environment()
        self.assertFalse(report["preflight_passed"])
        self.assertFalse(report["modules"]["pypdf"]["available"])
        self.assertFalse(report["word"]["automation_verified"])


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--temp-root", required=True)
    args, remaining = parser.parse_known_args()
    Path(args.temp_root).mkdir(parents=True, exist_ok=True)
    tempfile.tempdir = str(Path(args.temp_root).resolve())
    unittest.main(argv=[sys.argv[0], *remaining])
