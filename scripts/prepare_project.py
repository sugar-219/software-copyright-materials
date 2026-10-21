"""Inventory project evidence; candidates are not an approved filing scope."""
import argparse
import hashlib
import io
import json
import os
import re
import stat
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

SKIP_DIRS = {".git", ".venv", "venv", "node_modules", "__pycache__", "dist", "build", "vendor",
             "软件著作权申请资料"}
DOCUMENTS = {".md", ".rst", ".txt", ".docx", ".pdf", ".odt", ".doc", ".wps"}
CONFIG = {".json", ".yaml", ".yml", ".toml", ".ini", ".cfg", ".lock", ".xml", ".csv", ".tsv"}
BINARY = {".exe", ".dll", ".pyc", ".pyd", ".zip", ".7z", ".rar", ".png", ".jpg", ".jpeg",
          ".gif", ".webp", ".ico", ".mp4", ".mp3", ".ttf", ".woff", ".woff2", ".db", ".sqlite"}
MAX_BYTES = 2 * 1024 * 1024


def document_summary(path, raw):
    result = {"text_extracted": False, "excerpt": "", "headings": [],
              "excerpt_location": "提取文本开头；摘要不覆盖全文，关键结论须回看原文件"}
    try:
        if path.suffix.lower() in {".docx", ".odt"}:
            member = "word/document.xml" if path.suffix.lower() == ".docx" else "content.xml"
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                if archive.getinfo(member).file_size > 5 * 1024 * 1024:
                    raise ValueError("文档XML过大，需直接阅读")
                tree = ET.fromstring(archive.read(member))
            paragraphs = []
            for node in tree.iter():
                if node.tag.endswith("}p"):
                    text = ("".join(t.text or "" for t in node.iter() if t.tag.endswith("}t"))
                            if path.suffix.lower() == ".docx" else "".join(node.itertext()))
                    if text.strip():
                        paragraphs.append(text)
            text = "\n".join(paragraphs)
        elif path.suffix.lower() == ".pdf":
            from pypdf import PdfReader
            parts = []
            for page in PdfReader(io.BytesIO(raw)).pages:
                parts.append(page.extract_text() or "")
                if sum(map(len, parts)) >= 8000:
                    break
            text = "\n".join(parts)
        elif path.suffix.lower() in {".doc", ".wps"}:
            result["extraction_status"] = "仅登记；请使用可用文档工具读取"
            return result
        else:
            try:
                text = raw.decode("utf-16" if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig")
            except UnicodeError:
                text = raw.decode("gb18030", errors="strict")
        result.update(text_extracted=bool(text.strip()), excerpt=text[:1000],
                      headings=[item.strip() for item in re.findall(r"^#{1,6}\s+(.+)$", text, flags=re.M)[:24]],
                      extraction_status="已提取摘要" if text.strip() else "未提取到文字；需直接阅读或检查扫描件")
    except Exception as exc:
        result["extraction_status"] = "提取失败或能力不可用，请直接阅读：" + type(exc).__name__
    return result


def is_link(path):
    info = path.lstat()
    return stat.S_ISLNK(info.st_mode) or bool(getattr(info, "st_file_attributes", 0)
                                            & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))


def inventory(project, out_dir):
    project, out_dir = Path(project).resolve(strict=True), Path(out_dir).resolve()
    if not project.is_dir():
        raise ValueError("项目路径不是目录")
    if out_dir == project or project.is_relative_to(out_dir):
        raise ValueError("清单目录不能等于或包含项目根目录")
    report = {"project_root": str(project), "source_candidates": [], "documents": [],
              "configuration_evidence": [], "excluded": [], "needs_review": [],
              "limitations": ["清单仅提供路径、哈希和统计，不自动确认业务、版权归属或源码范围",
                              "未知文本进入候选；被排除的配置、依赖、生成目录仍需人工核对",
                              "文档摘要只供选读；失败、大文件及二进制文档需进一步阅读，不将摘要当成功能证明"]}
    def record(path, reason, group="excluded"):
        report[group].append({"path": path.relative_to(project).as_posix(), "reason": reason})
    for root, dirs, names in os.walk(project, followlinks=False):
        root = Path(root)
        for name in sorted(dirs):
            path = root / name
            if is_link(path) or name.lower() in SKIP_DIRS or path.resolve() == out_dir:
                dirs.remove(name)
                record(path, "链接、依赖/生成目录或本次清单目录；未遍历")
        dirs.sort()
        for name in sorted(names):
            path = root / name
            if is_link(path):
                record(path, "链接文件；不跟随")
                continue
            suffix = path.suffix.lower()
            if name.lower().startswith(".env") or suffix in {".pem", ".key", ".pfx"}:
                record(path, "可能含敏感配置/私钥；不读取正文")
                continue
            if suffix in BINARY:
                record(path, "媒体、二进制或打包文件")
                continue
            if path.stat().st_size > MAX_BYTES:
                record(path, "文件超过2MiB，请人工分类", "needs_review")
                continue
            try:
                raw = path.read_bytes()
            except OSError:
                record(path, "无法读取，请人工核对", "needs_review")
                continue
            item = {"path": path.relative_to(project).as_posix(), "bytes": len(raw),
                    "sha256": hashlib.sha256(raw).hexdigest()}
            if suffix in DOCUMENTS:
                item["requires_content_review"] = True
                item.update(document_summary(path, raw))
                report["documents"].append(item)
                continue
            encoding, text = None, None
            for candidate in ["utf-8-sig", "utf-16" if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else "gb18030"]:
                try:
                    text = raw.decode(candidate)
                    encoding = candidate
                    break
                except UnicodeError:
                    pass
            if text is None or "\0" in text or any(ord(c) < 32 and c not in "\r\n\t\f" for c in text):
                record(path, "二进制或编码待确认", "needs_review")
                continue
            item.update(encoding=encoding, physical_lines=len(text.splitlines()),
                        nonblank_lines=sum(bool(line.strip()) for line in text.splitlines()))
            if suffix in CONFIG or name.lower() in {"dockerfile", "makefile", "requirements.txt"}:
                report["configuration_evidence"].append(item)
            elif text.strip():
                item.update(selected=False, reason="待模型阅读并结合用户确认分类；不是自动入选")
                report["source_candidates"].append(item)
            else:
                record(path, "空文件")
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", required=True)
    parser.add_argument("--out-dir", required=True)
    args = parser.parse_args()
    report = inventory(args.project, args.out_dir)
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    (out / "project-evidence.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = ["# 项目证据与源码候选", "", *report["limitations"], ""]
    for group in ["documents", "configuration_evidence", "source_candidates", "excluded", "needs_review"]:
        lines.extend([f"## {group}", ""])
        lines.extend("- " + json.dumps(item, ensure_ascii=False) for item in report[group])
        lines.append("")
    (out / "project-evidence.md").write_text("\n".join(lines), encoding="utf-8")
    print(json.dumps({group: len(report[group]) for group in ["documents", "source_candidates", "needs_review"]}))


if __name__ == "__main__":
    main()
