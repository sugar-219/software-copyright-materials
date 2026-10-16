"""Audit source fidelity, DOCX layout, PDF fonts and pages; render every page."""
import argparse
import hashlib
import json
import re
from pathlib import Path
from urllib.parse import urlparse

from docx import Document
from docx.oxml.ns import qn
from PIL import Image, ImageDraw
from pypdf import PdfReader
import pypdfium2 as pdfium

from materials import FILES, read_json, resolve_source, write_json


def font_info(value):
    obj = value.get_object()
    name = str(obj.get("/BaseFont", ""))
    if "/DescendantFonts" in obj:
        child = obj["/DescendantFonts"][0].get_object()
        descriptor = child.get("/FontDescriptor")
    else:
        descriptor = obj.get("/FontDescriptor")
    embedded = False
    if descriptor:
        descriptor = descriptor.get_object()
        embedded = any(key in descriptor for key in ["/FontFile", "/FontFile2", "/FontFile3"])
    return {"name": name, "embedded": embedded}


def inspect_docx(path, errors):
    doc = Document(path)
    for section in doc.sections:
        if abs(section.page_width.mm - 210) > 0.2 or abs(section.page_height.mm - 297) > 0.2:
            errors.append(path.name + "：页面不是 A4")
    paragraphs = list(doc.paragraphs)
    for section in doc.sections:
        paragraphs.extend(section.header.paragraphs)
        paragraphs.extend(section.footer.paragraphs)
    allowed = {"宋体", "微软雅黑", "Consolas", "SimSun", "Microsoft YaHei"}
    for p in paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            props = r._element.rPr
            faces = props.find(qn("w:rFonts")) if props is not None else None
            if faces is None:
                errors.append(path.name + "：文本没有显式字体设置")
                continue
            for key in ["ascii", "hAnsi", "eastAsia", "cs"]:
                if faces.get(qn("w:" + key)) not in allowed:
                    errors.append(path.name + "：出现缺失或非约定字体")
                    break
    return doc


def inspect_pdf(path, render_dir, errors, pending):
    reader = PdfReader(path)
    fonts = {}
    for number, page in enumerate(reader.pages, 1):
        width, height = float(page.mediabox.width), float(page.mediabox.height)
        if abs(width - 595.276) > 1.5 or abs(height - 841.89) > 1.5:
            errors.append(f"{path.name} 第{number}页不是 A4")
        resources = page.get("/Resources", {}).get_object()
        for item in resources.get("/Font", {}).get_object().values() if "/Font" in resources else []:
            info = font_info(item)
            fonts[info["name"]] = info
        if "\ufffd" in (page.extract_text() or ""):
            pending.append(f"{path.name} 第{number}页提取文本出现替代字符，需核对显示和编码")
    allowed = ["simsun", "song", "宋体", "microsoftyahei", "msyh", "微软雅黑", "consolas"]
    for info in fonts.values():
        normalized = re.sub(r"[^a-z\u4e00-\u9fff]", "", info["name"].lower())
        if not any(name in normalized for name in allowed):
            errors.append(path.name + "：PDF出现非约定字体 " + info["name"])
        if not info["embedded"]:
            pending.append(path.name + "：PDF字体未确认内嵌 " + info["name"])
    render_dir.mkdir(parents=True, exist_ok=True)
    pdf = pdfium.PdfDocument(str(path))
    thumbnails = []
    try:
        for number in range(len(pdf)):
            page = pdf[number]
            textpage = page.get_textpage()
            width, height = page.get_size()
            for index in range(textpage.count_chars()):
                left, bottom, right, top = textpage.get_charbox(index)
                if right > left and top > bottom and (left < -1 or bottom < -1 or right > width + 1 or top > height + 1):
                    errors.append(f"{path.name} 第{number + 1}页存在页外文字")
                    break
            textpage.close()
            bitmap = page.render(scale=1.4)
            img = bitmap.to_pil().convert("RGB")
            img.save(render_dir / f"page-{number + 1:03}.png")
            img.thumbnail((280, 396))
            canvas = Image.new("RGB", (300, 424), "#dddddd")
            canvas.paste(img, ((300 - img.width) // 2, 5))
            ImageDraw.Draw(canvas).text((10, 405), str(number + 1), fill="black")
            thumbnails.append(canvas)
            bitmap.close()
            page.close()
    finally:
        pdf.close()
    for start in range(0, len(thumbnails), 12):
        batch = thumbnails[start:start + 12]
        sheet = Image.new("RGB", (900, 424 * ((len(batch) + 2) // 3)), "white")
        for i, img in enumerate(batch):
            sheet.paste(img, ((i % 3) * 300, (i // 3) * 424))
        sheet.save(render_dir / f"contact-{start // 12 + 1:02}.jpg", quality=88)
    return {"pages": len(reader.pages), "fonts": list(fonts.values()), "renders": str(render_dir)}


def inspect(work):
    work = Path(work).resolve(strict=True)
    data = read_json(work / "build-input.json")
    export = read_json(work / "word-export-report.json")
    build = read_json(work / "build-report.json")
    mapping = read_json(work / "source-map.json")
    output, project = Path(data["output_dir"]), Path(data["project_root"])
    errors, pending, pdfs = [], [], {}
    for item in mapping["files"]:
        path = resolve_source(project, item["path"])
        if hashlib.sha256(path.read_bytes()).hexdigest() != item["sha256"]:
            errors.append("项目源码在生成后发生变化：" + item["path"])
    expected_rows = [row["text"] for page in mapping["submitted_rows"] for row in page]
    source = inspect_docx(output / FILES["source"], errors)
    if [p.text for p in source.paragraphs] != expected_rows:
        errors.append("源程序Word正文与源码映射不一致")
    expected_pages = (list(range(1, mapping["full_pages"] + 1)) if mapping["full_pages"] <= 60
                      else list(range(1, 31)) + list(range(mapping["full_pages"] - 29, mapping["full_pages"] + 1)))
    if mapping["selected_pages"] != expected_pages:
        errors.append("源程序取页范围不符合一般交存的连续性")
    for item in export["documents"]:
        if item["kind"] != "source":
            inspect_docx(Path(item["docx"]), errors)
        report = inspect_pdf(Path(item["pdf"]), work / "renders" / item["kind"], errors, pending)
        pdfs[item["kind"]] = report
        if report["pages"] != item["pages"]:
            errors.append(item["kind"] + "：Word与PDF页数不同")
        if item["kind"] == "source" and report["pages"] != build["expected_source_pages"]:
            errors.append("源程序实际分页与取页计划不一致，需调整版式重新生成")
        if item["kind"] == "source" and report["pages"] == build["expected_source_pages"]:
            source_pdf = PdfReader(item["pdf"])
            for number, (page, rows) in enumerate(zip(source_pdf.pages, mapping["submitted_rows"]), 1):
                visible = re.sub(r"\s", "", page.extract_text() or "")
                cursor = 0
                for row in rows:
                    value = re.sub(r"\s", "", row["text"])
                    if not value:
                        continue
                    found = visible.find(value, cursor)
                    if found < 0:
                        errors.append(f"源程序PDF第{number}页内容与所选源码行不一致")
                        break
                    cursor = found + len(value)
        if item["kind"] == "manual" and report["pages"] > 60:
            errors.append("说明书超过60页，当前脚本未对说明书自动取页；需按已确认规则处理并复核Word/PDF对应关系")
        if item["kind"] in ["source", "manual"]:
            minimum = 50 if item["kind"] == "source" else 30
            below = [p["page"] for p in item["page_statistics"] if p["word_lines"] < minimum]
            if below:
                pending.append(f"{item['kind']} Word统计行数低于一般阈值的页面{below}；截图、末页等例外不能自动判定，需核实规则")
    expected = {FILES["source"], FILES["manual"], FILES["application"], "源程序鉴别材料.pdf", "软件说明书.pdf"}
    if data["software"]["development_mode"] == "合作开发":
        expected.add(FILES["agreement"])
    actual = {p.name for p in output.iterdir() if p.is_file() and not p.name.startswith("~$")}
    if actual != expected:
        errors.append("交付目录的文件集合不符合5/6文件约定")
    for issue in build["application_issues"]:
        pending.append(issue["field"] + "：" + issue["issue"])
    rules = data.get("rules", {})
    if not rules.get("checked_at") or not rules.get("sources"):
        pending.append("未记录本次官网规则核对日期或来源")
    for source in rules.get("sources", []):
        domain = urlparse(source.get("url", "")).hostname or ""
        if not any(domain == base or domain.endswith("." + base) for base in ["ncac.gov.cn", "ccopyright.com.cn"]):
            pending.append("规则来源不是国家版权局或版权中心官方域名，需要核实")
    pending.extend(rules.get("pending", []))
    pending.append("需要agent逐页查看渲染图，并核对真实功能、截图来源及用户确认事项；自动检查不代表登记审查通过")
    report = {"technical_checks_passed": not errors, "errors": list(dict.fromkeys(errors)),
              "pending_review": list(dict.fromkeys(pending)), "pdfs": pdfs,
              "delivery_files": sorted(actual), "status": "technical_errors" if errors else "pending_agent_review"}
    write_json(work / "inspection.json", report)
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--work", required=True)
    args = parser.parse_args()
    report = inspect(args.work)
    print(json.dumps({"technical_checks_passed": report["technical_checks_passed"],
                      "errors": report["errors"], "pending_count": len(report["pending_review"])}, ensure_ascii=False))
    return 0 if report["technical_checks_passed"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
