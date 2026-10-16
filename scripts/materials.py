"""Create editable filing drafts from explicit project evidence on Windows."""
import argparse
import hashlib
import json
import os
import re
import sys
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from PIL import Image, ImageFont

ROOT = Path(__file__).resolve().parents[1]
FILES = {
    "manual": "软件说明书.docx",
    "source": "源程序鉴别材料.docx",
    "application": "申请信息填写稿.docx",
    "agreement": "合作开发协议模板.docx",
}
ROWS_PER_PAGE = 55
CODE_SIZE = 10
CODE_LEADING = 12


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path, data):
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def font(run_or_style, latin="宋体", east="宋体", size=11):
    run_or_style.font.name = latin
    run_or_style.font.size = Pt(size)
    run_or_style.font.color.rgb = RGBColor(0, 0, 0)
    props = run_or_style._element.get_or_add_rPr()
    fonts = props.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        props.insert(0, fonts)
    for key in list(fonts.attrib):
        del fonts.attrib[key]
    for key, value in {"ascii": latin, "hAnsi": latin, "cs": latin, "eastAsia": east}.items():
        fonts.set(qn("w:" + key), value)


def run(paragraph, text, latin="宋体", east="宋体", size=11):
    value = paragraph.add_run(str(text))
    font(value, latin, east, size)
    return value


def field(paragraph, instruction):
    value = run(paragraph, "", size=9)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    code = OxmlElement("w:instrText")
    code.set(qn("xml:space"), "preserve")
    code.text = " " + instruction + " "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    value._r.extend([begin, code, end])


def base_document(kind):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin = sec.bottom_margin = Mm(18 if kind == "source" else 22)
    sec.left_margin = sec.right_margin = Mm(18 if kind == "source" else 23)
    sec.header_distance = sec.footer_distance = Mm(8)
    for name in ["Normal", "Title", "Heading 1", "Heading 2", "Caption", "Header", "Footer"]:
        style = doc.styles[name]
        heading = name in ["Title", "Heading 1", "Heading 2"]
        font(style, "微软雅黑" if heading else "宋体", "微软雅黑" if heading else "宋体",
             {"Title": 18, "Heading 1": 14, "Heading 2": 12, "Caption": 9,
              "Header": 9, "Footer": 9}.get(name, 11))
        style.paragraph_format.space_before = Pt(8 if heading else 0)
        style.paragraph_format.space_after = Pt(5 if heading else 3)
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.widow_control = True
        if heading:
            style.paragraph_format.keep_with_next = True
    style = doc.styles.add_style("SourceCode", WD_STYLE_TYPE.PARAGRAPH)
    font(style, "Consolas", "宋体", CODE_SIZE)
    fmt = style.paragraph_format
    fmt.space_before = fmt.space_after = Pt(0)
    fmt.line_spacing_rule, fmt.line_spacing = WD_LINE_SPACING.EXACTLY, Pt(CODE_LEADING)
    fmt.keep_with_next = fmt.keep_together = fmt.widow_control = False
    # Some bundled default templates carry a blue rule under Title/Subtitle.
    for border in list(doc.styles.element.xpath(".//w:pBdr")):
        border.getparent().remove(border)
    doc.styles["Header"].paragraph_format.tab_stops.clear_all()
    header = sec.header.paragraphs[0]
    header.paragraph_format.tab_stops.add_tab_stop(
        sec.page_width - sec.left_margin - sec.right_margin, WD_TAB_ALIGNMENT.RIGHT)
    run(header, "{{软件名称}}  {{版本号}}\t", size=9)
    field(header, "PAGE")
    settings = doc.settings.element
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    settings.append(update)
    return doc


def add_title(doc, name, version, title):
    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, title, "微软雅黑", "微软雅黑", 18)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, name + "  " + version)


def heading(doc, text, level=1):
    p = doc.add_paragraph(style="Heading " + str(level))
    run(p, text, "微软雅黑", "微软雅黑", 14 if level == 1 else 12)
    return p


def paragraph(doc, text):
    p = doc.add_paragraph()
    run(p, text)
    return p


def agreement_body(doc, name, version, party_count=2):
    add_title(doc, name, version, "软件合作开发协议")
    paragraph(doc, "待填写模板。各方应在填写、核对全部条款后打印签署；如合作方超过下列人数，可按相同格式增补。")
    heading(doc, "一 合作方")
    for i in range(1, party_count + 1):
        paragraph(doc, f"合作方{i}姓名或名称：【本人填写】")
        paragraph(doc, "证件类型及号码：【本人填写】")
        paragraph(doc, "联系地址及联系方式：【本人填写】")
    heading(doc, "二 合作开发的软件")
    paragraph(doc, "软件全称：" + name)
    paragraph(doc, "软件版本：" + version)
    paragraph(doc, "实际开发期间：【各方确认后填写】")
    heading(doc, "三 实际分工与成果")
    for i in range(1, party_count + 1):
        paragraph(doc, f"合作方{i}实际承担的工作与交付成果：【各方确认后填写】")
    paragraph(doc, "第三方代码、资源及其使用安排（如有）：【各方确认后填写】")
    heading(doc, "四 著作权归属与权利行使")
    paragraph(doc, "本软件著作权归属：【各方协商确认后填写】")
    paragraph(doc, "权利份额及计算方式（如有约定）：【各方协商确认后填写】")
    paragraph(doc, "使用、许可、转让及收益分配安排：【各方协商确认后填写】")
    heading(doc, "五 登记申请办理安排")
    paragraph(doc, "申请登记的著作权人名单及顺序：【各方确认后填写】")
    paragraph(doc, "办理代表及授权事项：【各方确认后填写】")
    paragraph(doc, "材料提供、核对及费用承担安排：【各方确认后填写】")
    heading(doc, "六 其他约定")
    paragraph(doc, "其他权利义务、争议解决及协议生效安排：【各方协商确认后填写】")
    heading(doc, "七 签署")
    for i in range(1, party_count + 1):
        paragraph(doc, f"合作方{i}签字或盖章：________________    日期：________________")
    paragraph(doc, "签署地点：【实际签署时填写】")


def initialize_templates():
    target = ROOT / "assets" / "templates"
    target.mkdir(parents=True, exist_ok=True)
    for kind, filename in FILES.items():
        doc = base_document(kind)
        if kind == "agreement":
            agreement_body(doc, "{{软件名称}}", "{{版本号}}")
        elif kind != "source":
            add_title(doc, "{{软件名称}}", "{{版本号}}", filename[:-5])
            paragraph(doc, "{{由项目真实内容填充正文}}")
        else:
            p = doc.add_paragraph(style="SourceCode")
            run(p, "{{从项目提取源程序}}", "Consolas", "宋体", CODE_SIZE)
        doc.save(target / filename)


def load_template(kind, software):
    doc = Document(ROOT / "assets" / "templates" / FILES[kind])
    body = doc._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    for section in doc.sections:
        for p in section.header.paragraphs:
            for r in p.runs:
                if "{{软件名称}}" in r.text:
                    r.text = r.text.replace("{{软件名称}}", software["name"]).replace("{{版本号}}", software["version"])
                    font(r, size=9)
    return doc


class CodeWrapper:
    def __init__(self, font_dir=None):
        directory = Path(font_dir or Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts")
        self.latin = ImageFont.truetype(str(directory / "consola.ttf"), CODE_SIZE * 4)
        self.east = ImageFont.truetype(str(directory / "simsun.ttc"), CODE_SIZE * 4)
        # Leave 4 mm of safety for differences between Pillow and Word shaping.
        self.width = (210 - 36 - 4) / 25.4 * 72 * 4
        self.widths = {}

    def wrap(self, text):
        expanded = text.expandtabs(4)
        if not expanded:
            return [""]
        chunks, current, width = [], "", 0
        for char in expanded:
            if ord(char) > 0xFFFF or (ord(char) < 32):
                raise ValueError("源程序包含控制字符或尚未支持字体的补充平面字符，请先确认其显示方案")
            if char not in self.widths:
                face = self.latin if ord(char) < 128 else self.east
                self.widths[char] = face.getlength(char)
            advance = self.widths[char]
            if current and width + advance > self.width and not unicodedata.combining(char):
                chunks.append(current)
                current, width = "", 0
            current += char
            width += advance
        chunks.append(current)
        assert "".join(chunks) == expanded
        return chunks


def resolve_source(project, relative):
    path = Path(relative)
    if path.is_absolute() or ".." in path.parts:
        raise ValueError("源码必须使用项目内相对路径：" + relative)
    resolved = (project / path).resolve(strict=True)
    if not resolved.is_relative_to(project):
        raise ValueError("源码链接指向项目外：" + relative)
    if not resolved.is_file():
        raise ValueError("源码不是普通文件：" + relative)
    return resolved


def collect_source(project, specs, wrapper):
    rows, files, seen = [], [], set()
    if not specs:
        raise ValueError("必须明确列出本次纳入的源文件及顺序")
    secret = re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----|\bAKIA[0-9A-Z]{16}\b|\bsk-(?:proj-)?[A-Za-z0-9_-]{24,}")
    for spec in specs:
        relative = spec["path"]
        path = resolve_source(project, relative)
        key = str(path).casefold()
        if key in seen:
            raise ValueError("源文件重复：" + relative)
        seen.add(key)
        raw = path.read_bytes()
        text = raw.decode(spec.get("encoding", "utf-8-sig"), errors="strict")
        if secret.search(text):
            raise ValueError("源文件疑似含真实凭据，先处理，不输出敏感值：" + relative)
        original = re.split(r"\r\n|\n|\r", text) if text else []
        if original and original[-1] == "" and text.endswith(("\n", "\r")):
            original.pop()
        files.append({"path": relative, "encoding": spec.get("encoding", "utf-8-sig"),
                      "sha256": hashlib.sha256(raw).hexdigest(), "physical_lines": len(original),
                      "nonblank_lines": sum(bool(line.strip()) for line in original)})
        for number, line in enumerate(original, 1):
            for segment, content in enumerate(wrapper.wrap(line), 1):
                rows.append({"path": relative, "line": number, "segment": segment, "text": content})
    if not rows:
        raise ValueError("纳入的源码为空")
    return files, rows


def select_pages(rows, per_page=ROWS_PER_PAGE):
    pages = [rows[i:i + per_page] for i in range(0, len(rows), per_page)]
    indices = list(range(len(pages))) if len(pages) <= 60 else list(range(30)) + list(range(len(pages) - 30, len(pages)))
    return pages, indices


def make_source(software, rows):
    doc = load_template("source", software)
    pages, indices = select_pages(rows)
    for position, page_index in enumerate(indices):
        for row_index, row in enumerate(pages[page_index]):
            p = doc.add_paragraph(style="SourceCode")
            if position and row_index == 0:
                p.paragraph_format.page_break_before = True
            run(p, row["text"], "Consolas", "宋体", CODE_SIZE)
    return doc, pages, indices


def make_manual(software, manual, data_dir):
    doc = load_template("manual", software)
    add_title(doc, software["name"], software["version"], "软件说明书")
    evidence = []
    if not manual.get("sections"):
        raise ValueError("说明书需要由 agent 编写真实章节，不能使用占位说明代替")
    for section in manual["sections"]:
        heading(doc, section["heading"])
        if not section.get("blocks"):
            raise ValueError("说明书存在空章节：" + section["heading"])
        for block in section["blocks"]:
            if "text" in block:
                for text in block["text"].split("\n"):
                    if text.strip():
                        paragraph(doc, text)
            elif "image" in block:
                if block.get("origin") not in ["runtime_capture", "user_provided"]:
                    raise ValueError("截图必须标记为实际运行截取或用户提供")
                path = Path(block["image"])
                if not path.is_absolute():
                    path = data_dir / path
                with Image.open(path) as img:
                    width, height = img.size
                    size = min(164 / width, 130 / height)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.keep_with_next = bool(block.get("caption"))
                p.add_run().add_picture(str(path), width=Mm(width * size), height=Mm(height * size))
                if block.get("caption"):
                    p = doc.add_paragraph(style="Caption")
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run(p, block["caption"], size=9)
                evidence.append({"image": str(path.resolve()), "origin": block["origin"],
                                 "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
            else:
                raise ValueError("说明书块仅支持 text 或 image")
    return doc, evidence


def make_application(software, values, stats, limits=None):
    doc = load_template("application", software)
    add_title(doc, software["name"], software["version"], "申请信息填写稿")
    paragraph(doc, "【本人填写】为隐私留空项；【待本人确认】为尚需确认的事实。")
    values = dict(values)
    if values.get("publication_status") == "未发表":
        values.setdefault("first_publication", "不适用（未发表）")
    values.update(name=software["name"], version=software["version"],
                  development_mode=software["development_mode"], source_lines=str(stats["nonblank_lines"]))
    issues = []
    for group in read_json(ROOT / "assets" / "fields.json"):
        heading(doc, group["title"])
        for original_item in group["fields"]:
            item = dict(original_item)
            changes = (limits or {}).get(item["key"], {})
            if set(changes) - {"min", "max"} or any(type(v) is not int or v < 0 for v in changes.values()):
                raise ValueError("字段限制只接受非负整数 min/max：" + item["key"])
            item.update(changes)
            value = values.get(item["key"], item.get("default", "【待本人确认】"))
            if not isinstance(value, str):
                raise ValueError("申请信息必须为字符串：" + item["key"])
            p = doc.add_paragraph()
            p.paragraph_format.keep_with_next = True
            run(p, item["label"] + "：").bold = True
            paragraph(doc, value)
            pending = (not value and not item.get("optional")) or "【" in value
            if pending:
                issues.append({"field": item["label"], "issue": "待填写或确认"})
            elif len(value) > item.get("max", float("inf")) or len(value) < item.get("min", 0):
                issues.append({"field": item["label"], "issue": "长度不符用户页面提示，需按当前页面确认", "length": len(value)})
    return doc, issues


def build(project, data_path, output, work):
    project = Path(project).resolve(strict=True)
    data_path = Path(data_path).resolve(strict=True)
    output, work = Path(output).resolve(), Path(work).resolve()
    if output == work or output.is_relative_to(work) or work.is_relative_to(output):
        raise ValueError("交付目录和工作目录必须分开且不互相包含")
    if project == output or project.is_relative_to(output):
        raise ValueError("交付目录不能覆盖项目根或其上级")
    if output.exists() and any(output.iterdir()):
        raise ValueError("请提供空交付目录，避免覆盖已有材料")
    if work.exists() and any(work.iterdir()):
        raise ValueError("请提供空的单次工作目录，避免覆盖已有记录")
    if not project.is_dir():
        raise ValueError("项目路径不是目录")
    data = read_json(data_path)
    software = data["software"]
    if software.get("confirmed") is not True:
        raise ValueError("软件名称、申请版本和开发方式需由用户确认后置 confirmed=true")
    for key in ["name", "version", "development_mode"]:
        if not isinstance(software.get(key), str) or not software[key].strip() or "【" in software[key]:
            raise ValueError("缺少已确认的软件信息：" + key)
    if software["development_mode"] not in ["单独开发", "合作开发", "委托开发", "下达任务开发"]:
        raise ValueError("开发方式需明确")
    files, rows = collect_source(project, data["source_files"], CodeWrapper())
    stats = {key: sum(item[key] for item in files) for key in ["physical_lines", "nonblank_lines"]}
    docs = {}
    docs["source"], pages, indices = make_source(software, rows)
    docs["manual"], images = make_manual(software, data["manual"], data_path.parent)
    docs["application"], issues = make_application(software, data.get("application", {}), stats,
                                                   data.get("field_limits"))
    if software["development_mode"] == "合作开发":
        count = data.get("cooperation", {}).get("party_count", 2)
        if type(count) is not int or count < 2:
            raise ValueError("合作协议人数需至少为 2；通用模板默认 2，更多合作方可增补")
        docs["agreement"] = load_template("agreement", software)
        agreement_body(docs["agreement"], software["name"], software["version"], count)
    output.mkdir(parents=True, exist_ok=True)
    work.mkdir(parents=True, exist_ok=True)
    for kind, doc in docs.items():
        doc.save(output / FILES[kind])
    mapping = {"files": files, "statistics": stats, "rows_per_page": ROWS_PER_PAGE,
               "full_pages": len(pages), "selected_pages": [i + 1 for i in indices],
               "submitted_rows": [pages[i] for i in indices]}
    write_json(work / "source-map.json", mapping)
    snapshot = dict(data)
    snapshot["project_root"] = str(project)
    snapshot["output_dir"] = str(output)
    snapshot["screenshot_evidence"] = images
    write_json(work / "build-input.json", snapshot)
    export = {"work_dir": str(work), "documents": []}
    for kind in docs:
        # Agreement and application are rendered to INTERNAL PDFs only for visual QA.
        pdf = output / (Path(FILES[kind]).stem + ".pdf") if kind in ["source", "manual"] else work / (kind + "-preview.pdf")
        export["documents"].append({"kind": kind, "docx": str(output / FILES[kind]), "pdf": str(pdf)})
    write_json(work / "export-job.json", export)
    write_json(work / "build-report.json", {"application_issues": issues,
               "rules": data.get("rules", {}), "expected_source_pages": len(indices),
               "expected_delivery_files": 6 if "agreement" in docs else 5,
               "status": "draft_pending_rule_content_and_visual_review"})
    return {"word_files": len(docs), "expected_source_pages": len(indices), "statistics": stats}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("init-templates")
    cmd = sub.add_parser("build")
    for argument in ["project", "data", "output", "work"]:
        cmd.add_argument("--" + argument, required=True)
    args = parser.parse_args()
    try:
        if args.command == "init-templates":
            initialize_templates()
            print("Templates created")
        else:
            print(json.dumps(build(args.project, args.data, args.output, args.work), ensure_ascii=False))
    except (ValueError, OSError, KeyError, UnicodeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
