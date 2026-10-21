"""Review declared feature coverage and produce readable drafts from the same input."""
import argparse
import hashlib
import json
from pathlib import Path


def review_business(project, data):
    report = {"enabled": "business" in data, "errors": [], "warnings": [],
              "features": [], "evidence_files": {},
              "limitation": "只检查显式关联、证据文件及结构；不证明功能已实现、文字语义一致或截图真实"}
    errors, warnings = report["errors"], report["warnings"]
    if not report["enabled"]:
        warnings.append("旧输入未提供 business 功能资料，无法自动核对功能与申请信息、手册和截图的覆盖关系")
        return report
    project = Path(project).resolve(strict=True)
    business = data["business"]
    if not isinstance(business, dict) or not isinstance(business.get("features"), list) or not business["features"]:
        errors.append("business.features 必须是非空功能列表")
        return report

    def texts(value, label, required=False):
        if not isinstance(value, list) or any(not isinstance(v, str) or not v.strip() for v in value):
            errors.append(label + " 必须是非空字符串组成的列表")
            return []
        if required and not value:
            errors.append(label + " 不能为空")
        return value

    for field in ["target_users", "operation_flow"]:
        if field in business:
            texts(business[field], "business." + field)
    features = {}
    for item in business["features"]:
        if not isinstance(item, dict) or not isinstance(item.get("id"), str) or not item["id"].strip():
            errors.append("每项功能必须有非空字符串 id")
            continue
        key = item["id"]
        if key in features:
            errors.append("功能 id 重复：" + key)
            continue
        features[key] = item
        for field in ["name", "purpose", "entry"]:
            if not isinstance(item.get(field), str) or not item[field].strip():
                errors.append(f"功能 {key} 缺少 {field}")
        for field in ["steps", "feedback"]:
            texts(item.get(field), f"功能 {key}.{field}", required=True)
        for field in ["visible_elements", "rules", "pending"]:
            if field in item:
                texts(item[field], f"功能 {key}.{field}")
        evidence = item.get("evidence", [])
        if not isinstance(evidence, list):
            errors.append(f"功能 {key}.evidence 必须是列表")
            continue
        for ref in evidence:
            if not isinstance(ref, dict) or not isinstance(ref.get("path"), str) or not ref["path"].strip():
                errors.append(f"功能 {key} 的证据必须提供 path")
                continue
            try:
                relative = Path(ref["path"])
                path = (project / relative).resolve(strict=True)
                if relative.is_absolute() or ".." in relative.parts or not path.is_relative_to(project):
                    raise ValueError("证据路径必须位于项目内")
                raw = path.read_bytes()
                report["evidence_files"][relative.as_posix()] = hashlib.sha256(raw).hexdigest()
                if "start_line" in ref or "end_line" in ref:
                    start, end = ref.get("start_line"), ref.get("end_line")
                    if type(start) is not int or type(end) is not int or not 1 <= start <= end:
                        raise ValueError("证据行范围无效")
                    if end > len(raw.decode(ref.get("encoding", "utf-8-sig"), errors="strict").splitlines()):
                        raise ValueError("证据行范围超过文件长度")
                elif not str(ref.get("location") or "").strip():
                    warnings.append(f"功能 {key} 的证据 {ref['path']} 未注明函数、章节或页码位置")
            except (ValueError, OSError, LookupError) as exc:
                errors.append(f"功能 {key} 的证据 {ref['path']} 无效：{exc}")

    app_ids = texts(business.get("application_feature_ids", []), "business.application_feature_ids")
    if len(app_ids) != len(set(app_ids)):
        errors.append("application_feature_ids 存在重复 id")
    if app_ids and not str(data.get("application", {}).get("functions") or "").strip():
        warnings.append("声明了申请信息功能关联，但 application.functions 尚无正文")
    sections_by_feature = {key: [] for key in features}
    for section in data.get("manual", {}).get("sections", []):
        section_ids = texts(section.get("feature_ids", []), "手册章节 feature_ids")
        if len(section_ids) != len(set(section_ids)):
            errors.append("手册章节 feature_ids 重复：" + str(section.get("heading", "")))
        for key in section_ids:
            if key not in features:
                errors.append("手册引用了未定义功能：" + key)
            else:
                sections_by_feature[key].append(section)
    for key in app_ids:
        if key not in features:
            errors.append("申请信息引用了未定义功能：" + key)

    for key, feature in features.items():
        sections = sections_by_feature[key]
        blocks = [block for section in sections for block in section.get("blocks", [])]
        images = [block["image"] for block in blocks if "image" in block]
        has_text = any(str(block.get("text") or "").strip() for block in blocks)
        placeholder = any("截图预留" in str(block.get("text", "")) for block in blocks)
        if key in app_ids and not sections:
            warnings.append(f"申请信息中的功能 {key} 未关联说明书章节")
        if sections and not has_text:
            warnings.append(f"功能 {key} 的说明书章节缺少操作正文")
        if not feature.get("evidence"):
            warnings.append(f"功能 {key} 缺少文件证据；须补充证据或记录真实的用户说明依据")
        if sections and not images:
            warnings.append(f"功能 {key} {'保留截图预留，实际截图待补' if placeholder else '缺少截图及可见截图预留'}")
        if key not in app_ids and not sections:
            warnings.append(f"功能 {key} 未被申请信息或手册引用，需核对申报范围")
        if isinstance(feature.get("pending", []), list):
            warnings.extend(f"功能 {key} 待核实：{item}" for item in feature.get("pending", []) if isinstance(item, str))
        report["features"].append({"id": key, "name": feature.get("name", key),
                                   "in_application": key in app_ids,
                                   "manual_sections": [section.get("heading", "") for section in sections],
                                   "screenshots": images, "screenshot_placeholder": placeholder,
                                   "evidence_count": len(feature.get("evidence", [])) if isinstance(feature.get("evidence", []), list) else 0})
    report["warnings"] = list(dict.fromkeys(warnings))
    return report


def require_business_review(project, data):
    report = review_business(project, data)
    if report["errors"]:
        raise ValueError("业务资料结构检查失败：" + "；".join(report["errors"]))
    return report


def business_pending(project, data, baseline=None):
    """Recheck current evidence; never treat an existing report as proof of reality."""
    report = review_business(project, data)
    errors = list(report["errors"])
    warnings = list(report["warnings"])
    if report["enabled"]:
        if baseline is None:
            warnings.append("缺少业务证据快照，需核对当前证据与材料的对应关系")
        elif baseline.get("evidence_files") != report["evidence_files"]:
            errors.append("业务证据文件在生成后发生变化，需重新核对业务口径及材料")
    return errors, warnings


def write_drafts(project, data_path, out_dir):
    data_path, out_dir = Path(data_path).resolve(strict=True), Path(out_dir).resolve()
    project = Path(project).resolve(strict=True)
    if out_dir == project or project.is_relative_to(out_dir):
        raise ValueError("草稿目录不能等于或包含项目根目录")
    data = json.loads(data_path.read_text(encoding="utf-8-sig"))
    report = review_business(project, data)
    software = data["software"]
    business = data.get("business") if isinstance(data.get("business"), dict) else {}
    users = business.get("target_users", [])
    users_text = "、".join(users) if isinstance(users, list) and all(isinstance(v, str) for v in users) else "待修正"
    lines = ["# 业务理解草稿", "", f"软件：{software['name']} {software['version']}", "",
             report["limitation"], "", "用途：" + str(business.get("positioning") or "待整理"),
             "用户：" + users_text, "", "## 主要流程", ""]
    flow = business.get("operation_flow", [])
    lines.extend(f"{i}. {step}" for i, step in enumerate(flow if isinstance(flow, list) else [], 1))
    draft_features = business.get("features", [])
    for item in draft_features if isinstance(draft_features, list) else []:
        if not isinstance(item, dict):
            continue
        lines.extend(["", "## " + str(item.get("name", "未命名功能")), ""])
        for field, label in [("scenario", "使用场景"), ("purpose", "用途"), ("entry", "入口"),
                             ("visible_elements", "可见内容"), ("steps", "操作"), ("rules", "限制及异常"),
                             ("feedback", "结果"), ("evidence", "证据"), ("pending", "待核实")]:
            value = item.get(field, "")
            if field == "evidence" and isinstance(value, list):
                value = "；".join(str(ref.get("path", "未指定文件")) + "（" +
                                  (f"第{ref['start_line']}至{ref.get('end_line', '?')}行" if "start_line" in ref
                                   else str(ref.get("location") or "位置待核对")) + "）"
                                  for ref in value if isinstance(ref, dict))
            elif isinstance(value, list):
                value = "；".join(json.dumps(v, ensure_ascii=False) if isinstance(v, dict) else str(v) for v in value)
            lines.append(f"{label}：{value or '未记录'}")
    lines.extend(["", "## 材料覆盖", "", "| 功能 | 申请信息 | 手册章节 | 截图数 |", "| --- | --- | --- | --- |"])
    for item in report["features"]:
        cells = [item["name"], "已关联" if item["in_application"] else "未关联",
                 "、".join(item["manual_sections"]) or "未关联", str(len(item["screenshots"]))]
        lines.append("| " + " | ".join(str(c).replace("|", "\\|").replace("\n", " ") for c in cells) + " |")
    lines.extend(["", "## 结构错误", "", *["- " + item for item in report["errors"]],
                  "", "## 待核实", "", *["- " + item for item in report["warnings"]]])
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "business-summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    draft = ["# 申请信息与说明书文字草稿", "", "来自同一输入；自动统计、字段规范化和协议版式以生成器为准。", "",
             f"软件：{software['name']} {software['version']}", "开发方式：" + software["development_mode"], "", "## 申请信息", ""]
    schema = json.loads((Path(__file__).resolve().parents[1] / "assets" / "fields.json").read_text(encoding="utf-8"))
    values = dict(data.get("application", {}))
    values.update(name=software["name"], version=software["version"], development_mode=software["development_mode"],
                  source_lines="生成时按确认范围自动统计")
    for group in schema:
        draft.extend(["### " + group["title"], ""])
        for item in group["fields"]:
            value = values.get(item["key"], item.get("default", "【待本人确认】"))
            if item.get("optional") and not str(value).strip():
                continue
            draft.extend(["**" + item["label"] + "**", "", str(value), ""])
    draft.extend(["## 说明书", ""])
    for section in data.get("manual", {}).get("sections", []):
        draft.extend(["### " + section["heading"], ""])
        for block in section.get("blocks", []):
            if "text" in block:
                draft.extend([block["text"], ""])
            elif "image" in block:
                path = Path(block["image"])
                path = path if path.is_absolute() else data_path.parent / path
                draft.extend([f"截图：{path.resolve().as_posix()}", str(block.get("caption", "")), ""])
    if data.get("cooperation"):
        draft.extend(["## 合作信息底稿", "", "以下为协议输入信息；正式协议由原有模板生成。", "",
                      "```json", json.dumps(data["cooperation"], ensure_ascii=False, indent=2), "```", ""])
    (out_dir / "materials-draft.md").write_text("\n".join(draft), encoding="utf-8")
    (out_dir / "business-review.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", required=True)
    parser.add_argument("--data", required=True)
    parser.add_argument("--out-dir", required=True)
    args = parser.parse_args()
    try:
        report = write_drafts(args.project, args.data, args.out_dir)
    except (ValueError, OSError, KeyError) as exc:
        parser.exit(1, str(exc) + "\n")
    print(json.dumps({"errors": report["errors"], "warnings": report["warnings"]}, ensure_ascii=False))
    return 2 if report["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
