"""Bind an existing user confirmation to input facts, source bytes and screenshots."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path

from business_review import require_business_review


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(",", ":")).encode("utf-8")).hexdigest()


def confirmation_path(data_path):
    path = Path(data_path)
    return path.with_name(path.stem + "-confirmation.json")


def input_groups(project, data_path):
    project, data_path = Path(project).resolve(strict=True), Path(data_path).resolve(strict=True)
    data = json.loads(data_path.read_text(encoding="utf-8-sig"))
    if data["software"].get("confirmed") is not True:
        raise ValueError("软件名称、版本和开发方式尚未确认")
    files = []
    for spec in data["source_files"]:
        relative = Path(spec["path"])
        path = (project / relative).resolve(strict=True)
        if relative.is_absolute() or ".." in relative.parts or not path.is_relative_to(project):
            raise ValueError("源码必须位于项目内：" + spec["path"])
        files.append({"spec": spec, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
    images = []
    for section in data.get("manual", {}).get("sections", []):
        for block in section.get("blocks", []):
            if "image" in block:
                path = Path(block["image"])
                path = path if path.is_absolute() else data_path.parent / path
                images.append({"image": block["image"], "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
    software = {k: v for k, v in data["software"].items() if k != "confirmed"}
    content = {key: data.get(key) for key in ["application", "manual", "cooperation", "field_limits", "fact_evidence"]}
    content["images"] = images
    if "business" in data:
        review = require_business_review(project, data)
        content.update(business=data["business"], business_evidence=review["evidence_files"])
    return {"identity": digest(software),
            "source": digest({"files": files, "exclusions": data.get("source_exclusions", []),
                              "layout": data.get("source_layout", {})}),
            "content": digest(content)}


def record_confirmation(project, data_path, note):
    if not note.strip():
        raise ValueError("须记录用户已确认的内容或既有确认依据")
    groups = input_groups(project, data_path)
    report = {"schema": 1, "project_root": str(Path(project).resolve(strict=True)),
              "confirmed_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
              "note": note, "groups": groups,
              "limitation": "仅记录agent依据的用户确认及内容变化；不证明真人操作、原创性、权属或登记通过"}
    path = confirmation_path(data_path)
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def verify_confirmation(project, data_path):
    path = confirmation_path(data_path)
    if not path.is_file():
        raise ValueError("缺少内容确认记录；先展示关键事实、源码范围及草稿，依据用户确认运行 confirm_input.py")
    report = json.loads(path.read_text(encoding="utf-8-sig"))
    if report.get("schema") != 1 or not report.get("note", "").strip():
        raise ValueError("内容确认记录无效")
    if report.get("project_root") != str(Path(project).resolve(strict=True)):
        raise ValueError("确认记录属于其他项目，需核对项目后重新确认")
    current = input_groups(project, data_path)
    changed = [key for key, value in current.items() if report.get("groups", {}).get(key) != value]
    if changed:
        labels = {"identity": "名称/版本/开发方式", "source": "源码内容/范围/顺序/排版", "content": "申请事实/说明书/协议/截图/业务资料及证据"}
        raise ValueError("内容确认已失效：" + "、".join(labels[key] for key in changed) + "；展示变化并取得相应确认后更新记录")
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=["record", "check"])
    parser.add_argument("--project", required=True)
    parser.add_argument("--data", required=True)
    parser.add_argument("--note")
    args = parser.parse_args()
    try:
        if args.command == "record":
            if not args.note:
                parser.error("record 需要 --note，不能自动把未确认内容记为通过")
            print(record_confirmation(args.project, args.data, args.note))
        else:
            verify_confirmation(args.project, args.data)
            print("内容指纹一致；不替代事实与材料核验")
    except (ValueError, OSError, KeyError) as exc:
        parser.exit(1, str(exc) + "\n")


if __name__ == "__main__":
    main()
