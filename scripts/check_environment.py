"""Read-only preflight; never install dependencies or launch Word."""
import argparse
import datetime
import importlib
import json
import os
import platform
import sys
from pathlib import Path


def check_environment():
    modules = {}
    for name in ["docx", "PIL", "pypdf", "pypdfium2"]:
        try:
            importlib.import_module(name)
            modules[name] = {"available": True}
        except Exception as exc:
            modules[name] = {"available": False, "error_type": type(exc).__name__}
    fonts_dir = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    fonts = {name: (fonts_dir / name).is_file() for name in ["simsun.ttc", "msyh.ttc", "consola.ttf"]}
    word = {"registered": False, "check": "registry_only", "automation_verified": False}
    if sys.platform == "win32":
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID") as key:
                word["registered"] = bool(winreg.QueryValueEx(key, "")[0])
        except OSError:
            pass
    templates_dir = Path(__file__).resolve().parents[1] / "assets" / "templates"
    templates = {name: (templates_dir / name).is_file() for name in [
        "软件说明书.docx", "源程序鉴别材料.docx", "申请信息填写稿.docx", "合作开发协议模板.docx"]}
    issues = []
    if sys.version_info < (3, 10):
        issues.append("需要 Python 3.10 或以上版本")
    if sys.platform != "win32":
        issues.append("完整生成与导出需要 Windows 和本机 Microsoft Word")
    if not all(item["available"] for item in modules.values()):
        issues.append("Python 依赖不完整；在独立环境安装 requirements.txt，不修改业务项目或全局 Python")
    if not word["registered"]:
        issues.append("未检测到 Word COM 注册；需核实本机 Microsoft Word")
    if not all(fonts.values()):
        issues.append("缺少约定字体文件")
    if not all(templates.values()):
        issues.append("随附模板不完整")
    return {"checked_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "python": {"executable": sys.executable, "version": platform.python_version()},
            "platform": sys.platform, "modules": modules, "fonts": fonts, "word": word,
            "templates": templates, "issues": issues, "preflight_passed": not issues,
            "limitations": ["注册表检查不等于 Word COM 可用；实际导出和分页仍需验证",
                            "预检失败时仍可整理证据和输入草稿，不宣称已生成完整材料"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out-dir", required=True)
    args = parser.parse_args()
    report = check_environment()
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    (out / "environment.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = ["# 环境检查", "", f"Python：{sys.executable}",
             f"预检：{'通过' if report['preflight_passed'] else '待处理'}", ""]
    for group in ["modules", "fonts", "templates"]:
        lines.append(f"{group}：{json.dumps(report[group], ensure_ascii=False)}")
    lines.extend(["", f"Word 注册：{report['word']['registered']}；尚未验证 COM 自动化", ""])
    lines.extend("- " + item for item in report["issues"] + report["limitations"])
    (out / "environment.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(json.dumps({"preflight_passed": report["preflight_passed"], "issues": report["issues"]}, ensure_ascii=False))
    return 0 if report["preflight_passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
