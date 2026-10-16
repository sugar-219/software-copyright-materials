# 制作输入与工具

本文件供执行 agent 使用。用户不需要手写 JSON；agent 根据项目证据及确认结果整理输入。

## 运行条件

- Windows，已安装可正常自动化的 Microsoft Word。
- 宋体 `simsun.ttc`、微软雅黑 `msyh.ttc`、Consolas `consola.ttf`。
- Python 3.10 及以上，安装 scripts/requirements.txt 中的包。
- 先探测现有运行时的模块；有合适的 Codex bundled Python 时直接使用。避免依赖未注册的 `py` 或默认 MSYS Python。
- Windows 命令行建议为 Python 加 `-X utf8`，避免中文控制台输出或校验器默认编码出现 GBK 解码错误；输入 JSON 仍明确使用 UTF-8。
- 遇到沙箱的路径解析或 Word COM 权限错误时，保留错误信息，按当前工具权限机制请求仅完成该本机操作所需的执行权限。不要为了绕过错误删除路径范围检查或改用修改原项目的办法。

## 输入示例

下面是结构示例，不是已核实的项目内容。`confirmed` 仅在用户已确认名称、版本和开发方式后设为 true；不能为了通过脚本检查自行推断。

```json
{
  "software": {
    "name": "用户确认的软件全称",
    "version": "2.3",
    "development_mode": "单独开发",
    "confirmed": false
  },
  "source_files": [
    {"path": "src/main.py", "encoding": "utf-8-sig"},
    {"path": "src/service.py", "encoding": "utf-8-sig"}
  ],
  "source_exclusions": [
    {"path": "vendor/", "reason": "第三方依赖，不计入本项目自主源码"}
  ],
  "application": {
    "rights_acquisition": "【待本人确认】",
    "category": "【待本人确认】",
    "original_or_modified": "【待本人确认】",
    "completion_date": "【待本人确认】",
    "publication_status": "【待本人确认】",
    "rights_scope": "【待本人确认】",
    "development_hardware": "【待本人确认】",
    "development_os": "【待本人确认】",
    "runtime_hardware": "根据软件真实要求填写",
    "development_tools": "根据项目和开发记录填写",
    "runtime_os": "根据已核实兼容环境填写",
    "runtime_dependencies": "根据实际运行依赖填写",
    "languages": "按实际项目填写",
    "purpose": "按实际项目填写",
    "industry": "按实际项目填写",
    "functions": "根据已经实现且核实的功能撰写完整说明",
    "technical_category": "根据实际情况选择页面类别",
    "technical_description": "根据实际技术撰写"
  },
  "manual": {
    "sections": [
      {
        "heading": "一 软件概述",
        "blocks": [{"text": "由agent依据真实项目撰写的正文。"}]
      },
      {
        "heading": "二 主要功能与操作",
        "blocks": [
          {"text": "实际操作入口、步骤和结果。"},
          {
            "image": "screenshots/actual-window.png",
            "origin": "runtime_capture",
            "caption": "图1 实际功能界面"
          }
        ]
      }
    ]
  },
  "cooperation": {"party_count": 2},
  "rules": {
    "checked_at": "本次真实核验日期",
    "sources": [
      {"url": "本次实际打开的官方页面地址", "sections": "条款或页面位置", "result": "实际核验结果"}
    ],
    "pending": ["无法核实的具体事项；没有则留空数组"]
  }
}
```

## 字段约定

- 所有申请字段及默认长度、选项均在 `assets/fields.json`，以其中的 key 对应 application。
- 软件全称、版本、开发方式从 software 自动同步到填写稿。源程序量自动使用 source_files 中原始非空行数，不采用折行后的显示行数。
- 实际源码范围必须先完整确定，不能只列一小部分代码后把该统计称作整个软件程序量。统计总行数、非空行数及范围保存在工作目录。
- 实际开发设备和操作系统不是当前运行脚本机器的配置；无法核实则待用户确认。
- 确认未发表后，首次发表日期和地点为不适用；已发表时由本人确认。
- 隐私默认 `【本人填写】`。owners 可按相同格式列出多位著作权人。
- 已核对页面长度与默认不同，输入 `field_limits` 覆盖对应字段的 min/max，并在 rules 记录依据。例如 `{"functions":{"min":400,"max":1200}}`。值只能为非负整数，不能通过设置过宽限制绕过真实要求。
- 图片路径相对 input.json 所在目录，或使用明确的绝对路径。origin 只接受 `runtime_capture` 或 `user_provided`；来源标记本身不能证明图片真实，agent仍需核实。
- 合作开发才读取 cooperation.party_count；未知人数可先生成默认两方、可增补的通用模板，或询问实际人数。权属、份额和签署信息始终由用户确认。
- 普通源程序使用严格解码，遇到其他编码在各源文件 encoding 中明确指定，不用替代字符吞掉乱码。

## 生成、导出与检查

1. 在工作根目录保存 input.json，另建空的单次运行子目录。交付目录和单次工作目录均须为空，且不能互相包含；截图和输入JSON放在单次运行目录之外。
2. 运行 materials.py build；它保留 source-map.json、build-input.json、build-report.json 和 export-job.json 于工作目录。
3. 使用 Windows PowerShell 运行 export_word.ps1 -Job export-job.json。脚本在独立 Word 实例中更新字段、保存最终DOCX并导出PDF，正常和异常结束均关闭其自身文档及实例，不结束用户其他 Word 进程。
4. 运行 inspect_materials.py --work 工作目录。检查结果为 inspection.json；每一页的PNG及联系表位于 renders/。
5. 必须核对 errors 和 pending_review，并逐页视觉检查；联系人隐私、官方例外规则等不能由脚本自动判定。

源代码每行先按字体宽度进行显示折行，再每55条显示行分一页。此数值是内部版式选择，不是官方要求；原始空行保留、制表符仅在显示时按4列展开，映射记录原始文件与行号。每页行数是否符合适用要求须结合实际Word统计与官方规则核实。

超过60页的源程序自动选取开头30页和末尾30页，最终稿页码连续，原始页到提交页的对应关系记录在工作目录。编号方式若与当前官方要求不同，需要调整模板及导出并重新验证。

说明书通常保持与软件功能相称的完整篇幅。当前工具不自动裁切超过60页的说明书，以免静默破坏图片或段落；出现该情形时会报错，由agent按核实规则处理后重新检查，不得当作完整通过。

源码范围验证、疑似私钥和常见凭据扫描只是辅助措施；需人工检查敏感配置、个人数据和第三方代码，不能据此宣称不存在其他敏感信息或许可问题。

## 修订与检查记录

- 导出脚本不覆盖已经存在的PDF；修订使用新的工作子目录和空暂存交付目录，核对后再按用户要求替换。
- 自动统计、图片、预览和检查记录留工作目录；最终仅交付约定的5或6个文件。
- 改动脚本后执行 test_materials.py --temp-root 独立测试目录；涉及排版、字体、导出时再做一次本机 Word 导出与视觉核对。
- 来源范围、内容和格式规则仍需结合每个真实项目确认，不能重复使用排版测试报告充当项目检查结果。
