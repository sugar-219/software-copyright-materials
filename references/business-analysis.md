# 功能资料、证据关联与可读草稿

新制作时使用。在原 input.json 中增加 business，不另建需要独立维护的业务数据库；草稿和检查报告均由同一输入生成。资料由 agent 阅读真实项目后填写，工具不代替业务判断或运行验证。

## 结构

以下仅演示字段结构，不是已核实功能，不可直接作为真实项目输入：

```json
{
  "business": {
    "positioning": "依据项目事实说明软件用途",
    "target_users": ["实际使用者"],
    "operation_flow": ["按实际业务填写主要流程"],
    "application_feature_ids": ["export"],
    "features": [{
      "id": "export",
      "name": "实际功能名称",
      "scenario": "用户在什么情形下需要此功能",
      "purpose": "此功能处理的具体事务",
      "entry": "真实页面入口、菜单或命令",
      "visible_elements": ["用户能看到的控件或命令行输出"],
      "steps": ["实际操作步骤"],
      "rules": ["实际输入限制、权限或异常处理；没有则留空数组"],
      "feedback": ["实际结果或反馈"],
      "evidence": [{"path": "src/main.py", "start_line": 10, "end_line": 30}],
      "pending": ["尚未核实的具体内容；没有则留空数组"]
    }]
  },
  "manual": {"sections": [{
    "heading": "实际章节名称",
    "feature_ids": ["export"],
    "blocks": [{"text": "按项目事实写出的自然段操作说明"}]
  }]}
}
```

- id 只用于内部稳定关联，名称按实际项目填写；不要用数组序号代替稳定 id。
- name、purpose、entry、steps、feedback 是功能资料必填项。页面控件、场景、限制、待核实内容按实际情况补充；没有限制不编造限制。CLI或后台任务可写真实命令、调用入口及可观察结果，不强制套页面结构。
- evidence 的 path 是项目内相对文件路径。可用 start_line/end_line 指定原始文本行，工具检查范围；也可用 location 指明函数、章节或 PDF 页码，文字定位仍需人工核对。文档编码不同可给 evidence.encoding。项目外资料先由 agent 直接阅读并将真实确认依据写入 fact_evidence；不为满足检查把附件塞进业务项目。
- application_feature_ids 表示申请信息主要功能文字实际涉及的功能。manual.sections[].feature_ids 表示该章节实际解释的功能；概述、安装等公共章节可以不关联。一章可涉及多项功能，多个章节也可分别说明同一项功能，按真实内容关联。
- 关联元数据不会自动写入 Word；申请信息仍使用 application.functions，说明书仍使用 manual.sections[].blocks 的自然段和图片，避免由脚本套句式编造正文。agent 必须对照这三处，不能只填 id 而不改实际文字。
- 截图从关联章节中的 image 块识别，不另建一份截图名单。暂不截图时在对应章节放可见的截图预留文字并保留待补事项；截图数量不证明对应关系真实。

不强制固定章节、FAQ、术语表、最低字数或检查轮数。阅读证据、定位问题、实际修改再复核；重复运行同一脚本不算独立业务核验。

## 可读草稿和覆盖检查

在写好统一输入后运行：

```text
<Python> -X utf8 -B <skill>/scripts/business_review.py --project <项目目录> --data <工作目录>/input.json --out-dir <工作目录>/preparation
```

生成：

- business-summary.md：用途、用户、流程、每项功能的操作资料和证据、材料覆盖表、结构错误及待核实事项。
- materials-draft.md：申请字段及说明书已有文字的可读预览，合作时附输入的合作信息。自动统计与规范化字段以生成器为准，协议最终文字及版式以既有模板为准。
- business-review.json：同一份结构化检查结果与证据文件哈希。

结构错误（如缺少必填步骤、重复或未定义 id、证据不存在、越界路径、行范围无效）会使命令返回非零，并阻止记录确认及 build/revise 写出材料。先修正输入再运行。

以下作为待人工处理的事实事项：申请信息功能未关联手册、有章节却没有操作正文、缺少证据或证据位置、缺少截图/保留截图预留，以及输入明确列出的 pending。它们允许保留草稿，不自动改为通过；最终检查会纳入原有 pending_review。

检查依据显式关联，不从字面关键词推断业务语义；即使关系齐全，也必须核对申请文字、章节、证据及截图实际讲的是同一功能。文件存在、行范围有效、哈希一致都不是功能真实性证明。字段限额继续以 assets/fields.json、当前核实页面及 field_limits 为准，不在这里新增一套永久限额。

## 确认与后续检查

先让用户审阅需要确认的可读草稿，沿用会话中已有的有效授权，不重复询问。按 preparation-and-confirmation.md 记录输入确认。新输入的 business 及其引用证据的实际字节会加入内容指纹；修改证据文档也会使记录失效。

生成器把 business-review.json 写入单次运行目录。inspect_materials.py 重新检查关联及证据哈希，证据在生成后改变会报错，遗漏项作为 fact 待办处理。检查报告不能替代最终 Word/PDF 与实际软件核验。

旧输入不含 business 时仍可制作、审查或修订，原有确认指纹兼容；报告明确提示无法自动核对覆盖关系，不虚构历史功能资料。新制作应补齐 business。局部修订仅增加或更正内部关联时可保留字节未变的材料，但必须重新核对这些关联是否真实；实际 application/manual 文字有变化时仍由原生成器重制相应材料。

## 文档摘要索引

prepare_project.py 现在为可处理的文档增加 excerpt、headings、extraction_status 和 excerpt_location。DOCX/ODT 用标准库读取，PDF 使用环境中已有的 pypdf；旧 DOC/WPS、提取失败、扫描件或超大文件保留待阅读状态。摘要通常只含提取文本开头1000字符，不覆盖全文；文档表格、复杂布局及后部条款需回看原文。

摘要放在内部工作目录，先核对是否含不宜传播的内容，不放入最终交付目录。索引方便选读，不自动决定行业、权属、功能或源码范围。

## 维护验证

运行 scripts/test_business_review.py --temp-root <独立测试目录>/business，并与原有两组测试一起验证。测试覆盖结构关联、证据变化失效、只读预览、摘要提取、旧确认兼容及生成/检查接入；不把测试资料或模拟记录当成真实项目检查依据。
