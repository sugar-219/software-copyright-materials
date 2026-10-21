# 软著申请材料制作 Skill

为已完成开发的软件准备中国软件著作权申请材料的 Codex skill。AI 根据真实项目核实功能、整理内容及官方要求，配套模板与脚本负责生成可编辑 Word、导出 PDF 和检查排版。

首版针对 **Windows + Microsoft Word**。使用时提供项目目录，由 agent 核对名称、版本、开发方式及缺失事实；个人信息可以留空，由申请人补充。

## 交付内容

| 材料 | 格式 | 条件 |
| --- | --- | --- |
| 软件说明书 | DOCX、PDF | 新制作时提供 |
| 源程序鉴别材料 | DOCX、PDF | 新制作时提供 |
| 申请信息填写稿 | 仅 DOCX | 新制作时提供 |
| 合作开发协议模板 | 仅 DOCX | 合作开发时生成 |

单独开发交付 5 个文件，合作开发交付 6 个文件。合作协议供填写、打印、签字后扫描，不交付空模板 PDF。截图、统计、检查记录和内部预览保存在工作目录。

## 安装与调用

将本仓库克隆或解压到个人 skills 目录，使入口文件位于：

```text
%USERPROFILE%\.agents\skills\software-copyright-materials\SKILL.md
```

也可以在个人 skills 目录执行：

```powershell
git clone https://github.com/sugar-219/software-copyright-materials.git
```

在项目聊天中调用：

```text
使用 $software-copyright-materials 为当前已完成的项目准备软著申请材料。
隐私由我填写，不确定的事项先与我确认。
```

若尚未发现新 skill，重新打开 Codex 后再检查。有关发现和调用机制，见 [OpenAI 官方 skills 文档](https://learn.chatgpt.com/docs/build-skills)。

## 工作流程

1. 预检 Python、依赖、字体和 Word 注册，检查项目源码、文档及实际功能。
2. 核对国家版权局及中国版权保护中心的适用规定和当前页面要求。
3. 整理统一底稿，记录功能操作资料、证据及申请信息和手册章节的对应关系。
4. 生成可读草稿和覆盖报告，核对缺项并依据真实用户确认绑定内容指纹。
5. 使用固定 Word 模板生成材料，从最终 Word 导出 PDF。
6. 检查源码映射、业务证据变化、分页、字体及交付文件，逐页查看渲染图后交付，并列明待办。

软件名称、实际申请版本和开发方式需要确认。开发完成日期、发表情况、权利归属及实际开发设备不能由源码或当前电脑配置推断。协议模板不预设分工、权利归属或份额。

支持新制作、既有材料审查及局部修订。已确认的合作方、日期和权利约定统一带入；没有简称时省略整栏，协议个人信息字段按需显示。局部修订使用 `revise`，仅重新生成受影响的文档；未改材料须通过文件哈希及基线记录核验后复用，名称、版本或源码范围变化时重新制作。

检查结果将未决事项区分为隐私待填、事实待确认、规则待核对和视觉检查，并保存证据及相关文件哈希。源码分页保留原文件映射，空行不用于凑足每页代码行数，不改写业务源码。

## 业务分析、草稿与内容确认

项目清单登记源码候选、文档及排除理由，可读取文档附开头摘要和提取状态。摘要仅用于选读，不能替代原文件和实际运行验证。业务判断由 agent 阅读项目后形成，不按文件名或关键词自动认定功能。

在同一输入中记录主要功能的用途、入口、步骤、反馈和证据，并用功能 id 关联申请信息和手册章节。工具生成：

- `business-summary.md`：业务理解、功能操作资料、证据和覆盖表。
- `materials-draft.md`：申请字段及说明书文字预览，自动统计和规范化结果以生成器为准。
- `business-review.json`：结构检查、覆盖缺口及业务证据文件哈希。

重复或未定义 id、缺少必要操作资料、无效证据路径或行范围会阻止生成；缺少章节关联、证据、正文或截图等事项进入待办。显式关联不证明文字语义一致或功能真实，仍需对照软件、源码及截图核验。

**生成前需要内容确认记录。** Agent 根据用户已确认的信息整理输入并运行 `confirm_input.py record`；仅设置 `software.confirmed=true` 不再足够。用户不用手写 JSON，也不用逐条执行这些内部脚本。源码、关键文字、截图或引用业务证据改变后，原记录失效，应核对变化及已有授权后更新，不能自动把失效记录改为通过。

旧输入不含 `business` 时保持兼容，并明确提示无法自动检查覆盖关系；已有材料不补造历史业务资料或确认。局部修订仅改变内部关联时可复用未变材料，实际文字改变时仍重制相应文档。

详细结构和命令见 [功能资料与草稿检查](references/business-analysis.md)及[准备与确认工具](references/preparation-and-confirmation.md)。

## 排版与运行条件

- A4 纵向，正文宋体，标题微软雅黑。
- 源程序英文及符号使用 Consolas，中文使用宋体。
- 本机安装 Word、宋体、微软雅黑和 Consolas；本仓库不分发这些字体。
- Python 3.10 及以上，使用能导入所需包的独立环境。

如需安装 Python 依赖，在选定的独立环境中执行：

```powershell
python -m pip install -r scripts/requirements.txt
```

Agent 应优先检查已有工作区运行时，避免修改业务项目或全局环境。Windows 命令行可使用 `python -X utf8` 避免中文输出编码问题。

## 文件结构

```text
SKILL.md                         执行入口与工作规则
agents/openai.yaml               显示名称及默认调用提示
assets/fields.json               申请字段、页面选项及长度提示
assets/templates/                四份 Word 模板
references/format-requirements.md 格式依据、来源与待核实细节
references/input-and-tools.md    agent 的制作输入及工具说明
references/preparation-and-confirmation.md 环境、证据与确认
references/business-analysis.md  功能资料、可读草稿与覆盖检查
scripts/check_environment.py     启动环境预检
scripts/prepare_project.py       项目证据、摘要及源码候选清单
scripts/business_review.py       业务草稿、材料预览及覆盖报告
scripts/confirm_input.py         内容指纹记录及校验
scripts/materials.py             生成或局部修订材料、提取源程序
scripts/export_word.ps1          本机 Word 导出与页面统计
scripts/inspect_materials.py     源码、字体、PDF和交付集合检查
scripts/test_materials.py        材料机械规则测试
scripts/test_preparation.py      准备工具及确认测试
scripts/test_business_review.py  业务覆盖与证据测试
scripts/requirements.txt         Python依赖
```

## 验证与当前边界

2026-10-15 在本机完成了 13 项测试，覆盖取页边界、长行与中文、原始和显示行数区分、源码映射、版本同步、合作输出、防覆盖及缺失输入。另用 Word 验证了小样张和 63 页源码样本取前后各 30 页后的 60 页输出，检查了 PDF 字体嵌入及可见内容。测试样张不是申报材料。

2026-10-17 更新后完成 28 项回归测试，新增可选字段、合作信息一致性、嵌套表格字体、源码正文完整性、页眉和检查证据校验。本机 Word 导出并逐页查看了两人、四人、八人协议样张；局部修订样例只重新导出协议，其余三类文档按哈希复用，并验证名称、源码或旧材料被改动时拒绝不当复用。测试中刻意保留的短说明书末页仍标记为规则待核对。

2026-10-21 增加环境预检、文档摘要索引、结构化功能资料、可读草稿、内容指纹及业务覆盖检查，完成 56 项测试和命令行联调。联调验证生成前接受有效确认，并在业务证据改变后拒绝沿用旧记录；保留旧输入兼容及局部复用。模板和 Word 导出逻辑未变，本次未重跑 Word 导出或逐页视觉检查，历史样张结果不作为本次重跑结果。

需要结合每个真实项目重新核实功能、截图、代码范围和申请事实。官方核验由执行 agent 完成，脚本不会仅凭一个网址证明已完成核验。

- 官网细则无法确认时，先完成可核实草稿，再请求相应提示或截图。
- 说明书超过 60 页时，当前脚本要求进一步处理，不自动裁切段落或图片。
- 部分特殊源码字符可能需要先确认字体显示方案。
- 自动行数和技术检查不能替代适用规则、例外情况、权属或登记结果的判断。

具体执行步骤见 [SKILL.md](SKILL.md)，格式来源见 [格式要求参考](references/format-requirements.md)。
