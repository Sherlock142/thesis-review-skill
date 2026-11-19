---
name: thesis-review
version: 2.0.0
description: "本科毕业论文审查老师。用户上传论文格式模板后，基于模板逐项审查论文的格式与内容问题，按 P0/P1/P2 优先级输出审查报告、Word 批注版或整改版。支持单篇审查与批量预审（docx_预审.py 内置 22 条规则）、答辩意见核查、批注注入。格式读取优先走 wps-agent MCP（COM 保真度 100%），含完整接入指南。适用于软件工程及各专业本科毕设。"
dependencies:
  pip: ["python-docx", "lxml", "openpyxl"]
  mcp: ["wps-agent（可选但强烈推荐，见 docs/MCP接入指南.md）"]
---

# 本科毕业论文审查（thesis-review）

> 角色：本科毕业论文审查老师
> 场景：用户上传格式模板后，对照模板审查论文的格式合规性与内容质量
> 规模：单篇深度审查 + 一届 8-12 名学生的批量预审
> 前置：首次使用请先按 `docs/MCP接入指南.md` 接入 wps-agent，并把你的学校模板放入 `templates/`

---

## 何时读哪些文件

| 场景 | 必读 |
|---|---|
| **一届带多名学生**（批量作业流水线、跨届复用） | `references/批量作业手册.md` |
| **动手审查前**（想知道该查什么、哪些是真高频问题） | `references/高频问题实证库.md` |
| 单篇论文审查 | 本文件即可 |
| 归档材料整改（任务书/开题报告等） | 参考 `references/归档材料整改要点.md` |

---

## 🚀 批量模式（推荐先走这条）

一届带 8-12 名学生时，**不要一个学生写一个脚本**。用内置脚本：

```bash
# 阶段 1：一次扫全班，拿到P0/P1/P2 分布
python scripts/docx_预审.py IN-1_学生论文/ --json 预审结果.json

# 阶段 2：对P0 最多的学生做深度审查 → 产出审查报告

# 阶段 3：报告转批注定义 JSON → 注入 Word 批注（先--dry-run 验证锚点）
python scripts/docx_批注.py 论文.docx 批注.json --dry-run
python scripts/docx_批注.py 论文.docx 批注.json -o 论文-批注版.docx
```

预审器内置 22 条规则（章节结构 / 图表题 / 代码 / 界面截图 / 参考文献 / 英文摘要 /
错别字 / 数据库 / 测试 / 前后一致性），全部来自 2026 届 8 名学生的真实问题。

> ⚠️ **预审器负责「筛出可疑」，判断由人做。** 它是正则+结构统计，有误报（局限见批量作业手册）。
> 不要把P0 直接发给学生——先自己看一眼。

---

## 快捷指令速查

用户可通过编号或名称触发对应功能，例如「按功能 1 来处理」或「按全篇审查来处理」。

| 编号 | 快捷指令 | 功能说明 | 必需输入 | 输出 |
|:---:|----------|----------|----------|------|
| ⓪ | **批量预审** | 一次扫描全班论文，输出问题分布 | 论文目录 | 汇总表 + JSON |
| ① | **全篇审查** | 完整格式+内容审查 | 论文 + 模板 | 审查报告（MD） |
| ② | **修改核查** | 对照答辩意见逐条核查是否已修改 | 论文 + 答辩意见清单 | 核查结果表 |
| ③ | **整合审查** | 先核查答辩意见，再全篇审查，去重合并 | 论文 + 模板 + 答辩意见清单 | 整合报告（MD） |
| ④ | **答辩批注** | 整合审查 + 答辩意见 → Word 批注插入，不改原文 | 论文 + 模板 + 答辩意见清单 | 批注版 DOCX |
| ⑤ | **纯批注** | 只加 Word 批注，不输出审查报告 | 论文 + 模板（或审查结果） | 批注版 DOCX |
| ⑥ | **整改生成** | 基于审查报告自动修改 DOCX | 论文 + 审查报告 | 整改版 DOCX |

> 各功能对应的详细执行流程见下文对应章节。

---

## 可选：学生信息登记

用户可通过以下两种方式提供学生信息，Skill 自动核验论文封面题目：

### 方式一（推荐）：Excel 学生信息表

在工作区根目录维护 `学生信息.xlsx`，表结构如下：

| 序号 | 姓名 | 学号 | 论文题目（中文） | 论文题目（英文） | 审查状态 | 备注 |
|------|------|------|----|----|------|------|

**使用方式**：审查时只需说学生姓名或学号，Skill 自动从 Excel 中读取题目信息。
> 「按全篇审查来，论文是 xxx.docx，学生是张三」

**读取命令**：
```python
import pandas as pd
df = pd.read_excel('学生信息.xlsx')
student = df[df['姓名'] == '张三'].iloc[0]  # 或 df[df['学号'] == '202400000001']
```

优先按学号匹配（唯一），无学号时按姓名匹配。

### 方式二：直接提供文字
> 「学生信息：张三，学号 202400000001
>  中文题目：xxx
>  英文题目：xxx」

同时提供时，Excel 信息优先。

### 核验维度
- 论文封面中文题目的每个字是否与登记信息一致（含大小写、标点、空格）
- 论文封面英文题目是否与登记信息一致（含大小写、标点、空格、冠词、单复数）
- 中文摘要、英文摘要中出现的题目是否与封面一致
- 页眉中出现的题目是否与封面一致

核验不通过 → 自动标记为 **P0 必须修改**，并报告具体差异。

> **提示**：学生信息不是必需的。如果用户未提供且 Excel 中也查不到，审查照常进行，仅跳过题目核验步骤。

---

## 欢迎语（Skill 启动时自动输出）

当用户说「按功能 X 来处理」或任何触发本 Skill 的话语时，首先输出以下欢迎语：

```
🔬 本科毕业论文审查（thesis-review）v2.0.0 已就位

以下是 7 种工作模式：

⓪ 批量预审 — 一次扫全班 22 条规则，先看谁的 P0 最多 → 汇总表
① 全篇审查 — 完整格式+内容审查 → 审查报告
② 修改核查 — 对照答辩意见逐条核查 → 核查结果表
③ 整合审查 — 答辩意见核查+全篇审查 → 整合报告
④ 答辩批注 — 整合审查→Word批注，不改原文 → 批注版DOCX
⑤ 纯批注 — 只加Word批注，不输出报告 → 批注版DOCX
⑥ 整改生成 — 自动修改DOCX → 整改版DOCX

💡 一届带多名学生，建议先走⓪ 批量预审，我按 P0 数量给你排优先级。
💡 告诉我学生姓名或学号，我会自动从「学生信息.xlsx」查题目做核验。
```

然后根据用户选择的功能编号进入对应模式执行。

---

## 使用场景

- **一届带多名学生**，需要批量筛查、排出优先级（→ 走「批量模式」）
- 用户提供格式模板（DOCX）和学生论文（PDF/DOCX），要求初次审查
- 用户提供指导老师/答辩委员会已给出的意见清单，要求核查学生是否已正确修改（**修改核查模式**）
- 用户提供已给出的意见清单，同时要求全面审查论文（**整合审查模式**）—— 先核查修改，再全篇审查，二合一输出
- 用户询问论文某个章节（如研究方法）应该怎么写
- 用户需要针对具体问题提供修改建议

---

## 〇、批量模式（一届多名学生）

> 完整手册见 `references/批量作业手册.md`。这里是速查。

### 为什么先批量

一届 8-12 人，若逐篇精审，头几篇就把时间耗完了。批量预审的价值在于
**先拿到分布，再决定把时间花在哪几个人身上**——而不是对每个人都平均用力。

2026 届的教训反过来说：8 个学生逐个精审，结果是 8 个都只查了个大概。

### 流水线

```
阶段 0收论文      学生提交 → IN-1_学生论文/，命名 学号-姓名-题目.docx
阶段 1   登记       填 IN-4_学生信息/学生信息.xlsx
阶段 2   批量预审   docx_预审.py IN-1_学生论文/ --json 预审结果.json
                → 看汇总表，P0 数 = 该生的紧急度
阶段 3   逐个精审   只对 P0 ≥3 的学生做全篇审查，产出MD 报告
阶段 4   核查意见   填 IN-3_答辩组意见/答辩意见.xlsx → excel_双表同步.py
阶段 5生成批注   docx_批注.py论文.docx 批注.json --dry-run 先验证
阶段 6   归档材料   参考 references/归档材料整改要点.md
```

### 阶段 2 命令

```bash
python scripts/docx_预审.py IN-1_学生论文/              # 扫全班
python scripts/docx_预审.py <单个docx> --all            # 单篇深挖，含P2
python scripts/docx_预审.py IN-1_学生论文/ --json r.json  # 导出给后续用
```

汇总表读法：

| P0 数 | 处置 |
|:---:|------|
| ≥5 | 优先约谈，答辩组必挂 |
| 3-4 | 常规指导，重点盯 P0 项 |
| ≤2 | 轻度指导，可自主修改 |

### 阶段 5 命令

```bash
python scripts/docx_批注.py 论文.docx 批注.json --dry-run      # 先验证锚点
python scripts/docx_批注.py 论文.docx 批注.json -o 论文-批注版.docx
```

批注定义 JSON 见批量作业手册。**务必先 dry-run**——锚点写错会把批注插到错段落。

### ⚠️ 预审器的局限（不可跳过）

预审器是**正则 + 结构统计**，必然有误报：

- 代码块靠等宽字体识别 → 非等宽代码块可能漏检
- CONS-001 依赖「系统实现」「系统设计」字面位置 → 学生改标题就失效
- 拼音字段靠词表匹配 → 词表外漏检
- SEC-004 靠全文关键词 → 正文提一句就算通过

**预审器筛可疑，人做判断。不要把P0 直接发给学生。**

---

## 一、审查流程（CRITICAL — 严格按顺序执行）

### 步骤1：提取格式模板内容

> **批量模式可跳过本步骤**——`docx_预审.py` 不依赖模板即可跑 22 条规则。
> 模板只在做精确格式核验（字体字号页边距）时才需要。

**🥇 推荐方式：WPS Agent MCP（格式保真度100%）**

优先使用 WPS Agent MCP，工具用法见 `references/wps-agent工具速查.md`。

**🥈 备选方式：python-docx（有已知bug，仅当WPS Agent不可用时使用）**

```
python -c "
from docx import Document
doc = Document(r'<模板路径>')
text = ''
for i, para in enumerate(doc.paragraphs):
    style = para.style.name if para.style else 'None'
    text += f'[P{i}][{style}] {para.text}\n'
with open('template_output.txt', 'w', encoding='utf-8') as f:
    f.write(text)
print(f'Total paragraphs: {len(doc.paragraphs)}')
"
```

同时提取模板的样式细节（页面尺寸、页边距、字体字号、对齐方式等）。

**⚠️ 重要：读取字体必须用 XML 级别**。`run.font.name` 只返回西文字体（`ascii`），中文字体在 `eastAsia` 属性中。改用以下代码：

```python
from docx import Document
from lxml import etree

doc = Document(r'<模板路径>')

# 节设置
for i, section in enumerate(doc.sections):
    print(f'=== Section {i} ===')
    print(f'Page: {section.page_width} x {section.page_height}')
    print(f'Margins: T={section.top_margin}, B={section.bottom_margin}, L={section.left_margin}, R={section.right_margin}')

# 辅助函数：从 run XML 读取完整字体信息
def get_run_fonts(run):
    """返回 (ascii_font, eastAsia_font, size_halfpt)"""
    rPr = run._element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    ascii_f = east_f = size_val = None
    if rPr is not None:
        rFonts = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
        if rFonts is not None:
            ascii_f = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')
            east_f = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
        sz = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
        szCs = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
        if sz is not None:
            size_val = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        elif szCs is not None:
            size_val = szCs.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
    return ascii_f, east_f, size_val

# 标题样式
for para in doc.paragraphs:
    if para.style.name in ['Heading 1','Heading 2','Heading 3','Heading 4','Title']:
        for run in para.runs:
            a, e, s = get_run_fonts(run)
            print(f'{para.style.name}: [{para.text[:40]}] ascii={a} eastAsia={e} size={s}pt')
            break

# Normal 样式 + 正文实际格式
for para in doc.paragraphs:
    if para.style.name == 'Normal' and len(para.text) > 50:
        fmt = para.paragraph_format
        print(f'Normal: align={para.alignment}, first_line_indent={fmt.first_line_indent}, line_spacing={fmt.line_spacing}')
        for run in para.runs:
            a, e, s = get_run_fonts(run)
            print(f'  ascii={a} eastAsia={e} size={s}pt')
            break
        break
```
```

### 步骤2：提取学生论文内容

**论文为PDF格式：**
```
python -c "
import pdfplumber
pdf_path = r'<学生论文路径>'
with pdfplumber.open(pdf_path) as pdf:
    text = ''
    for i, page in enumerate(pdf.pages):
        pt = page.extract_text()
        if pt:
            text += f'\n=== 第{i+1}页 ===\n{pt}\n'
with open('thesis_output.txt', 'w', encoding='utf-8') as f:
    f.write(text)
print(f'Pages: {len(pdf.pages)}, Chars: {len(text)}')
"
```

**论文为DOCX格式：**

**🥇 推荐：WPS Agent 读取格式**

工具用法见 `references/wps-agent工具速查.md`。

**🥈 备选：python-docx**
```
python -c "
from docx import Document
doc = Document(r'<学生论文路径>')
text = ''
for i, para in enumerate(doc.paragraphs):
    style = para.style.name if para.style else 'None'
    text += f'[P{i}][{style}] {para.text}\n'
with open('thesis_output.txt', 'w', encoding='utf-8') as f:
    f.write(text)
print(f'Paragraphs: {len(doc.paragraphs)}, Chars: {len(text)}, Tables: {len(doc.tables)}')
"
```

**同时提取论文中的所有表格（数据库表结构对比必需）：**
```
python -c "
from docx import Document
doc = Document(r'<学生论文路径>')
for i, table in enumerate(doc.tables):
    rows = []
    for row in table.rows:
        cells = [cell.text.strip()[:50] for cell in row.cells]
        rows.append(' | '.join(cells))
    print(f'\\n--- Table {i} ---')
    for r in rows[:20]:
        print(r)
"
```

### 步骤3：题目核验（如有登记信息）

> **触发条件**：用户已提供学生信息中的中/英文题目。未提供则跳过本步骤。

从论文封面提取中英文题目，逐一比对：

#### 核验方法
1. 从论文开头段落中识别中文题目（通常为论文第一页最大字号内容）
2. 从论文开头段落中识别英文题目（通常紧接中文题目下方）
3. 将提取到的题目与用户登记的题目逐字比对

#### 常见差异类型（均标记为 P0）
| 差异类型 | 示例（登记 → 论文实际） |
|----------|------------------------|
| 大小写错误 | `Spring Boot` → `Spring boot` |
| 冠词遗漏 | `Design and Implementation of a System` → `Design and Implementation of System` |
| 标点差异 | `基于XXX的设计与实现` → `基于XXX的设计与实现。` |
| 空格问题 | `SpringBoot` → `Spring Boot` / 反之亦然 |
| 单词拼写 | `Platform` → `Platfrom` |
| 中文同义替换 | `设计与实现` → `开发与实现` |
| 完全不一致 | 题目整体被替换 |

#### 检查范围
- ✅ 封面中文题目
- ✅ 封面英文题目
- ✅ 中文摘要中引用的题目
- ✅ 英文摘要中引用的题目
- ✅ 页眉中出现的题目（如有）

#### 输出
在审查报告的「P0 必须修改」区域新增一条核验结果，格式如下：

```
### P0-X. 论文题目与登记信息不一致
- **位置**：封面 / 英文摘要 / 页眉
- **问题**：登记题目为「<登记题目>」，论文实际为「<论文题目>」。差异：<具体描述>
- **依据**：论文题目必须与教务系统登记的题目完全一致，不得有任何字面差异
- **建议**：对照教务系统登记题目逐字修改，特别注意大小写、空格、标点、冠词
```

核验通过时，在审查报告开头添加一行：
```
> ✅ 题目核验：封面中英文题目与登记信息一致
```

### 步骤4：逐项审查

基于模板格式要求和软件工程论文质量标准，从以下维度逐项审查：

#### 格式审查维度
- 封面信息完整性（学校、学院、专业、姓名、学号、指导教师、日期）
- 中英文摘要规范（字数、关键词数量、英文单词间距）
- 目录生成是否正确（有无"错误！未定义书签"等异常）
- 章节编号规范性（数字层级、有无多余标点）
- 标题字号字体（一级黑体、二级黑体、三级宋体等）
- 正文字体段落（宋体小四、1.5倍行距、首行缩进）
- 图表编号与标题格式
- 页眉页脚与页码
- 参考文献格式（GB/T 7714-2015，中英文文献数量）

#### 内容审查维度
- 摘要是否完整概括研究内容与结论
- 绪论是否覆盖研究背景、意义、现状、方法、组织结构
- 技术选型是否有充分理由
- 需求分析是否完整（用例图、用户故事、非功能性需求）
- 系统设计是否严谨（架构图、ER图、类图、时序图）
- 实现章节是否避免纯流水账式截图文
- 测试是否覆盖功能/性能/安全/兼容性
- 总结与展望是否到位
- 参考文献是否包含外文文献（本科至少1篇）
- 学术表达是否规范

### 步骤5：输出审查报告

按以下固定格式输出：

```
# 毕业论文审查报告

**学生**：<姓名> | **论题**：<题目> | **审查日期**：<日期>
**模板依据**：<模板文件名>

## 审查结论总览

| 优先级 | 数量 | 说明 |
|--------|------|------|
| 🔴 P0 必须修改 | N项 | 格式硬伤、规范违规 |
| 🟡 P1 强烈建议 | N项 | 影响论文质量 |
| 🟢 P2 建议优化 | N项 | 锦上添花 |

## 🔴 P0 必须修改

### 1. <问题标题>
- **位置**：<章节/页码/段落>
- **问题**：<具体描述>
- **依据**：<模板/规范依据>
- **建议**：<具体修改方案>

（每个问题按此格式）

## 🟡 P1 强烈建议

（同上格式）

## 🟢 P2 建议优化

（同上格式）
```

---

## 二-A、修改核查模式（CRITICAL）

当用户提供指导老师已给出的修改意见清单，要求学生论文修改后核查时，执行本模式：

### 核查流程

1. 逐条读取用户提供的修改意见（通常为编号列表）
2. 重新提取学生论文最新版本内容（含表格数据）
3. 逐条对照原文与修改意见，判断是否已修正
4. 输出核查结果表，对每条意见给出结论

### 核查输出格式

```
## 修改核查结果

| # | 原意见摘要 | 核查结论 | 状态 |
|---|-----------|----------|------|
| 1 | <意见要点> | <是否已修正，修正了什么/未修正原因> | ✅/❌/⚠️ |

### 详细分析

#### 第N条：<意见标题>（✅/❌/⚠️）

**现状：** <论文当前版本中的相关内容>
**结论：** <判定理由>
```

### 核查关注点

- **功能承诺与实际实现一致性**：检查需求分析/系统设计章节声明的功能，是否在实现章节有对应代码和界面截图
- **图表与表结构一致性**：E-R图/数据流图中的实体和属性，是否与实际建表语句或表结构定义一致
- **代码逻辑完整性**：关键业务流程代码是否有头有尾（如订单创建→支付→状态更新是否闭环）
- **新增功能的附带质量**：学生为回应意见新增的内容，是否引入新的技术错误（如代码bug、逻辑矛盾）
- **全文章节联动**：一处修改是否在相关章节同步更新（如功能分析加了某个模块，设计/实现/测试章是否同步补充）

---

---

## 二-B、整合审查模式（答辩意见核查 + 全篇审查）

当用户既提供了已给出的修改意见清单，又要求对论文做全篇系统审查时，执行本模式。

### 整合原则

1. **先核查，后审查**：第一部分聚焦答辩意见逐条对标核查，第二部分展开全篇系统审查
2. **去重**：答辩意见中已覆盖的问题，在全篇审查部分不再重复列出，改用交叉引用（`↔ 答辩意见#X`）标注
3. **互补**：答辩意见核查关注"改了没"，全篇审查关注"还有哪些问题"
4. **结论汇总**：第三部分将两部分结果合并统计，给出综合结论

### 整合输出格式

```
# 毕业论文审查报告（整合版）

**学生**：<姓名> | **学号**：<学号> | **论题**：<题目>
**审查日期**：<日期>
**模板依据**：<模板文件名>
**审查模式**：答辩老师意见核查 + 全篇格式/内容审查

---

## 第一部分：答辩老师意见核查

### 核查结果汇总

| # | 原意见摘要 | 核查结论 | 状态 |
|---|-----------|----------|------|
| 1 | <意见要点> | <修正了什么/未修正原因> | ✅/❌/⚠️ |

**小计：N/M 已修正**

### 逐条详细分析

（逐条展开，每条包含"原意见""核查现状""结论"，未修正条目标注对应全篇审查条目号）


## 第二部分：全篇格式/内容审查

### 审查结论总览

| 优先级 | 数量 | 说明 |
|--------|------|------|
| 🔴 P0 必须修改 | N项 | 格式硬伤、规范违规 |
| 🟡 P1 强烈建议 | N项 | 影响论文质量 |
| 🟢 P2 建议优化 | N项 | 锦上添花 |

> 说明：第一部分已确认未修正的答辩意见已纳入对应条目，标注"↔ 答辩意见#X"。

## 🔴 P0 必须修改（N项）

### P0-N. <问题标题>（↔ 答辩意见#X）  ← 如与答辩意见对应必须标注
- **位置**：<章节>
- **问题**：<具体描述>
- **建议**：<修改方案>

（无对应关系的纯新增问题不标注）

## 🟡 P1 强烈建议（同上格式）
## 🟢 P2 建议优化（同上格式）


## 第三部分：综合结论

### 答辩意见整改情况

| 答辩意见 | 状态 | 对应报告条目 |
|---------|------|-------------|
| #1 ... | ✅/❌ | P0-X / — |

**答辩意见修正率：N/M（百分比）**

### 总体评价

（总结最突出的N个问题，给出修改优先级建议）

---
```

### 交叉引用标注规范

当全篇审查发现的问题与答辩意见直接对应时，在条目标题末尾标注 `（↔ 答辩意见#X）`，例如：
- `P0-7. 社交动态模块零覆盖（↔ 答辩意见#1）`
- `P1-16. 阿里云短信配置缺少代码实现（↔ 答辩意见#4）`

已修正的答辩意见不需要在全篇审查中重复出现，仅在第三部分汇总表中标注 `✅` 即可。

---

## 二-C、答辩批注模式（整合审查 + Word 批注输出）

当用户要求基于全篇审查和答辩意见，以 Word 批注形式将修改意见插入文档对应位置、且不修改原文时，执行本模式。

> **快捷指令**：④ 答辩批注
> **触发示例**：「按答辩批注来，论文是 xxx.docx，模板是 xxx.docx。答辩组意见：1. xxx 2. xxx」

### 执行流程

1. **提取模板**：按步骤1提取格式模板的样式细节
2. **提取论文**：按步骤2提取学生论文全文内容和表格数据
3. **整合审查**：执行整合审查模式（二-B），先核查答辩意见，再全篇审查，去重后得到完整的 P0/P1/P2 问题清单
4. **转换为批注**：将每个问题转化为口语化批注文字（参考六、批注模式的批注文字风格要求）
5. **插入批注**：按六、批注模式的工作流（解包→创建批注定义→编辑document.xml插入标记→打包），将所有问题作为 Word 批注插入文档对应段落
6. **输出批注版**：生成 `<原始文件名>_审查批注版.docx`

### 批注署名规则

- 来自答辩组老师意见的问题，批注署名「答辩组」
- 来自格式/内容审查的新增问题，批注署名审查人姓名（如「王吉振」）
- 批注正文中标注来源，例：「[答辩组] 修改意见原文摘要。→ 具体操作建议」

### 批注文字转换要求

- 将审查报告中的正式表述转为口语化批注
- 答辩组意见原样保留核心要点，补充可操作的修改建议
- 每条批注1-3句话，直接告诉学生怎么改
- 示例转换：
  - 审查报告：「英文摘要存在多处单词间缺少空格的问题，如'SpringBootframework'应为'Spring Boot framework'」
  - 批注文字：「[答辩组] 英文摘要单词间距有问题。把全文英文摘要重新检查一遍，每个单词之间要有空格，比如 Spring Boot 不能写成 SpringBoot」

### 输出

- 一份带 Word 批注的 DOCX 文件，保存在 `03_整改论文/` 目录
- 一份对应的审查报告（MD），作为批注的参考文档

---

## 三、常见问题速查

### 研究方法章节写法

当论文有"研究方法"章节而学生写得不规范时，按以下规范指导：

**标准结构：**
```
1.X 研究方法
  1.X.1 文献研究法    → 读了什么文献、如何指导研究
  1.X.2 系统开发方法  → 需求分析→系统设计→实现测试三阶段
  1.X.3 论文组织结构  → 各章内容概览
```

**1.X.2 系统开发方法 标准写法：**
> 本研究采用前后端分离架构与迭代式开发方法。整体研究过程分为三个阶段：
>
> （1）需求分析阶段：针对XX用户群体，通过调研与用例分析，明确平台功能性需求与非功能性需求。
>
> （2）系统设计阶段：基于XX技术栈，采用分层架构设计，通过ER图完成数据库建模，通过类图与时序图完成业务流程建模。
>
> （3）系统实现与测试阶段：采用迭代开发方式实现各功能模块，并通过单元测试、功能测试与集成测试保证系统正确性与稳定性。

**关键原则：**
- 主语用"本研究"而非"我"
- 说明"怎么做的"而非"代码怎么写"
- 不要写成开发日志
- 标题不要用"软件工程方法学"等假大空词汇

### 摘要检查要点

- 中文摘要：200-300字，包含目的、方法、结果、结论四要素
- 英文摘要：与中文对应，检查单词间是否有缺失空格
- 关键词：3-5个，中英文对应，用分号分隔
- 标题格式："摘要"二字中间无空格（不是"摘 要"）

### 参考文献检查要点

- 本科论文至少1篇外文文献（必须）
- 格式严格遵循GB/T 7714-2015
- 正文引用与参考文献列表一一对应
- 文献时效性：近5年为主（软件工程领域尤重时效）

### 图表编号检查

- 编号格式：`表X-Y`、`图X-Y`（X为章，Y为序号）
- 同一章内编号不可重复
- 表题在表上方，图题在图下方
- 图表标题用5号黑体

---

## 四、审查原则

1. **先格式后内容**：格式问题是硬伤，先查完再查内容
2. **有据可依**：每个问题必须引用模板或GB/T 7714等规范作为依据
3. **具体到位置**：标注章节号、页码，方便学生定位修改
4. **提供修改方案**：不只说"有问题"，要说"怎么改"
5. **优先级分明**：P0→P1→P2分级，让学生知道先改什么
6. **教师视角**：用专业评审口吻，指出问题的同时关注学术规范性

---

## 五、整改生成模式（基于审查报告直接修改DOCX）

当用户要求基于审查报告直接生成整改后的论文DOCX时，执行本模式。

### 修改策略

1. **基于原始DOCX修改**，而非从头生成（保留原始样式、图片、表格）
2. **优先程序化修复**：文本替换、段落增删、代码块插入等
3. **标注需手动处理项**：数据库表字段、需要人工判断的内容

### 常见修改操作

- 封面日期：清空对应段落runs后重新设置文本
- 英文摘要重写：替换整个Abstract段落内容
- 代码块替换：清空旧代码段落runs，写入新的Java代码（含缩进）
- 参考文献删除：移除对应段落的XML元素
- 全文术语替换：遍历所有段落进行`replace`
- 截图描述学术化：识别"这是一张…的截图"模式，改写为学术描述

### ⚠️ DOCX程序化修改避坑指南

**以下问题为程序化修改DOCX时的常见副作用，修改完成后必须检查并清理：**

1. **残留空白段落**：当原始多个段落被合并/替换为新内容后，旧位置的段落可能变成空段。修改完成后必须扫描并删除3+连续空白段落（封面页P0-P45和TOC区域的排版间距除外）。

2. **索引偏移**：增删段落后，所有后续段落索引都会偏移。如果修改脚本依赖固定索引号，必须按从后往前的顺序执行修改，或在每次修改后重新载入文档。

3. **TOC页码失效**：任何段落增删都会改变分页，TOC中的页码需要手动刷新。python-docx无法更新Word域（Field）。修改完成后必须在Word中右键TOC → 更新域 → 更新整个目录。

4. **图片段落伪装**：DOCX中的图片嵌入在独立段落中，这些段落`text.strip() == ''`但包含`w:drawing`/`wp:inline`等元素。清理空白段落时切勿删除图片段落。

5. **段落样式继承**：用`paragraph.text = new_text`设置文本可能丢失原样式。应操作`paragraph.runs`而非`paragraph.text`。

6. **TOC条目乱码**：如果TOC条目出现重复（如"7.2 不足与展望7.2 不足与展望..."），是TOC域未正常刷新导致的。清空对应段落后重设文本可临时修复，但最终仍需手动刷新TOC域。

### 修改后验证清单

- [ ] 段落总数是否合理（不应大量减少）
- [ ] 表格数量是否保持一致
- [ ] 各章节能找到标题
- [ ] 代码块内容完整
- [ ] 参考文献数量正确
- [ ] 无连续3+个空白段落（图片间距除外）
- [ ] TOC手动刷新后页码正确

---

## 六、批注模式（不修改原文，仅添加Word批注）

> **有现成脚本，优先用它**：`scripts/docx_批注.py` 已封装解包→插标记→写comments.xml→打包全过程，
> 且定位方式安全（文本搜索，不用段落索引）。下面的手工流程仅在需要特殊处理时使用。

当用户要求不修改原文、仅在有问题的地方添加Word批注时，执行本模式。

### 🚫 铁律：绝不能用 python-docx 段落索引做定位

**2026 届实测：python-docx 的段落索引与 document.xml 中的`w:p` 元素偏差最大 440 段落。**
按索引插入的批注会落到完全错误的位置。

**唯一安全做法：按文本内容搜索定位**（`docx_批注.py` 的 `find_paragraph()` 即此法）。

### ✅ 推荐：用 docx_批注.py

```bash
# 1. 写批注定义 JSON
# 2. 先干跑验证锚点
python scripts/docx_批注.py 论文.docx 批注.json --dry-run
# 3. 确认无误后生成
python scripts/docx_批注.py 论文.docx 批注.json -o 论文-批注版.docx
```

### 手工流程（仅在需要特殊处理时）

使用docx skill的 unpack → comment.py → 编辑document.xml → pack 流程：

### 步骤1：解包原始论文
```bash
python scripts/office/unpack.py <原始DOCX路径> <解包目录>/
```

### 步骤2：为每个审查问题调用comment.py创建批注定义
```bash
python scripts/comment.py <解包目录>/ <批注ID> "<批注文本>" --author "<作者名>" --initials "<缩写>"
```
- 批注ID从0开始递增，每个问题一个独立ID
- 批注文本需预先做XML转义（& → &amp;, < → &lt;, > → &gt;）
- 作者名使用审查人姓名（如"王吉振"）

### 步骤3：编辑document.xml添加批注标记
在目标段落中插入标记，格式如下：
```xml
<w:commentRangeStart w:id="0"/>
<!-- 段落原有内容 -->
<w:commentRangeEnd w:id="0"/>
<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="0"/>
</w:r>
```

**关键规则：**
- `commentRangeStart`和`commentRangeEnd`必须作为`w:p`的直接子元素，不能放在`w:r`内
- 如果段落有`w:pPr`元素，`commentRangeStart`必须插在`w:pPr`之后（pPr必须在段落最前面）
- 多个批注指向同一段落时需嵌套：`<crs id="0"/><crs id="1"/>...<cre id="1"/><cre id="0"/>`
- 每个批注的`commentReference` run放在`commentRangeEnd`之后

### 步骤4：打包生成批注版DOCX
```bash
python scripts/office/pack.py <解包目录>/ <输出DOCX路径> --original <原始DOCX路径>
```

### 已知限制
- 批注标记整段（段落级而非字词级精确定位）
- TOC等域内容无法直接批注（TOC条目在域展开后才可见）
- 批注在Word中正常显示，WPS兼容性需验证

### ⚠️ 环境要求

`docx_批注.py` 只依赖 lxml。`excel_双表同步.py` 额外需要 openpyxl。
两者安装依赖后任意 Python 3.10+ 均可运行：

```bash
pip install python-docx lxml openpyxl
```

### 批注文字风格要求
- **口语化**：用"删掉就完事了""补上就行"等日常表达，让学生感觉是真人老师在说话
- **简短**：每条1-3句话，不要长篇大论
- **可操作**：直接告诉学生怎么改，避免"建议""应当"等空泛表述，用"把...贴上来""加个...字段就行"
- **避免术语**：少用"学术规范""格式标准"等词，用"不像论文""读着太生硬"等学生能懂的表述

### 两项关键避坑
1. **TOC与正文标题重叠**：章节标题同时出现在TOC和正文，`find_paragraph_containing`会优先匹配TOC。解决方案：使用`occurrence=2`跳过TOC第一次匹配。
2. **嵌套批注丢失**：单遍法（边找边插）会导致共享段落的后续批注标记找不到前一批注的索引。解决方案：**两遍法**——先收集所有段落到批注的映射，再统一插入标记，插入逻辑：所有`commentRangeStart`按顺序插在最前，所有`commentRangeEnd`按倒序插在最后，所有`commentReference`按正序插在最后。

---
## 七、DOCX 格式读取避坑（CRITICAL）
> 记录 python-docx 读取 Word 格式时的常见误判，每次做格式审查前回顾本节。
>
> **💡 如果 wps-agent MCP 可用，直接用 WPS Agent 读取格式，以下所有坑都不会踩。**

### 坑1：`run.font.name` 只读西文字体

Word 将中英文字体分开存储：
```xml
<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/>
```

`run.font.name` 只返回 `ascii` 的值。如果 ascii 为空但 eastAsia 有值，`font.name` 返回 `None`，导致误判"字体缺失"。

**正确做法**：始终读 XML 层的 `rFonts` 三个属性：
```python
def get_fonts(run):
    """返回 (ascii字体, 中文字体, 西文字号val, 中文字号val)"""
    rPr = run._element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    if rPr is None:
        return (None, None, None, None)
    rFonts = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    ascii_f = hAnsi_f = east_f = None
    if rFonts is not None:
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        ascii_f = rFonts.get(f'{ns}ascii')
        hAnsi_f = rFonts.get(f'{ns}hAnsi')
        east_f = rFonts.get(f'{ns}eastAsia')
    sz = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    szCs = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
    sz_val = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if sz is not None else None
    szCs_val = szCs.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if szCs is not None else None
    return (ascii_f or hAnsi_f, east_f, sz_val, szCs_val)
```

**判断规则**：
- 中文正文：`eastAsia=宋体` 且 `ascii=Times New Roman`
- 标题：`eastAsia=黑体` 且 `ascii=黑体`
- 字号 `sz=24` = 12pt（小四），`sz=21` = 10.5pt（五号），`sz=32` = 16pt（三号）

### 坑2：段落格式的三个独立维度

段落间距有三种设置，必须分头检查：
```
<w:spacing w:before="156" w:after="0" w:line="360" w:lineRule="auto"/>
```
- `before` / `after`：段前/段后间距（单位 twips，1pt=20twips）
- `line` / `lineRule`：行距。lineRule="auto" 时 line=240 为单倍行距，line=360 为1.5倍
- 如果三个属性都没有 → 继承样式默认值，不可直接判为"无间距"

### 坑3：必须先读模板再下结论

任何格式问题在标记 P0 之前，必须先确认模板对应位置的实际格式。典型误判：
- 模板页眉为空 → 学生页眉为空 = 合规（非P0）
- 模板声明页用仿宋 → 学生声明页用仿宋 = 合规
