# 本科毕业论文审查 Skill（thesis-review）

一个可部署到 AI Agent 的论文审查能力包。用户上传自己学校的论文格式模板后，
Agent 即可对照模板审查论文的**格式 + 内容**，按 P0/P1/P2 输出审查报告、Word 批注版或整改版。

- 格式核验走 **wps-agent MCP**（COM 读取生效值，保真度高）——不是普通 Python 库那种读不出字体/行距的降级方案
- 内容预审内置 **22 条规则**（来自真实审查经验），一次扫全班
- 支持单篇深度审查 + 批量预审

---

## 快速开始（3 步）

```bash
# 1. 安装 wps-agent MCP 服务端（开源项目，官方仓库）
git clone https://github.com/alllyx520-bot/wps-agent.git
cd wps-agent && pip install -r requirements.txt

# 2. 装本 Skill 的 Python 依赖
pip install python-docx lxml openpyxl

# 3. 把 Skill 放进你的 Agent
cp -r skills/thesis-review ~/.workbuddy/skills/
```

然后按 [**MCP接入指南.md**](docs/MCP接入指南.md) 配置 MCP 服务端并重启 Agent。
接入成功后，把你的学校论文模板放进 `templates/`，对 Agent 说「审查这篇论文」即可。

> 详细配置、验证方法、常见问题见 [docs/MCP接入指南.md](docs/MCP接入指南.md)。

---

## 项目结构

```
thesis-review-skill/
├── README.md                      本文件
├── LICENSE                        MIT 协议
├── docs/
│   └── MCP接入指南.md              从零接入 wps-agent（核心）
├── skills/
│   └── thesis-review/
│       ├── SKILL.md                Skill 主文件（审查方法论 / 7 种模式）
│       ├── scripts/
│       │   ├── docx_预审.py        批量预审，22 条规则，输出 P0/P1/P2
│       │   ├── docx_批注.py        Word 批注注入，文本搜索定位
│       │   ├── excel_双表同步.py   选题表 ↔ 答辩意见按学号交叉同步
│       │   ├── 环境检查.py         依赖自检
│       │   └── 新建学生审查.py     新建学生建档
│       └── references/
│           ├── wps-agent工具速查.md   MCP 工具用法（格式读取核心）
│           ├── 高频问题实证库.md      该查什么
│           ├── 批量作业手册.md        一届多人的流水线
│           └── 归档材料整改要点.md    任务书/开题等表格整改
└── templates/
    └── README.md                  模板放置说明（请放你的学校模板）
```

---

## 三种输出

| 模式 | 触发词 | 输出 |
|---|---|---|
| 审查报告 | 全篇审查 / 出报告 | `.md` 报告，含核查结果 + 全篇问题 |
| 批注版 | 加批注 / 批注版 | DOCX，审查意见以 Word 批注插入（不改原文） |
| 整改版 | 直接改 / 整改版 | DOCX，按意见改好的干净版本（无批注） |

统一审查流程：答辩意见三态核查（✅已修正 / ❌未修正 / ⚠️部分修正）→ 全篇格式内容审查 → 去重合并。

---

## 能力边界

- **接入 wps-agent**：格式 + 内容全量审查，含字体/字号/行距/缩进的模板逐字段对比。
- **未接入（降级）**：只能查内容结构（章节、图件、参考文献、题文一致性等），查不了版式细节。审查报告会如实注明所用模式。

预审脚本是正则 + 结构统计，**有误报**，只用于筛可疑，判断需人工。不要把 P0 直接发给学生。

---

## 协议

MIT — 自由使用、修改、分发，保留署名即可。
