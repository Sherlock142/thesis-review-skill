# WPS Agent MCP 接入指南

> 为什么必须接入：论文版式（字体/字号/行距/缩进）藏在 DOCX 的样式继承链里，
> 普通 Python 库读不出来。接入 wps-agent 后，Agent 通过 COM 读取生效值，格式核验才准确。
> 本指南让你从零到能用，约 15 分钟。

---

## 0. 前置条件

- **Windows** 系统（wps-agent 走 COM，依赖 Word 或 WPS）
- 已安装 **Microsoft Word** 或 **WPS Office 2019+**（本机实测 COM 的 WPS ProgID 实际指向 Word，二者有一即可）
- 已安装 **WorkBuddy**（或任意支持 MCP 的 Agent 工具，如 Cursor / Cline / opencode）
- Python 3.10+

---

## 1. 安装 wps-agent（MCP 服务端）

wps-agent 是开源项目，直接克隆官方仓库：

```bash
git clone https://github.com/alllyx520-bot/wps-agent.git
cd wps-agent

# 创建虚拟环境（推荐 conda，也可用 venv）
conda create -n wps-agent python=3.11 -y
conda activate wps-agent

# 或用 venv
python -m venv .venv
.venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

验证：确认 `mcp_server.py` 和 `requirements.txt` 都在目录里。

---

## 2. 配置 MCP 服务端

### WorkBuddy

编辑 `~/.workbuddy/mcp.json`，在 `mcpServers` 里加入：

```json
"wps-agent": {
  "command": "C:/path/to/wps-agent/.venv/Scripts/python.exe",
  "args": ["C:/path/to/wps-agent/mcp_server.py"],
  "type": "stdio",
  "env": {
    "PYTHONPATH": "C:/path/to/wps-agent"
  }
}
```

把两个 `C:/path/to/wps-agent` 换成你实际的克隆路径（用正斜杠 `/`）。

> 配置不会自动生效。重启 WorkBuddy，到连接器管理页找到 wps-agent，点「信任 / 启用」。

### 其他 Agent 工具（Cursor / Cline / opencode）

MCP 配置格式略有差异，但核心都是三要素：
- command：python 解释器绝对路径
- args：`mcp_server.py` 的绝对路径
- env：`PYTHONPATH` 指向 wps-agent 目录

wps-agent 仓库的 `opencode_config/mcp-config.jsonc` 提供了 opencode 的现成范例。

---

## 3. 安装论文审查 Skill

把本项目的 `skills/thesis-review/` 整个目录，复制到你的用户级 Skill 目录：

```bash
# WorkBuddy
cp -r skills/thesis-review ~/.workbuddy/skills/

# 结果应形如：
# ~/.workbuddy/skills/thesis-review/SKILL.md
# ~/.workbuddy/skills/thesis-review/scripts/
# ~/.workbuddy/skills/thesis-review/references/
```

重启 WorkBuddy，Skill 即被识别（触发词见 SKILL.md 的 description）。

---

## 4. 安装 Python 依赖

```bash
pip install python-docx lxml openpyxl
```

`docx_预审.py`（批量预审）和 `docx_批注.py`（批注注入）用这些库。
wps-agent 自己的依赖在它的 `requirements.txt` 里，已单独装过。

---

## 5. 放好你的模板

把**你自己学校的论文格式模板**（.docx）放到项目根目录的 `templates/`：

```
thesis-review-skill/
└── templates/
    └── 你的学校论文模板.docx
```

> 模板只在做精确版式核验（字体/字号/页边距）时才需要。
> 没有模板也能用——Agent 会退化为只查内容结构问题，并在报告里注明。

---

## 6. 验证接入成功

对 Agent 说：

```
检查 wps-agent MCP 是否正常
```

Agent 应调用 `document.health_check`，返回类似：

```json
{"healthy": true, "application": "Microsoft Word", ...}
```

再验证完整能力：让 Agent 读一份论文的某段格式，应返回可读的
`alignment: "justify"`、`name_far_east: "宋体"`、`line_spacing: 20.0` 等值，
而不是 `?` / `inherit` / `None`。

---

## 常见问题

| 症状 | 原因 / 解决 |
|---|---|
| MCP 连不上 | 路径用了反斜杠或相对路径；改用正斜杠绝对路径，确认 venv python 存在 |
| `health_check` 失败 | 未装 Word/WPS，或 COM 组件未注册；先装 Office 再试 |
| 中文路径乱码 | COM 对中文路径显示可能乱码，但读取值不受影响；可先把文件放纯英文路径排查 |
| 读出来全是 `?`/`inherit` | 说明退化成了 python-docx 模式，wps-agent 没接上——回第 2、6 步 |
| Skill 没被识别 | 没重启 WorkBuddy；或 SKILL.md 不在 `~/.workbuddy/skills/thesis-review/` 下 |

---

## 能力边界（重要）

| 能力 | wps-agent 模式 | 仅 python-docx 降级模式 |
|---|---|---|
| 章节/图件/参考文献等内容问题 | ✅ | ✅ |
| 字体 / 字号 | ✅ 精确 | ❌ 读不出 |
| 行距 / 缩进 / 对齐 | ✅ 精确 | ❌ 读不出 |
| 与模板逐字段对比 | ✅ | ⚠️ 仅内容层面 |

**结论**：要「完整工作能力」，wps-agent 是必需的。降级模式只能做内容预审。
