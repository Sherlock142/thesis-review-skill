# WPS Agent MCP 工具速查（格式读取核心）

> 前提：已按 `docs/MCP接入指南.md` 接入 wps-agent。
> 为什么必须用它：DOCX 的字体/字号/行距/缩进大量藏在**样式继承链**（`basedOn` + `docDefaults`）里。
> python-docx 只读直接格式，读不出继承值，会返回 `?`/`inherit`/`None`，导致漏检。
> WPS Agent 走 COM 解析完整生效值，格式保真度接近 100%。

---

## 核心工作流（检查论文格式）

```
1. document.open          → 打开论文（拿到 doc_index，通常=1）
2. content.full_structure → 了解段落类型分布
3. format.get_paragraph_format → 逐段读对齐/缩进/行距/段间距
4. format.get_run_font         → 读字体（含中文字体 name_far_east）
5. document.close(save_changes:false) → 只读检查，务必丢弃改动
```

---

## 一、document — 文档生命周期

| action | 说明 | 关键参数 |
|---|---|---|
| `open` | 打开文档 | `filepath` 必填 |
| `close` | 关闭 | `doc_index`；只读检查时传 `save_changes:false` |
| `health_check` | 探活 | 无。返回 `healthy` / `application` |
| `save` | 另存 | `filepath` 可选 |

```javascript
{ action: "open", filepath: "C:/path/论文.docx" }
{ action: "close", doc_index: 1, save_changes: false }
```

## 二、content — 内容读取

| action | 说明 |
|---|---|
| `full_text` | 全文纯文本 |
| `full_structure` | 段落结构概览（含每段摘要） |
| `paragraphs` | 指定范围段落文本。`start`/`count` 或 `from_para`/`to_para` |
| `runs_detail` | 读某段全部 run 及其字体（定位可改 run 必备） |

> ⚠️ 段落索引：`format` 按 COM 计数，`content` 可能不同。以 `format` 为准。
> 中文路径在本机 COM 可能显示乱码，但不影响读取值。

## 三、format — 格式读写（核心）

### get_paragraph_format — 读段落格式

```javascript
{ action: "get_paragraph_format", para_index: 45, doc_index: 1 }
```

返回可读字段：

| 字段 | 含义 | 示例 |
|---|---|---|
| `alignment` | 对齐（可读字符串） | `"justify"` / `"center"` / `"left"` |
| `first_line_indent` | 首行缩进 pt（负=悬挂） | `24.0` |
| `left_indent` / `right_indent` | 左右缩进 pt | `0.0` |
| `line_spacing_rule` | 行距规则 | `"exactly"` / `"multiple"` / `"single"` |
| `line_spacing` | 行距值 pt | `20.0` |
| `space_before` / `space_after` | 段前/段后 pt | `0.0` |

> 行距单位是**磅不是倍数**。1.5 倍行距 = `18`（1.5×12pt），传 `1.5` 会把行距压扁。

### get_run_font — 读字体

```javascript
{ action: "get_run_font", para_index: 45, run_index: 1, doc_index: 1 }
```

返回：`name`（西文）、`name_far_east`（**中文字体**）、`size`（pt）、`bold`、`italic`、`underline`、`color_rgb`。

> 中文字体必须看 `name_far_east`。只看 `name`（西文）会误判——这是 python-docx `run.font.name` 的经典坑。

### 只读检查的典型调用组合

对照模板核验时，对关键段落（封面/摘要/正文/标题/参考文献）分别调用：

```javascript
{ action: "get_paragraph_format", para_index: N }
{ action: "get_run_font", para_index: N, run_index: 1 }
```

把返回的值与模板同位置段落对比，差异即为格式问题。

---

## 已知限制

1. **段落索引不一致**：`format` 按 COM 计数，`content` 可能不同，以 `format` 为准。
2. **实际驱动 Word**：本机 `health_check` 返回 `application: Microsoft Word`（COM 的 WPS ProgID 指向 Word）。只要有 Word 或 WPS 均可，格式都保真。
3. **`search.find` 有时返回空**：此时用 `full_text` 手动定位段落。
4. **`replace_paragraph_text` 有严重 bug**：会静默删除段落，禁止使用。改文本用 `replace_runs` + `preserve_format:true`。
5. **只读检查务必关闭不保存**：`document.close(save_changes:false)`，避免污染原稿。

---

## 降级方案（WPS 不可用时）

若 wps-agent 连不上，用内置脚本做**内容结构**预审（无法读字体/行距等版式细节）：

```bash
python skills/thesis-review/scripts/docx_预审.py IN-1_学生论文/
```

> 明确边界：降级模式只能查章节缺失、图件、参考文献数量、题文一致性等内容问题；
> 字体/字号/缩进/行距这类版式核验必须有 wps-agent。审查报告里要如实说明用了哪种模式。
