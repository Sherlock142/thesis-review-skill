#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文批量预审器 —— 2027 届通用版

设计目标：把去年「一个学生写一个脚本」的做法，变成「一个脚本扫全班」。

用法
----
  # 扫描单个文件
  python docx_预审.py "IN-1_学生论文/202400000001-张三-xxx.docx"

  # 扫描整个目录（推荐：一次扫全班）
  python docx_预审.py IN-1_学生论文/

  # 输出 JSON 供后续批注/整改脚本消费
  python docx_预审.py IN-1_学生论文/ --json out.json

环境
----
  python-docx（读段落/表格结构） + lxml（精确定位）
  ⚠️ openpyxl 只在系统 Python 3.10 有；本脚本不依赖 openpyxl。
  ⚠️ **绝不可用 python-docx 的段落索引做 XML 定位**——实测偏差最大 440 段落。
     本脚本的定位一律走「文本内容搜索」，与 Skill 约定一致。
"""

import os
import sys
import re
import json
import zipfile
import argparse
from datetime import datetime
from types import SimpleNamespace

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    print("缺少 python-docx。请运行：pip install python-docx")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────
# 规则表：每条 = (规则ID, 级别, 说明, 检测函数)
# 级别：P0 = 学院格式硬性要求 / 答辩必挂；P1 = 明显质量问题；P2 = 建议
# ─────────────────────────────────────────────────────────────

RULES = []


def rule(rid, level, desc):
    def deco(fn):
        RULES.append((rid, level, desc, fn))
        return fn
    return deco


# ── 一、章节结构类 ──

@rule("SEC-001", "P0", "缺少「总结与展望」章节（本科论文必备）")
def r_sec_summary(ctx):
    if not any(re.search(r"(总结|结论|展望)", h["text"]) for h in ctx.headings):
        return ["论文中未找到「总结」「展望」或「结论」类章节——本科论文必备，答辩组常重点核查"]
    return []


@rule("SEC-002", "P1", "章节编号与标题之间缺少空格（如「5系统实现」）")
def r_sec_style(ctx):
    # 真正的错误：数字编号后直接跟汉字/字母，没有空格
    bad = [h["text"] for h in ctx.headings
           if re.match(r"^\d+(\.\d+)*[一-龥A-Za-z]", h["text"])]
    if bad:
        return [f"编号与标题缺空格：{b[:40]}" for b in bad[:5]]
    return []


@rule("SEC-003", "P1", "章节编号格式不统一（半角/全角点号混用）")
def r_sec_numfmt(ctx):
    half = [h["text"] for h in ctx.headings if re.match(r"^\d+\.\d+", h["text"])]
    full = [h["text"] for h in ctx.headings if re.match(r"^\d+．\d+", h["text"])]
    if half and full:
        return [f"同时存在半角点号({half[0][:16]})与全角点号({full[0][:16]})两种格式"]
    return []


@rule("SEC-004", "P1", "疑似缺少核心图件（架构图/用例图/E-R 图/功能模块图）")
def r_sec_figs(ctx):
    # 图题可能写作「图x-x」「图 x-x」「图x－x」，统一抽出编号
    cap_blob = "\n".join(ctx.captions) + "\n" + "\n".join(
        p["text"] for p in ctx.paragraphs)
    need = {
        "架构图": [r"架构", r"总体结构", r"系统结构", r"技术架构", r"部署"],
        "用例图": [r"用例"],
        "E-R图": [r"E-?R\s*图", r"E-?R图", r"实体.{0,3}关系"],
        "功能模块图": [r"功能模块", r"功能结构", r"系统功能", r"功能组成", r"功能设计"],
    }
    missing = []
    for name, pats in need.items():
        if not any(re.search(p, cap_blob, re.I) for p in pats):
            missing.append(name)
    if missing:
        return [f"全文未提及：{'、'.join(missing)}"]
    return []


# ── 二、图题 / 表题 ──

@rule("CAP-001", "P0", "图表题编号重复或跳号")
def r_cap_dup(ctx):
    figs = [c for c in ctx.captions if re.match(r"^图\s*\d+-\d+", c)]
    tabs = [c for c in ctx.captions if re.match(r"^表\s*\d+-\d+", c)]
    out = []
    for label, group in (("图", figs), ("表", tabs)):
        if not group:
            continue
        seen, dup = set(), []
        for c in group:
            m = re.match(rf"^([{label}]?\s*\d+-\d+)", c)
            if not m:
                continue
            k = m.group(1).replace(" ", "")
            if k in seen:
                dup.append(k)
            seen.add(k)
        if dup:
            out.append(f"{label}题编号重复：{'、'.join(sorted(set(dup))[:5])}")
    return out


@rule("CAP-002", "P1", "图题/表题格式不统一（有空格与无空格混用）")
def r_cap_fmt(ctx):
    withsp = [c for c in ctx.captions if re.match(r"^[图表]\s+\d", c)]
    nosp = [c for c in ctx.captions if re.match(r"^[图表]\d", c)]
    if withsp and nosp:
        return [f"混用两种格式：「{withsp[0][:14]}」（有空格）与「{nosp[0][:14]}」（无空格）"]
    return []


@rule("CAP-003", "P1", "多数图在正文中无「如图X-X」引用，图文可能脱节")
def r_cap_ref(ctx):
    figs = [c for c in ctx.captions if re.match(r"^图\s*\d+-\d+", c)]
    refs = set()
    for p in ctx.paragraphs:
        for m in re.finditer(r"如图\s*(\d+-\d+)", p["text"]):
            refs.add(m.group(1))
    n = [re.match(r"^图\s*(\d+-\d+)", c).group(1) for c in figs
         if re.match(r"^图\s*(\d+-\d+)", c)]
    if not n:
        return []
    orphan = [x for x in n if x not in refs]
    if len(orphan) > max(3, len(n) * 0.6):
        return [f"{len(orphan)}/{len(n)} 个图在正文中无「如图X-X所示」引用，图文可能脱节"]
    return []


# ── 三、代码类 ──

@rule("CODE-001", "P0", "疑似伪代码/中文变量名（代码块中含中文标识符）")
def r_code_cn(ctx):
    bad = []
    for blk in ctx.code_blocks:
        for l in blk.split("\n"):
            stripped = l.strip()
            if not re.search(r"[\u4e00-\u9fa5]", stripped):
                continue
            # 去掉行注释后还剩中文 → 真问题
            nocomment = re.sub(r"//.*|#.*|/\*.*?\*/", "", stripped)
            nocomment = re.sub(r'"[^"]*"|\'[^\']*\'', "", nocomment)  # 去字符串
            if re.search(r"[\u4e00-\u9fa5]", nocomment):
                bad.append(stripped[:60])
    if bad:
        return [f"代码标识符含中文（疑似伪代码）：{len(bad)} 处，例：{bad[0]}"]
    return []


@rule("CODE-002", "P1", "代码占位式写法（出现 TODO / xxx / 待补充 等）")
def r_code_placeholder(ctx):
    for blk in ctx.code_blocks:
        if re.search(r"\bTODO\b|待补充|待完善|此处省略|\.\.\.\s*$", blk, re.M):
            for l in blk.split("\n"):
                if re.search(r"\bTODO\b|待补充|待完善|此处省略", l):
                    return [f"代码含占位标记：{l.strip()[:50]}"]
    return []


@rule("CODE-003", "P1", "疑似整章为 API 接口文档（函数签名密度过高）")
def r_code_api(ctx):
    total = sum(len(b) for b in ctx.code_blocks)
    if total < 400:
        return []
    sig = sum(len(re.findall(r"(public|private|def|function)\s+\w+\s*\(", b))
              for b in ctx.code_blocks)
    words = max(1, len(ctx.full_text))
    if sig >= 8 and sig * 100 / words > 2:
        return [f"函数签名 {sig} 处且正文叙述极少，疑似写成 API 文档而非系统实现章节"]
    return []


# ── 四、界面截图 ──

@rule("IMG-001", "P0", "系统实现章节疑似无界面截图（实现章节 0 张图）")
def r_img_nofig(ctx):
    if "实现" not in ctx.section_hint and "实现" not in ctx.full_text[:2000]:
        return ["正文中未定位到「系统实现」章节，请人工确认"]
    imgs_in_impl = ctx.imgs_after_impl
    if imgs_in_impl == 0:
        return ["「系统实现」相关章节内 0 张图片——实现章节应有界面截图，纯代码堆砌是答辩高频问题"]
    return []


# ── 五、参考文献 ──

@rule("REF-001", "P1", "参考文献数量偏少（本科建议 ≥15 条）")
def r_ref_count(ctx):
    if ctx.ref_count == 0:
        return ["未识别到参考文献条目"]
    if ctx.ref_count < 15:
        return [f"仅 {ctx.ref_count} 条，建议补至 15 条以上"]
    return []


@rule("REF-002", "P1", "参考文献含非学术来源（网页/百科/商业站点）")
def r_ref_source(ctx):
    bad = [r for r in ctx.refs
           if re.search(r"(baike|zhihu|sohu|163\.com|csdn\.net|百度百科|知乎|论坛)", r, re.I)]
    if bad:
        return [f"{len(bad)} 条疑似非学术来源，例：{bad[0][:40]}"]
    return []


@rule("REF-003", "P2", "参考文献条目缺页码或卷号")
def r_ref_incomplete(ctx):
    miss = []
    for r in ctx.refs:
        if re.search(r"\[J\]", r) and not re.search(r"[:：]\s*\d", r):
            miss.append(r[:36])
    if len(miss) > max(2, len(ctx.refs) * 0.4):
        return [f"{len(miss)}/{len(ctx.refs)} 条期刊文献缺页码/卷号，例：{miss[0]}"]
    return []


# ── 六、英文摘要 ──

@rule("ABS-001", "P1", "英文摘要疑含中文标点（逗号/引号用全角）")
def r_abs_cnpunct(ctx):
    m = re.search(r"(?is)abstract[\s\S]{0,1200}", ctx.full_text)
    if not m:
        return []
    seg = m.group(0)
    bad = re.findall(r"[\u4e00-\u9fa5，。；：？！“”（）【】]", seg)
    if bad:
        return [f"Abstract 段内出现 {len(bad)} 个中文/全角字符（含中文标点）"]
    return []


@rule("ABS-002", "P1", "英文摘要长度异常")
def r_abs_len(ctx):
    m = re.search(r"(?is)abstract[\s\S]{0,1500}", ctx.full_text)
    if not m:
        return []
    words = len(re.findall(r"[A-Za-z']+", m.group(0)))
    if words and (words < 80 or words > 600):
        return [f"Abstract 约 {words} 词，建议 150–400 词"]
    return []


# ── 七、错别字与术语 ──

TYPOS = [
    ("登陆", "登录"), ("身分", "身份"), ("帐号", "账号"), ("其它", "其他"),
    ("做为", "作为"), ("既使", "即使"), ("按装", "安装"), ("必需要", "必须"),
]


@rule("TYPO-001", "P2", "常见错别字")
def r_typo(ctx):
    hits = []
    for w, right in TYPOS:
        n = ctx.full_text.count(w)
        if n:
            hits.append(f"「{w}」→「{right}」×{n}")
    return [f"共 {len(hits)} 类：{h}" for h in hits]


@rule("TYPO-002", "P2", "疑似存在「拟引入/拟采用/计划实现」等未完成表述")
def r_future(ctx):
    hits = re.findall(r"[^\n。]{0,24}(拟引入|拟采用|计划实现|将在后续|有待完善)", ctx.full_text)
    if hits:
        return [f"{len(hits)} 处未来时表述（论文应描述已实现内容），例：{hits[0].strip()[:36]}"]
    return []


# ── 八、数据库设计 ──

@rule("DB-001", "P1", "数据库字段疑似使用中文拼音命名")
def r_db_pinyin(ctx):
    hits = set()
    for t in ctx.tables:
        for cell in t["headers"] + t["cells_flat"]:
            if re.fullmatch(r"[a-z]{2,10}", str(cell).strip()) and _looks_pinyin(str(cell)):
                hits.add(str(cell).strip())
    if hits:
        return [f"{len(hits)} 个疑似拼音字段名，例：{'、'.join(list(hits)[:5])}——要求改英文命名"]
    return []


PYININ = re.compile(r"^(m[aei]?ng|ha[oa]o|y[ua]?n|zh[aei]ng|sh[ou]|g[ou]ai|"
                     r"x[ie]ng|ji[ao]|l[ia]o|d[ai]|z[hiao]|l[ie]|p[ai]|q[ie]|"
                     r"k[ei]|c[ao]|l[ei]|s[an]|g[ong]|z[ao]|x[iao]|sh[ang]|"
                     r"j[ia]|j[ing]|ch[eng]|t[ian]|l[un]|w[en]|z[hi])$", re.I)


def _looks_pinyin(s):
    return bool(PYININ.match(s.strip()))


@rule("DB-002", "P1", "字段允许为空但业务上应为必填（E-R 图与表结构常见矛盾）")
def r_db_nullable(ctx):
    hits = []
    for t in ctx.tables:
        for row in t["rows"][:60]:
            line = " ".join(str(c) for c in row)
            if re.search(r"NULL|允许为空|可空", line) and re.search(r"user_id|admin|角色|状态", line, re.I):
                hits.append(t["title"][:24])
    return [f"表「{h}」存在可空且语义应为必填的字段，请核对 E-R 图与正文描述是否一致" for h in dict.fromkeys(hits)][:3]


# ── 九、测试类 ──

@rule("TEST-001", "P1", "测试数据数量与展示数量不符（声称 N 条仅展示 M 条）")
def r_test_count(ctx):
    claims = re.findall(r"(?:共|计|合计)\s*(\d+)\s*(?:条|个|组|项)", ctx.full_text)
    if not claims:
        return []
    n = max(int(c) for c in claims)
    if n >= 20 and len(ctx.tables) and n > 0:
        # 粗略估计：表格里的行数
        rows = sum(len(t["rows"]) for t in ctx.tables if re.search(r"测试|用例", t["title"]))
        if rows and rows < n * 0.5:
            return [f"文中声称 {n} 条测试，仅展示约 {rows} 条——测试数据数量与展示不符会被质疑可信度，建议附录补全"]
    return []


@rule("TEST-002", "P2", "测试章节缺边界/异常场景用例")
def r_test_edge(ctx):
    if not re.search(r"测试", ctx.full_text):
        return []
    has = re.search(r"(边界|异常|非法|空值|越界|异常场景)", ctx.full_text)
    if not has:
        return ["测试部分未见边界/异常场景用例，建议补充"]
    return []


# ── 十、论文-实现一致性（需求与实现脱节是高频问题）──

@rule("CONS-001", "P0", "代码中出现但需求分析未提及的功能（前后脱节）")
def r_consistency(ctx):
    # 抽取代码里的业务关键词，与需求分析章节文本比对
    impl_start = ctx.full_text.find("系统实现")
    req_end = ctx.full_text.find("系统设计")
    if impl_start < 0 or req_end < 0:
        return []
    req_text = ctx.full_text[:req_end]
    impl_text = ctx.full_text[impl_start:impl_start + 6000]
    words = set(re.findall(r"[一-龥]{2,6}(?:管理|模块|功能|秒杀|砍价|评价|收藏|购物车)", impl_text))
    missing = [w for w in words if w.replace("模块", "") not in req_text and w not in req_text]
    if len(missing) >= 2:
        return [f"实现章节出现需求分析未提及的功能：{'、'.join(missing[:4])}——需求与实现脱节是答辩高频问题，需在需求分析章节补充"]
    return []


# ─────────────────────────────────────────────────────────────
# 上下文构建
# ─────────────────────────────────────────────────────────────

CODE_FONTS = {"Consolas", "Courier New", "宋体-代码", "等线-代码", "Lucida Console", "monospace"}


def build_context(path):
    doc = Document(path)

    paragraphs = []
    headings = []
    full = []
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if not t:
            continue
        paragraphs.append({"i": i, "text": t, "style": p.style.name})
        full.append(t)
        if p.style.name.startswith("Heading") or re.match(r"^\d+(\.\d+)*\s", t) or re.match(r"^第[一二三四五六七八九十]+章", t):
            headings.append({"i": i, "text": t, "style": p.style.name})

    full_text = "\n".join(full)

    # 图题表题：短、含"图x-y"/"表x-y"
    captions = [p["text"] for p in paragraphs
                if re.match(r"^[图表]\s*\d+[\-－—]\d+", p["text"]) and len(p["text"]) < 60]

    # 代码块：字体为等宽 或 多行缩进
    code_blocks = []
    for p in doc.paragraphs:
        fname = ""
        try:
            rPr = p._element.find(qn('w:r'))
            if rPr is not None:
                rf = rPr.find(qn('w:rFonts'))
                if rf is not None:
                    fname = rf.get(qn('w:ascii')) or ""
        except Exception:
            pass
        if fname in CODE_FONTS or (len(p.text) > 120 and re.search(r"[{};]\s*$", p.text.strip())):
            if p.text.strip():
                code_blocks.append(p.text)

    # 表格
    tables = []
    for ti, t in enumerate(doc.tables):
        grid = [[c.text.strip() for c in r.cells] for r in t.rows]
        title = ""
        # 向上找最近的表题
        for p in reversed(doc.paragraphs[:max(0, 200)]):
            pass
        tables.append({
            "idx": ti,
            "title": title,
            "headers": grid[0] if grid else [],
            "rows": grid[1:] if len(grid) > 1 else [],
            "cells_flat": [c for row in grid for c in row],
        })
    # 给表格配标题（简化：按出现顺序匹配表题）
    cap_iter = iter([c for c in captions if c.startswith("表")])
    for t in tables:
        try:
            t["title"] = next(cap_iter)
        except StopIteration:
            t["title"] = f"表{t['idx'] + 1}"

    # 图片计数 + 实现章节后的图片
    imgs_total = len(doc.inline_shapes)
    imgs_after_impl = 0
    seen_impl = False
    for p in doc.paragraphs:
        if re.search(r"系统实现|功能实现|详细实现", p.text):
            seen_impl = True
        if seen_impl and p._element.findall('.//' + qn('w:drawing')) or \
           p._element.findall('.//' + qn('w:pict')):
            imgs_after_impl += 1

    # 参考文献
    refs = []
    in_ref = False
    for p in paragraphs:
        if re.match(r"^\s*(参考文献|参 考 文 献)\s*$", p["text"]):
            in_ref = True
            continue
        if in_ref:
            if re.match(r"^\s*致\s*谢|^\s*附\s*录|^\s*致谢", p["text"]):
                break
            if re.match(r"^\s*\[\d+\]", p["text"]):
                refs.append(p["text"])

    section_hint = ""
    for p in paragraphs:
        if re.search(r"系统实现", p["text"]):
            section_hint = "实现"
            break

    return SimpleNamespace(**{
        "path": path,
        "paragraphs": paragraphs,
        "headings": headings,
        "full_text": full_text,
        "captions": captions,
        "code_blocks": code_blocks,
        "tables": tables,
        "refs": refs,
        "ref_count": len(refs),
        "imgs_total": imgs_total,
        "imgs_after_impl": imgs_after_impl,
        "section_hint": section_hint,
    })


def scan(path):
    ctx = build_context(path)
    findings = []
    for rid, level, desc, fn in RULES:
        try:
            msgs = fn(ctx) or []
        except Exception as e:
            msgs = [f"规则执行异常：{e}"]
        for m in msgs:
            findings.append({"rule": rid, "level": level, "desc": desc, "msg": m})
    order = {"P0": 0, "P1": 1, "P2": 2}
    findings.sort(key=lambda f: (order.get(f["level"], 9), f["rule"]))
    stats = {
        "file": os.path.basename(path),
        "size_kb": round(os.path.getsize(path) / 1024, 1),
        "paragraphs": len(ctx.paragraphs),
        "headings": len(ctx.headings),
        "tables": len(ctx.tables),
        "captions": len(ctx.captions),
        "images_total": ctx.imgs_total,
        "images_in_impl": ctx.imgs_after_impl,
        "code_blocks": len(ctx.code_blocks),
        "references": ctx.ref_count,
        "p0": sum(1 for f in findings if f["level"] == "P0"),
        "p1": sum(1 for f in findings if f["level"] == "P1"),
        "p2": sum(1 for f in findings if f["level"] == "P2"),
    }
    return stats, findings


def render(stats, findings, show_all):
    L = []
    L.append("=" * 72)
    L.append(f"【{stats['file']}】  {stats['size_kb']} KB")
    L.append("  结构：段落 %d | 标题 %d | 表 %d | 图题表题 %d | 图片 %d（实现章节 %d）"
             % (stats["paragraphs"], stats["headings"], stats["tables"],
                stats["captions"], stats["images_total"], stats["images_in_impl"]))
    L.append("  代码块 %d | 参考文献 %d" % (stats["code_blocks"], stats["references"]))
    L.append(f"  → P0 {stats['p0']} 项 / P1 {stats['p1']} 项 / P2 {stats['p2']} 项")
    L.append("=" * 72)
    for f in findings:
        if f["level"] == "P2" and not show_all:
            continue
        L.append(f"  [{f['level']}] {f['rule']}  {f['desc']}")
        L.append(f"         → {f['msg']}")
    if not show_all:
        n = stats["p2"]
        if n:
            L.append(f"  （另有 {n} 项 P2 建议，加 --all 查看）")
    L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="论文批量预审器（跨年通用）")
    ap.add_argument("target", help="docx 文件或目录")
    ap.add_argument("--json", help="结果写入 JSON 文件")
    ap.add_argument("--all", action="store_true", help="显示 P2 建议")
    a = ap.parse_args()

    if os.path.isdir(a.target):
        files = sorted(os.path.join(a.target, f) for f in os.listdir(a.target)
                       if f.lower().endswith(".docx") and not f.startswith("~$"))
    else:
        files = [a.target]
    if not files:
        print("未找到 docx 文件")
        return

    all_results = []
    for f in files:
        try:
            stats, findings = scan(f)
        except Exception as e:
            print(f"!! 读取失败 {os.path.basename(f)}: {e}")
            continue
        print(render(stats, findings, a.all))
        all_results.append({"stats": stats, "findings": findings})

    # 汇总
    print("=" * 72)
    print("【全班汇总】")
    print(f"{'文件':<44}{'P0':>4}{'P1':>4}{'P2':>4}{'图':>5}{'表':>5}")
    print("-" * 72)
    for r in all_results:
        s = r["stats"]
        print(f"{s['file'][:42]:<44}{s['p0']:>4}{s['p1']:>4}{s['p2']:>4}"
              f"{s['images_total']:>5}{s['tables']:>5}")
    print("-" * 72)
    print(f"{'合计':<44}"
          f"{sum(r['stats']['p0'] for r in all_results):>4}"
          f"{sum(r['stats']['p1'] for r in all_results):>4}"
          f"{sum(r['stats']['p2'] for r in all_results):>4}")

    if a.json:
        with open(a.json, "w", encoding="utf-8") as fp:
            json.dump({"scanned_at": datetime.now().isoformat(timespec="seconds"),
                       "target": a.target, "results": all_results},
                      fp, ensure_ascii=False, indent=2)
        print(f"\nJSON 已写入：{a.json}")


if __name__ == "__main__":
    main()