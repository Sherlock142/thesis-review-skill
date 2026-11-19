#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 批注批量注入器 —— 跨年通用版

原理
----
解包 DOCX → 在 word/document.xml 中，为「文本搜索定位到的段落」插入
commentRangeStart / commentRangeEnd / commentReference 标记 →
写入 word/comments.xml → 重新打包。

⚠️ 铁律：**绝不能用 python-docx 的段落索引做 XML 定位**。
   2026 届实测偏差最大 440 段落。本脚本一律走文本内容搜索。

批注定义文件（JSON）
-------------------
[
  {"find": "5系统实现", "level": "P0",
   "comment": "章节编号与标题之间缺少空格，应改为「5 系统实现」。学院模板要求编号后空一格。"},
  {"find": "登陆", "level": "P2",
   "comment": "「登陆」应为「登录」。全文统一修改。",
   "all": true}
]

字段说明
--------
  find    必填。用于在正文里搜索的锚点文本（取一段足够唯一的片段）
  comment 必填。批注内容
  level   P0 / P1 / P2，仅用于报告中分类，可省略
  all     可选。true = 替换全文所有出现处；false/缺省 = 仅第一处
  nth     可选。指定第 n 处（从 1 开始），与 all 配合可只改第 2、3 处
  skip    可选。锚点文本中已有的文字，用它做更精确定位（正则可选）

用法
----
  # 预览（不写文件，只报告哪些批注能命中）
  python docx_批注.py 论文.docx 批注.json --dry-run

  # 生成批注版
  python docx_批注.py 论文.docx 批注.json -o 论文-批注版.docx

依赖：lxml（必需）；python-docx（可选，仅校验用）
"""

import os
import re
import sys
import json
import shutil
import zipfile
import argparse
import tempfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def q(tag):
    return "{%s}%s" % (W, tag)


def para_text(p):
    """提取段落纯文本（拼接所有 w:t）"""
    return "".join(t.text or "" for t in p.iter(q("t")))


def find_paragraph(root, anchor, skip=None):
    """按文本搜索定位段落 —— 核心方法，不用索引"""
    cands = []
    for p in root.iter(q("p")):
        txt = para_text(p)
        if not txt:
            continue
        if skip:
            probe = skip if skip in txt else txt
        else:
            probe = txt
        if anchor in probe:
            cands.append(p)
    return cands


def make_comment(cid, author, initials, text):
    """构造 <w:comment> 元素"""
    c = etree.SubElement(etree.Element("dummy"), q("comment"))
    c.set(q("id"), str(cid))
    c.set(q("author"), author)
    c.set(q("initials"), initials)
    c.set(q("date"), "2026-10-13T00:00:00Z")
    p = etree.SubElement(c, q("p"))
    pPr = etree.SubElement(p, q("pPr"))
    ps = etree.SubElement(pPr, q("pStyle"))
    ps.set(q("val"), "CommentText")
    r = etree.SubElement(p, q("r"))
    rPr = etree.SubElement(r, q("rPr"))
    rs = etree.SubElement(rPr, q("rStyle"))
    rs.set(q("val"), "CommentReference")
    ann = etree.SubElement(r, q("annotationRef"))
    r2 = etree.SubElement(p, q("r"))
    t = etree.SubElement(r2, q("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return c


def insert_markers(p, cid):
    """
    在段落 p 内插入批注范围标记：
      <w:commentRangeStart w:id=N/> ...段落内容...
      <w:commentRangeEnd w:id=N/><w:r><w:commentReference w:id=N/></w:r>
    """
    start = etree.Element(q("commentRangeStart"))
    start.set(q("id"), str(cid))
    end = etree.Element(q("commentRangeEnd"))
    end.set(q("id"), str(cid))
    r = etree.Element(q("r"))
    rPr = etree.SubElement(r, q("rPr"))
    rs = etree.SubElement(rPr, q("rStyle"))
    rs.set(q("val"), "CommentReference")
    ref = etree.SubElement(r, q("commentReference"))
    ref.set(q("id"), str(cid))

    # 插到段落第一个 run 之前
    first_run = None
    for child in p:
        if child.tag == q("r"):
            first_run = child
            break
    if first_run is not None:
        first_run.addprevious(start)
    else:
        p.append(start)
    p.append(end)
    p.append(r)


COMMENTS_SKELETON = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
{body}
</w:comments>"""


def process(src, spec_file, out=None, author="王吉振", initials="王",
            dry_run=False):
    with open(spec_file, encoding="utf-8") as f:
        specs = json.load(f)

    workdir = tempfile.mkdtemp(prefix="docx_anno_")
    with zipfile.ZipFile(src) as z:
        z.extractall(workdir)

    doc_path = os.path.join(workdir, "word", "document.xml")
    tree = etree.parse(doc_path)
    root = tree.getroot()

    # 收集已有批注，决定新 ID 起点
    existing = 0
    cpath = os.path.join(workdir, "word", "comments.xml")
    if os.path.exists(cpath):
        try:
            croot = etree.parse(cpath).getroot()
            existing = len(croot.findall(q("comment")))
        except Exception:
            existing = 0

    cid = existing
    applied, missed = [], []

    for i, sp in enumerate(specs):
        anchor = sp.get("find")
        text = sp.get("comment")
        level = sp.get("level", "P2")
        do_all = sp.get("all", False)
        nth = sp.get("nth")
        skip = sp.get("skip")

        if not anchor or not text:
            missed.append({"index": i, "reason": "spec 缺 find 或 comment"})
            continue

        cands = find_paragraph(root, anchor, skip)

        if not cands:
            missed.append({"index": i, "level": level,
                           "anchor": anchor, "reason": "全文未找到锚点文本"})
            continue

        targets = cands if do_all else cands[:1]
        if nth:
            targets = [cands[nth - 1]] if nth <= len(cands) else []

        if not targets:
            missed.append({"index": i, "anchor": anchor,
                           "reason": f"仅 {len(cands)} 处命中，指定第 {nth} 处不存在"})
            continue

        for p in targets:
            cid += 1
            c = make_comment(cid, author, initials, text)
            insert_markers(p, cid)
            applied.append({
                "id": cid, "level": level, "anchor": anchor,
                "matched_text": para_text(p)[:60],
            })

    print("=" * 68)
    print(f"批注定义 {len(specs)} 条 → 实际注入 {len(applied)} 条"
          f"，未命中 {len(missed)} 条")
    print("=" * 68)
    if applied:
        print("\n【已注入】")
        for a in applied:
            print(f"  [{a['level']}] #{a['id']} 锚点「{a['anchor'][:24]}」")
            print(f"         命中段落：{a['matched_text']}")
    if missed:
        print("\n【未命中 —— 需人工核对锚点】")
        for m in missed:
            print(f"  #{m['index']} 锚点「{m.get('anchor','')[:28]}」→ {m['reason']}")

    if dry_run:
        shutil.rmtree(workdir, ignore_errors=True)
        print("\n[dry-run] 未写入任何文件")
        return

    # 写 comments.xml
    body = "\n".join(
        etree.tostring(make_comment(a["id"], author, initials,
                                    next(s["comment"] for s in specs
                                         if s.get("find") == a["anchor"])),
                       encoding="unicode")
        for a in applied)
    with open(cpath, "w", encoding="utf-8") as f:
        f.write(COMMENTS_SKELETON.format(body=body))

    tree.write(doc_path, xml_declaration=True, encoding="UTF-8",
               standalone=True)

    # 补关系与内容类型
    rels_path = os.path.join(workdir, "word", "_rels", "document.xml.rels")
    if os.path.exists(rels_path):
        rt = etree.parse(rels_path)
        rroot = rt.getroot()
        RNS = "http://schemas.openxmlformats.org/package/2006/relationships"
        has = any(e.get("Target") == "comments.xml" for e in rroot)
        if not has:
            e = etree.SubElement(rroot, "{%s}Relationship" % RNS)
            e.set("Id", "rIdCommentsNew")
            e.set("Type", "http://schemas.openxmlformats.org/officeDocument/"
                          "2006/relationships/comments")
            e.set("Target", "comments.xml")
            rt.write(rels_path, xml_declaration=True,
                     encoding="UTF-8", standalone=True)

    ct_path = os.path.join(workdir, "[Content_Types].xml")
    if os.path.exists(ct_path):
        ct = etree.parse(ct_path)
        croot2 = ct.getroot()
        CNS = "http://schemas.openxmlformats.org/package/2006/content-types"
        has = any(e.get("PartName") == "/word/comments.xml" for e in croot2)
        if not has:
            e = etree.SubElement(croot2, "{%s}Override" % CNS)
            e.set("PartName", "/word/comments.xml")
            e.set("ContentType", "application/vnd.openxmlformats-officedocument"
                                 ".wordprocessingml.comments+xml")
            ct.write(ct_path, xml_declaration=True, encoding="UTF-8",
                     standalone=True)

    # 打包
    out = out or (os.path.splitext(src)[0] + "-批注版.docx")
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml 必须第一个
        z.write(ct_path, "[Content_Types].xml")
        for root_dir, dirs, files in os.walk(workdir):
            for fn in files:
                full = os.path.join(root_dir, fn)
                arc = os.path.relpath(full, workdir).replace("\\", "/")
                if arc == "[Content_Types].xml":
                    continue
                z.write(full, arc)

    shutil.rmtree(workdir, ignore_errors=True)
    print(f"\n✓ 已生成：{out}")


def main():
    ap = argparse.ArgumentParser(description="Word 批注批量注入器（文本搜索定位）")
    ap.add_argument("docx")
    ap.add_argument("spec", help="批注定义 JSON")
    ap.add_argument("-o", "--out")
    ap.add_argument("--author", default="王吉振")
    ap.add_argument("--initials", default="王")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    process(a.docx, a.spec, a.out, a.author, a.initials, a.dry_run)


if __name__ == "__main__":
    main()