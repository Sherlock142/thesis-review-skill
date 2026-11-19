#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
两张xlsx 交叉同步工具（选题一览表 ↔ 答辩意见）

学生端唯一数据源：选题一览表.xlsx（根目录）。
IN-4_学生信息/学生信息.xlsx 已停用（学生+选题信息统一至一览表）。

规则（来自工作区约定，不可违背）
------------------------------
  以「学号」为唯一匹配键

  场景 A：写入答辩意见表时
    → 在 选题一览表 中查找该学号
    → 未找到       → 报告「需先在一览表登记」（不自动追加，一览表是唯一源）
    → 找到但题目为空 → 自动填充题目

  场景 B：写入一览表时
    → 在 答辩意见表 中查找该学号
    → 找到但姓名为空 → 自动填充姓名
    → 找到但题目为空 → 自动填充题目

  冲突检测（必须先问再改）
    学号匹配 + 两表题目都非空 + 值不一致
    → 🚫 禁止自动覆盖！打印冲突清单并退出，等人工确认

用法
----
  # 同步（自动补空字段，遇冲突则停下报告；学生端默认为一览表）
  python excel_双表同步.py

  # 只看差异，不写
  python excel_双表同步.py --dry-run

  # 导出冲突清单供人工核对
  python excel_双表同步.py --冲突报告 冲突.csv

  # 覆盖学生端路径（一般不需要）
  python excel_双表同步.py --学生表 <path.xlsx>

环境
----
  ⚠️ 需 openpyxl：pip install openpyxl
     用装有 openpyxl 的 Python 解释器运行即可
"""

import os
import sys
import csv
import argparse
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment
except ImportError:
    print("缺少 openpyxl。请先安装：pip install openpyxl")
    sys.exit(1)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
F_STU = os.path.join(ROOT, "选题一览表.xlsx")  # 学生端唯一数据源（默认）
F_OPN = os.path.join(ROOT, "IN-3_答辩组意见", "答辩意见.xlsx")

# 一览表表头（顺序即列序）
STU_COLS = ["题号", "正式题名", "班级", "姓名", "学号"]
OPN_COLS = ["序号", "学号", "姓名", "论文题目", "答辩意见内容",
            "意见来源", "提交日期", "备注"]


def norm_id(x):
    """学号归一化：去空格、转 str、全角数字转半角"""
    if x is None:
        return ""
    s = str(x).strip().replace("　", "")
    trans = str.maketrans("０１２３４５６７８９", "0123456789")
    s = s.translate(trans)
    # 明显的占位符（XX / 空）视为无学号
    if not s or re.fullmatch(r"[xXＸｘ]{2,}", s):
        return ""
    return s


def norm_txt(x):
    if x is None:
        return ""
    return str(x).strip()


import re  # noqa: E402  （放在 norm_id 之后以免上面的函数引用不到）


def read_sheet(path, cols):
    if not os.path.exists(path):
        print(f"!! 表不存在：{path}")
        return None, []
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    # 定位表头行（允许前几行有标题横幅）
    hdr_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [norm_txt(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if any(v == "学号" for v in vals):
            hdr_row = r
            break
    if hdr_row is None:
        hdr_row = 1
    hdrs = [norm_txt(ws.cell(hdr_row, c).value) for c in range(1, ws.max_column + 1)]
    idx = {h: i for i, h in enumerate(hdrs) if h}
    rows = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        rec = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if all(norm_txt(v) == "" for v in rec):
            continue
        rows.append({"row": r, "cells": rec, "idx": idx})
    return (ws, hdr_row), rows


def main():
    ap = argparse.ArgumentParser(description="两张 xlsx 交叉同步（学号为唯一键）")
    ap.add_argument("--dry-run", action="store_true", help="只报告差异不写文件")
    ap.add_argument("--冲突报告", help="把冲突清单导出为 CSV")
    ap.add_argument("--学生表", help="覆盖学生端表格路径（默认：选题一览表.xlsx）")
    a = ap.parse_args()
    f_stu = a.学生表 or F_STU

    print("=" * 68)
    print("双表交叉同步 —— 学号为唯一匹配键")
    print("=" * 68)

    stu_meta, stu_rows = read_sheet(f_stu, STU_COLS)
    opn_meta, opn_rows = read_sheet(F_OPN, OPN_COLS)
    if stu_meta is None or opn_meta is None:
        sys.exit(1)

    # 建索引
    def g(rec, key):
        i = rec["idx"].get(key)
        if i is None:
            return ""
        return norm_txt(rec["cells"][i]) if i < len(rec["cells"]) else ""

    def g_topic(rec):
        """题目列兼容：一览表叫「正式题名」，旧表叫「论文题目（中文）/论文题目」"""
        for k in ("正式题名", "论文题目（中文）", "论文题目"):
            v = g(rec, k)
            if v:
                return v
        return ""

    n_with_id = sum(1 for r in stu_rows if norm_id(g(r, "学号")))
    print(f"选题一览表：{len(stu_rows)} 行（已登记学生 {n_with_id} 人） | 答辩意见表：{len(opn_rows)} 条\n")

    stu_by_id = {}
    for r in stu_rows:
        sid = norm_id(g(r, "学号"))
        if sid:
            stu_by_id.setdefault(sid, []).append(r)

    opn_by_id = {}
    for r in opn_rows:
        sid = norm_id(g(r, "学号"))
        if sid:
            opn_by_id.setdefault(sid, []).append(r)

    conflicts, fills, newrows = [], [], []

    # 冲突检测
    for sid, orecs in opn_by_id.items():
        srecs = stu_by_id.get(sid)
        if not srecs:
            continue
        t_opn = g(orecs[0], "论文题目")
        for srec in srecs:
            t_stu = g_topic(srec)
            if t_opn and t_stu and t_opn != t_stu:
                conflicts.append({
                    "学号": sid,
                    "一览表题目": t_stu,
                    "答辩意见表题目": t_opn,
                    "姓名(信息表)": g(srec, "姓名"),
                    "姓名(意见表)": g(orecs[0], "姓名"),
                })

    if conflicts:
        print("🚫 检测到论文题目冲突 —— 已停止自动同步，需人工确认\n")
        print(f"{'学号':<16}{'一览表题目':<40}{'意见表题目':<40}")
        print("-" * 96)
        for c in conflicts:
            print(f"{c['学号']:<16}{c['一览表题目'][:38]:<40}{c['答辩意见表题目'][:38]:<40}")
        print("\n规则：两表题目都非空且不一致时禁止自动覆盖。")
        print("请人工确认以哪个为准，再手动修改。")
        if a.冲突报告:
            with open(a.冲突报告, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=list(conflicts[0].keys()))
                w.writeheader()
                w.writerows(conflicts)
            print(f"\n冲突清单已导出：{a.冲突报告}")
        sys.exit(2)

    print("✅ 无题目冲突，可安全同步\n")

    # 补空字段
    for sid, orecs in opn_by_id.items():
        name = g(orecs[0], "姓名")
        topic = g(orecs[0], "论文题目")
        srecs = stu_by_id.get(sid)
        if not srecs:
            newrows.append({"来源": "答辩意见表", "学号": sid, "姓名": name,
                            "论文题目": topic})
            continue
        srec = srecs[0]
        if not g(srec, "姓名") and name:
            fills.append(("选题一览表", sid, "姓名", "", name))
            if not a.dry_run:
                srec["cells"][srec["idx"]["姓名"]] = name
        st = g_topic(srec)
        if not st and topic:
            fills.append(("选题一览表", sid, "题目", "", topic))
            if not a.dry_run:
                key = "正式题名" if "正式题名" in srec["idx"] else (
                    "论文题目（中文）" if "论文题目（中文）" in srec["idx"] else "论文题目")
                srec["cells"][srec["idx"][key]] = topic

    for sid, srecs in stu_by_id.items():
        srec = srecs[0]
        name = g(srec, "姓名")
        topic = g_topic(srec)
        orecs = opn_by_id.get(sid)
        if not orecs:
            continue
        orec = orecs[0]
        if not g(orec, "姓名") and name:
            fills.append(("答辩意见表", sid, "姓名", "", name))
            if not a.dry_run:
                orec["cells"][orec["idx"]["姓名"]] = name
        if not g(orec, "论文题目") and topic:
            fills.append(("答辩意见表", sid, "论文题目", "", topic))
            if not a.dry_run:
                orec["cells"][orec["idx"]["论文题目"]] = topic

    print(f"【补空字段】{len(fills)} 处")
    for tbl, sid, field, old, new in fills:
        print(f"  {tbl}  {sid}  {field}：{'(空)' if not old else old} → {new}")

    print(f"\n【需先在一览表登记】{len(newrows)} 人（学号不在选题一览表中）")
    for n in newrows:
        print(f"  来自答辩意见表  {n['学号']}  {n['姓名']}  {n['论文题目'][:36]}")
    if newrows:
        print("  → 一览表是学生+选题唯一数据源：请先在 选题一览表.xlsx 登记，再重跑本脚本")

    if a.dry_run:
        print("\n[dry-run] 未写入任何文件")
        return

    # 写入（一览表不自动追加行——唯一数据源，新增学生必须先在一览表登记）
    if fills:
        touched = {t[0] for t in fills}
        for path, meta, rows, tbl in ((f_stu, stu_meta, stu_rows, "选题一览表"),
                                      (F_OPN, opn_meta, opn_rows, "答辩意见表")):
            if tbl not in touched:
                continue
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            for rec in rows:
                for i, v in enumerate(rec["cells"]):
                    if v is not None:
                        ws.cell(rec["row"], i + 1, v)
            wb.save(path)
            print(f"✓ 已写入 {os.path.basename(path)}")

    print("\n完成。下次新增数据后重跑本脚本即可。")


if __name__ == "__main__":
    main()