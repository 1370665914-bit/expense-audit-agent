# -*- coding: utf-8 -*-
"""
AI 填单 —— 报销草稿管理
======================
「扔凭证替代填表单」:上传发票/截图/行程单 → 本地识别 → 生成报销草稿。
AI 只生成草稿,识别不出的字段留空并提示;补全与提交的责任在报销人。
草稿存在 数据/报销草稿.xlsx,可以直接用 Excel 打开补全,补全后一键送 AI 审批。
"""
import re
from pathlib import Path

import pandas as pd

import engine
import recognize

VOUCHER_DIR = engine.DATA_DIR / "凭证文件"
DRAFT_FILE = engine.DATA_DIR / "报销草稿.xlsx"
COLS = ["报销单号", "员工姓名", "部门", "报销日期", "费用类型", "金额",
        "发票号", "事由", "凭证文件", "识别方式", "识别备注"]


def load_drafts():
    if not DRAFT_FILE.exists():
        return []
    df = pd.read_excel(DRAFT_FILE).fillna("")
    for c in COLS:
        if c not in df.columns:
            df[c] = ""
    return df[COLS].to_dict("records")


def _save(rows):
    engine.DATA_DIR.mkdir(exist_ok=True)
    pd.DataFrame(rows, columns=COLS).to_excel(DRAFT_FILE, index=False)


def _next_no(rows):
    mx = 0
    for r in rows:
        m = re.match(r"CD(\d+)", str(r["报销单号"]))
        if m:
            mx = max(mx, int(m.group(1)))
    return "CD%03d" % (mx + 1)


def save_voucher(file_storage):
    """把上传的凭证文件存到本机 数据/凭证文件/,重名自动加序号"""
    VOUCHER_DIR.mkdir(parents=True, exist_ok=True)
    name = Path(file_storage.filename).name
    dest = VOUCHER_DIR / name
    i = 1
    while dest.exists():
        dest = VOUCHER_DIR / ("%s-%d%s" % (Path(name).stem, i, Path(name).suffix))
        i += 1
    file_storage.save(dest)
    return dest


def add_from_file(saved_path):
    """识别一份凭证 → 追加一行草稿;返回草稿行(含重复预警)"""
    rows = load_drafts()
    try:
        f = recognize.recognize(saved_path)
    except Exception as e:
        f = {"发票号": "", "开票日期": "", "金额": None, "商户": "",
             "费用类型": "", "识别方式": "失败", "识别备注": "识别失败:%s" % e}
    notes = [f.get("识别备注", "")]
    inv = f.get("发票号", "")
    if inv:
        ledger = engine.load_ledger()
        if inv in ledger:
            notes.append("⚠ 此发票已在「已报销台账」(历史单号 %s),疑似重复报销" % ledger[inv].get("单号", "?"))
        if any(str(r["发票号"]) == inv for r in rows):
            notes.append("⚠ 草稿中已有同号发票")
    row = {
        "报销单号": _next_no(rows), "员工姓名": "", "部门": "",
        "报销日期": f.get("开票日期", ""), "费用类型": f.get("费用类型", ""),
        "金额": f.get("金额", ""), "发票号": inv,
        "事由": f.get("商户", ""), "凭证文件": Path(saved_path).name,
        "识别方式": f.get("识别方式", ""), "识别备注": " ".join(n for n in notes if n),
    }
    rows.append(row)
    _save(rows)
    return row


def delete_draft(no):
    rows = [r for r in load_drafts() if str(r["报销单号"]) != str(no)]
    _save(rows)


def clear_drafts():
    _save([])


def draft_rel_path():
    """草稿文件相对 finance-tools 的路径(供设置「最近文件」送审)"""
    return str(DRAFT_FILE.relative_to(engine.ROOT))
