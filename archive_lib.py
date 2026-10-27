# -*- coding: utf-8 -*-
"""
电子会计档案 —— 本地档案库
==========================
把发票/回单/报表等原始凭证归档到本机 档案库/年/月/ 目录:
  · 归档时计算 SHA-256 防篡改指纹,之后可随时「验真」(重新计算比对)
  · 完全相同的文件拒绝重复归档;发票号重复的档案给出预警
  · 发票类档案自动识别元数据(发票号/金额/日期),支持检索
"""
import datetime
import hashlib
import json
import shutil
from pathlib import Path

import engine
import recognize

ARCHIVE_DIR = engine.BASE / "档案库"
INDEX_FILE = ARCHIVE_DIR / "档案索引.json"


def load_index():
    if INDEX_FILE.exists():
        try:
            return json.loads(INDEX_FILE.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _save_index(idx):
    ARCHIVE_DIR.mkdir(exist_ok=True)
    INDEX_FILE.write_text(json.dumps(idx, ensure_ascii=False, indent=2), encoding="utf-8")


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def archive_file(src_path, original_name, note=""):
    """归档一份文件;返回 (档案条目, 错误消息)"""
    idx = load_index()
    digest = sha256_of(src_path)
    same = [e for e in idx if e["指纹"] == digest]
    if same:
        return None, "内容与档案 %s(%s)完全相同,已拒绝重复归档" % (same[0]["档案号"], same[0]["原文件名"])

    ext = Path(original_name).suffix.lower()
    meta = {"发票号": "", "开票日期": "", "金额": None, "费用类型": ""}
    if ext == ".pdf" or ext in recognize.IMAGE_EXTS:
        try:
            f = recognize.recognize(src_path)
            for k in meta:
                meta[k] = f.get(k) or meta[k]
        except Exception:
            pass

    now = datetime.datetime.now()
    seq = 1 + sum(1 for e in idx if e["档案号"].startswith("DA-%s" % now.strftime("%Y%m")))
    no = "DA-%s-%03d" % (now.strftime("%Y%m"), seq)
    folder = ARCHIVE_DIR / now.strftime("%Y") / now.strftime("%m")
    folder.mkdir(parents=True, exist_ok=True)
    dest = folder / ("%s_%s" % (no, Path(original_name).name))
    shutil.copy2(src_path, dest)

    warn = ""
    if meta["发票号"] and any(e["发票号"] == meta["发票号"] for e in idx):
        warn = "发票号与已有档案重复"
    entry = {
        "档案号": no, "原文件名": Path(original_name).name,
        "路径": str(dest.relative_to(ARCHIVE_DIR)), "指纹": digest,
        "发票号": meta["发票号"], "金额": meta["金额"], "开票日期": meta["开票日期"],
        "类别": meta["费用类型"] or ("表格" if ext in (".xlsx", ".xls", ".csv") else "其他"),
        "归档时间": now.strftime("%Y-%m-%d %H:%M"), "备注": note, "预警": warn,
    }
    idx.append(entry)
    _save_index(idx)
    return entry, None


def find(no):
    for e in load_index():
        if e["档案号"] == no:
            return e
    return None


def file_path(entry):
    return ARCHIVE_DIR / entry["路径"]


def verify(no):
    """验真:重新计算指纹与归档时比对,判断档案是否被篡改"""
    e = find(no)
    if e is None:
        return False, "找不到档案 %s" % no
    p = file_path(e)
    if not p.exists():
        return False, "档案文件已丢失:%s" % e["路径"]
    ok = sha256_of(p) == e["指纹"]
    return ok, ("指纹一致,文件自归档以来未被改动" if ok else "⚠ 指纹不一致!文件内容与归档时不同,可能被篡改")


def stats():
    idx = load_index()
    month = datetime.datetime.now().strftime("%Y%m")
    total_amt = sum(e["金额"] for e in idx if isinstance(e.get("金额"), (int, float)))
    return {
        "总数": len(idx),
        "本月": sum(1 for e in idx if e["档案号"].startswith("DA-%s" % month)),
        "金额合计": round(total_amt, 2),
        "预警数": sum(1 for e in idx if e.get("预警")),
    }
