# -*- coding: utf-8 -*-
"""
审核引擎 —— AI 财务审核智能体的核心
================================
职责:读取报销明细 Excel → 多维规则检查 → 风险评分 → 按信任滑杆分级
     全程在本机运行,不联网、不调用任何外部 AI 服务。

设计思路(借鉴 Seal AI):
  1. 多维风险检查   —— 内置 10+ 条审计规则,每条给出人话理由
  2. 风险评分分级   —— 自动通过 / 需人工复核 / 拦截 三档
  3. 信任滑杆       —— 试用期/转正期/专家期,决定 AI 自主放行的范围
  4. 防幻觉兜底     —— 关键要素缺失时 AI 不猜测,强制转人工
  5. 大白话配规则   —— 财务在 Excel 里用中文写规则,引擎自动解析
"""
import json
import re
import hashlib
import datetime
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent        # 审核智能体 文件夹
ROOT = BASE.parent                            # finance-tools 文件夹
DATA_DIR = BASE / "数据"
SETTINGS_FILE = BASE / "设置.json"
REVIEW_FILE = BASE / "复核记录.json"
LEDGER_FILE = BASE / "已报销台账.json"
RULES_XLSX = BASE / "审核规则.xlsx"

DEFAULT_SETTINGS = {
    "信任滑杆": 20,      # 0-100:  <30 试用期 / 30-69 转正期 / >=70 专家期
    "低风险线": 25,      # 风险分 <= 低风险线 视为低风险
    "高风险线": 60,      # 风险分 >= 高风险线 直接拦截
    "抽检比例": 5,       # AI 放行单据中抽取 N% 提示人工抽检
    "最近文件": "",
}

# ---------------------------------------------------------------------------
# 基础:设置 与 复核记录 的读写
# ---------------------------------------------------------------------------

def load_settings():
    s = dict(DEFAULT_SETTINGS)
    if SETTINGS_FILE.exists():
        try:
            s.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
        except Exception:
            pass
    return s


def save_settings(new_values):
    s = load_settings()
    s.update(new_values)
    SETTINGS_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")
    return s


def load_ledger():
    """已报销台账:{发票号: {单号, 姓名, 金额, 日期, 来源文件, 登记时间}}"""
    if LEDGER_FILE.exists():
        try:
            return json.loads(LEDGER_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def register_ledger(result, reviews):
    """把一批审核结果中「放行」的发票登记入台账,防止跨批次重复报销。
    登记范围:自动通过、人工通过;人工拒绝和(未被人工翻案的)拦截不登记。"""
    ledger = load_ledger()
    added = 0
    for rec in result["记录"]:
        inv = rec["发票号"]
        if not inv or inv in ledger:
            continue
        rv = reviews.get("%s::%s" % (result["文件"], rec["报销单号"]), {})
        decision = rv.get("决定", "")
        if decision == "人工拒绝":
            continue
        if rec["等级"] != "自动通过" and decision != "人工通过":
            continue
        ledger[inv] = {
            "单号": rec["报销单号"], "姓名": rec["员工姓名"],
            "金额": rec["_金额数值"], "日期": rec["日期文本"],
            "来源文件": result["文件"],
            "登记时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        added += 1
    LEDGER_FILE.write_text(json.dumps(ledger, ensure_ascii=False, indent=2), encoding="utf-8")
    return added, len(ledger)


def load_reviews():
    if REVIEW_FILE.exists():
        try:
            return json.loads(REVIEW_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_review(file_name, voucher_no, decision, note=""):
    reviews = load_reviews()
    key = "%s::%s" % (file_name, voucher_no)
    reviews[key] = {
        "决定": decision,           # 人工通过 / 人工拒绝 / 撤销
        "备注": note,
        "时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    if decision == "撤销":
        reviews.pop(key, None)
    REVIEW_FILE.write_text(json.dumps(reviews, ensure_ascii=False, indent=2), encoding="utf-8")
    return reviews


# ---------------------------------------------------------------------------
# 数据文件发现 与 读取(列名自动映射,兼容不同财务软件导出格式)
# ---------------------------------------------------------------------------

COLUMN_ALIASES = {
    "报销单号": ["报销单号", "单号", "单据编号", "单据号", "凭证号"],
    "员工姓名": ["员工姓名", "姓名", "报销人", "申请人", "经办人"],
    "部门":     ["部门", "所属部门", "申请部门"],
    "报销日期": ["报销日期", "日期", "单据日期", "申请日期", "业务日期"],
    "费用类型": ["费用类型", "费用类别", "费用科目", "科目", "费用项目"],
    "金额":     ["金额", "报销金额", "金额(元)", "金额（元）", "价税合计"],
    "发票号":   ["发票号", "发票号码", "发票编号", "票据号"],
    "事由":     ["事由", "摘要", "备注", "说明", "报销事由"],
}


def list_data_files():
    """列出可审核的 Excel:finance-tools 根目录 + 审核智能体/数据 文件夹"""
    files = []
    for folder in (ROOT, DATA_DIR):
        if not folder.exists():
            continue
        for p in sorted(folder.glob("*.xlsx")):
            if p.name.startswith("~$") or p.name in ("审核规则.xlsx", "审核结果.xlsx", "异常清单.xlsx"):
                continue
            files.append(str(p.relative_to(ROOT)))
    return files


def resolve_data_path(rel_name):
    p = ROOT / rel_name
    return p if p.exists() else None


def load_data(path):
    """读取明细并把列名统一成标准名;返回 (DataFrame, 缺失的关键列列表)"""
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    rename = {}
    for std, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            if a in df.columns and std not in rename.values():
                rename[a] = std
                break
    df = df.rename(columns=rename)
    missing = [c for c in ("金额", "员工姓名") if c not in df.columns]
    for std in COLUMN_ALIASES:
        if std not in df.columns:
            df[std] = ""
    if (df["报销单号"] == "").all():
        df["报销单号"] = ["行%03d" % (i + 2) for i in range(len(df))]
    return df, missing


# ---------------------------------------------------------------------------
# 规则配置表(Excel):不存在时自动生成带示例的默认表 —— "写文档即配系统"
# ---------------------------------------------------------------------------

DEFAULT_LIMITS = [
    ("餐饮费", 300, "单人单笔餐饮限额"),
    ("交通费", 200, "市内交通单笔限额"),
    ("办公用品", 1000, "办公用品单笔限额"),
    ("差旅费", 2000, "差旅单笔限额"),
    ("招待费", 1500, "业务招待单笔限额"),
]

DEFAULT_PARAMS = [
    ("大额整数线", 5000, "金额达到此数且为整千,触发【可疑整数大额】"),
    ("当月高频笔数", 15, "同人当月报销超过此笔数,触发【高频报销】"),
    ("贴线比例", 90, "金额达到限额的此百分比但未超,触发【贴线报销】"),
    ("同额多笔笔数", 3, "同人相同金额出现达到此笔数,触发【同额多笔】"),
]

# (编号, 名称, 级别, 风险分, 说明)   级别: 拦截 / 复核 / 提示 / 兜底
BUILTIN_RULES = [
    ("R01", "发票号重复", "拦截", 100, "同一发票号在本批数据中出现多次,疑似重复报销"),
    ("R02", "发票连号",   "复核", 40,  "同一员工多张发票号码连续,疑似一次性代开"),
    ("R03", "单笔超标",   "复核", 35,  "单笔金额超过该费用类型限额(见【费用限额】表)"),
    ("R04", "拆分报销",   "复核", 45,  "同人同日同类多笔,单笔不超限但合计超限,疑似拆分发票"),
    ("R05", "可疑整数大额", "复核", 30, "金额达到大额线且为整千整百,真实消费很少如此"),
    ("R06", "非工作日消费", "提示", 15, "交通/餐饮费发生在周末"),
    ("R07", "高频报销",   "提示", 20,  "同人当月报销笔数超过阈值"),
    ("R08", "贴线报销",   "提示", 15,  "金额达到限额的90%但未超,疑似按限额凑票"),
    ("R10", "同额多笔",   "复核", 25,  "同人多笔完全相同的金额,疑似复制报销"),
    ("R11", "要素缺失",   "兜底", 0,   "关键信息缺失或无法解析 —— AI 不猜测,强制转人工(防幻觉兜底)"),
    ("R12", "历史已报销", "拦截", 100, "发票号已在「已报销台账」登记过,跨批次重复报销"),
]

DEFAULT_PLAIN_RULES = [
    ("事由 含 酒、酒水、烟、礼品 且 金额 超过 500", "拦截", "是", "含酒水字样超500元不予报销"),
    ("费用类型 含 招待 且 金额 超过 1000", "复核", "是", "大额招待重点复核"),
    ("事由 含 打车 且 日期 是 周末", "提示", "是", "周末打车关注是否加班"),
    ("部门 含 财务 且 金额 超过 3000", "提示", "否", "示例:某部门大额提示(未启用)"),
]


def ensure_rules_workbook():
    if RULES_XLSX.exists():
        return
    with pd.ExcelWriter(RULES_XLSX, engine="openpyxl") as w:
        pd.DataFrame(DEFAULT_LIMITS, columns=["费用类型", "单笔限额(元)", "说明"]).to_excel(
            w, sheet_name="费用限额", index=False)
        pd.DataFrame(DEFAULT_PLAIN_RULES, columns=["规则内容", "处理方式", "是否启用", "备注"]).to_excel(
            w, sheet_name="大白话规则", index=False)
        pd.DataFrame(
            [(r[0], r[1], r[3], "是", r[4]) for r in BUILTIN_RULES],
            columns=["规则编号", "规则名称", "风险分", "是否启用", "说明"],
        ).to_excel(w, sheet_name="内置规则", index=False)
        pd.DataFrame(DEFAULT_PARAMS, columns=["参数名", "值", "说明"]).to_excel(
            w, sheet_name="参数", index=False)


def load_rule_config():
    """读取 审核规则.xlsx → dict(限额表 / 参数 / 内置规则开关与分值 / 大白话规则)"""
    ensure_rules_workbook()
    cfg = {"限额": {}, "参数": {}, "内置": {}, "大白话": [], "解析警告": []}
    xls = pd.ExcelFile(RULES_XLSX)

    if "费用限额" in xls.sheet_names:
        for _, row in xls.parse("费用限额").iterrows():
            try:
                cfg["限额"][str(row["费用类型"]).strip()] = float(row["单笔限额(元)"])
            except Exception:
                continue

    for name, default, _ in DEFAULT_PARAMS:
        cfg["参数"][name] = default
    if "参数" in xls.sheet_names:
        for _, row in xls.parse("参数").iterrows():
            try:
                cfg["参数"][str(row["参数名"]).strip()] = float(row["值"])
            except Exception:
                continue

    for rid, name, level, score, desc in BUILTIN_RULES:
        cfg["内置"][rid] = {"名称": name, "级别": level, "风险分": score, "启用": True, "说明": desc}
    if "内置规则" in xls.sheet_names:
        for _, row in xls.parse("内置规则").iterrows():
            rid = str(row.get("规则编号", "")).strip()
            if rid in cfg["内置"]:
                try:
                    cfg["内置"][rid]["风险分"] = int(float(row["风险分"]))
                except Exception:
                    pass
                cfg["内置"][rid]["启用"] = str(row.get("是否启用", "是")).strip() != "否"

    if "大白话规则" in xls.sheet_names:
        for i, (_, row) in enumerate(xls.parse("大白话规则").iterrows(), start=1):
            text = str(row.get("规则内容", "")).strip()
            if not text or text == "nan":
                continue
            if str(row.get("是否启用", "是")).strip() == "否":
                continue
            level = str(row.get("处理方式", "复核")).strip()
            if level not in ("拦截", "复核", "提示"):
                level = "复核"
            checker, err = parse_plain_rule(text)
            if err:
                cfg["解析警告"].append("第%d条「%s」:%s" % (i, text, err))
                continue
            cfg["大白话"].append({
                "编号": "R09-%d" % i, "名称": "自定义:%s" % text, "级别": level,
                "风险分": {"拦截": 100, "复核": 35, "提示": 15}[level],
                "文本": text, "检查": checker,
                "备注": str(row.get("备注", "")).strip(),
            })
    return cfg


# ---------------------------------------------------------------------------
# 大白话规则解析器 —— 财务用中文写规则,秒级变成可执行检查
# 支持的写法(条件之间用「且」连接):
#   <字段> 含 <词1、词2…>      字段可为: 事由/费用类型/部门/员工姓名/发票号
#   金额 超过|大于|低于|小于 <数字>
#   日期 是 周末
# ---------------------------------------------------------------------------

def parse_plain_rule(text):
    conds = []
    for part in re.split(r"[且,，;；]", text):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(事由|摘要|费用类型|部门|员工姓名|姓名|发票号)\s*含\s*(.+)$", part)
        if m:
            field = {"摘要": "事由", "姓名": "员工姓名"}.get(m.group(1), m.group(1))
            words = [w.strip() for w in re.split(r"[、,，/|]", m.group(2)) if w.strip()]
            conds.append(("contains", field, words))
            continue
        m = re.match(r"^金额\s*(超过|大于|>=?|高于|低于|小于|<=?)\s*([0-9.]+)\s*元?$", part)
        if m:
            op = "gt" if m.group(1) in ("超过", "大于", ">", ">=", "高于") else "lt"
            conds.append(("amount", op, float(m.group(2))))
            continue
        if re.match(r"^日期\s*是\s*周末$", part):
            conds.append(("weekend",))
            continue
        return None, "看不懂「%s」。支持:字段 含 词 / 金额 超过 数字 / 日期 是 周末" % part
    if not conds:
        return None, "空规则"

    def checker(rec):
        for c in conds:
            if c[0] == "contains":
                val = str(rec.get(c[1], "") or "")
                if not any(w in val for w in c[2]):
                    return False
            elif c[0] == "amount":
                amt = rec.get("_金额数值")
                if amt is None:
                    return False
                if c[1] == "gt" and not amt > c[2]:
                    return False
                if c[1] == "lt" and not amt < c[2]:
                    return False
            elif c[0] == "weekend":
                d = rec.get("_日期对象")
                if d is None or d.weekday() < 5:
                    return False
        return True

    return checker, None


# ---------------------------------------------------------------------------
# 审核主流程
# ---------------------------------------------------------------------------

def _fmt_money(x):
    return ("{:,.2f}".format(x)).rstrip("0").rstrip(".") if x == x else "?"


def _invoice_parts(inv):
    """把发票号拆成 (前缀, 数字尾巴),用于连号检测"""
    m = re.match(r"^(.*?)(\d{3,})$", str(inv).strip())
    if not m:
        return None
    return m.group(1), int(m.group(2))


def run_audit(rel_name):
    """核心入口:返回包含每单结果与汇总统计的 dict(可直接渲染页面)"""
    settings = load_settings()
    cfg = load_rule_config()
    path = resolve_data_path(rel_name)
    if path is None:
        raise FileNotFoundError("找不到数据文件:%s" % rel_name)

    df, missing_cols = load_data(path)

    # ---- 逐行转成记录,并解析金额/日期(解析失败 = 触发防幻觉兜底) ----
    records = []
    for idx, row in df.iterrows():
        rec = {}
        for col in COLUMN_ALIASES:
            v = row.get(col, "")
            rec[col] = "" if pd.isna(v) else (str(v).strip() if not isinstance(v, str) else v.strip())
        try:
            rec["_金额数值"] = round(float(row.get("金额")), 2)
        except Exception:
            rec["_金额数值"] = None
        d = pd.to_datetime(row.get("报销日期"), errors="coerce")
        rec["_日期对象"] = None if pd.isna(d) else d
        rec["日期文本"] = "" if pd.isna(d) else d.strftime("%Y-%m-%d")
        rec["星期"] = "" if pd.isna(d) else "周" + "一二三四五六日"[d.weekday()]
        rec["命中"] = []
        records.append(rec)

    def enabled(rid):
        return cfg["内置"][rid]["启用"]

    def hit(rec, rid, detail):
        r = cfg["内置"][rid]
        rec["命中"].append({"编号": rid, "名称": r["名称"], "级别": r["级别"],
                            "风险分": r["风险分"], "详情": detail})

    limits = cfg["限额"]
    P = cfg["参数"]

    def limit_of(ftype):
        if ftype in limits:
            return limits[ftype]
        for k, v in limits.items():   # "餐饮费" 也能匹配 "餐饮"
            if k and (k in ftype or ftype in k):
                return v
        return None

    # ---- R01 发票号重复 ----
    inv_count = {}
    for rec in records:
        if rec["发票号"]:
            inv_count[rec["发票号"]] = inv_count.get(rec["发票号"], 0) + 1
    if enabled("R01"):
        for rec in records:
            n = inv_count.get(rec["发票号"], 0)
            if n > 1:
                hit(rec, "R01", "发票号 %s 在本批数据中出现 %d 次" % (rec["发票号"], n))

    # ---- R12 历史已报销(比对已报销台账,跨批次查重) ----
    if enabled("R12"):
        ledger = load_ledger()
        for rec in records:
            e = ledger.get(rec["发票号"])
            if e and e.get("单号") != rec["报销单号"]:
                hit(rec, "R12", "发票号 %s 已于 %s 由 %s 报销过(单号 %s,来源 %s)"
                    % (rec["发票号"], e.get("日期") or "?", e.get("姓名") or "?",
                       e.get("单号") or "?", e.get("来源文件") or "?"))

    # ---- R02 发票连号(同人同前缀,号码差1的相邻票) ----
    if enabled("R02"):
        by_person = {}
        for rec in records:
            parts = _invoice_parts(rec["发票号"])
            if parts:
                by_person.setdefault((rec["员工姓名"], parts[0]), []).append((parts[1], rec))
        for (_, _), items in by_person.items():
            items.sort(key=lambda t: t[0])
            run = [items[0]]
            for cur in items[1:]:
                if cur[0] - run[-1][0] == 1:
                    run.append(cur)
                else:
                    if len(run) >= 2:
                        for _, r in run:
                            hit(r, "R02", "与本人另外 %d 张发票号码连续(%s…)" % (len(run) - 1, run[0][1]["发票号"]))
                    run = [cur]
            if len(run) >= 2:
                for _, r in run:
                    hit(r, "R02", "与本人另外 %d 张发票号码连续(%s…)" % (len(run) - 1, run[0][1]["发票号"]))

    # ---- R03 单笔超标 / R08 贴线报销 ----
    for rec in records:
        amt, lim = rec["_金额数值"], limit_of(rec["费用类型"])
        if amt is None or lim is None:
            continue
        if enabled("R03") and amt > lim:
            hit(rec, "R03", "%s %s 元,超过限额 %s 元(超 %.0f%%)"
                % (rec["费用类型"], _fmt_money(amt), _fmt_money(lim), (amt / lim - 1) * 100))
        elif enabled("R08") and amt >= lim * P["贴线比例"] / 100:
            hit(rec, "R08", "%s %s 元,达到限额 %s 元的 %.0f%%,贴线未超"
                % (rec["费用类型"], _fmt_money(amt), _fmt_money(lim), amt / lim * 100))

    # ---- R04 拆分报销 ----
    if enabled("R04"):
        groups = {}
        for rec in records:
            if rec["_金额数值"] is not None and rec["日期文本"]:
                groups.setdefault((rec["员工姓名"], rec["日期文本"], rec["费用类型"]), []).append(rec)
        for (_, day, ftype), items in groups.items():
            lim = limit_of(ftype)
            if lim is None or len(items) < 2:
                continue
            total = sum(r["_金额数值"] for r in items)
            if total > lim and all(r["_金额数值"] <= lim for r in items):
                for r in items:
                    hit(r, "R04", "%s当日%s共 %d 笔、合计 %s 元,单笔均不超限但合计超过限额 %s 元"
                        % (day, ftype, len(items), _fmt_money(total), _fmt_money(lim)))

    # ---- R05 可疑整数大额 ----
    if enabled("R05"):
        for rec in records:
            amt = rec["_金额数值"]
            if amt is not None and amt >= P["大额整数线"] and amt % 1000 == 0:
                hit(rec, "R05", "金额 %s 元,达到大额线且为整千" % _fmt_money(amt))

    # ---- R06 非工作日消费 ----
    if enabled("R06"):
        for rec in records:
            d = rec["_日期对象"]
            if d is not None and d.weekday() >= 5 and any(k in rec["费用类型"] for k in ("交通", "餐饮")):
                hit(rec, "R06", "%s(%s)发生%s" % (rec["日期文本"], rec["星期"], rec["费用类型"]))

    # ---- R07 高频报销 ----
    if enabled("R07"):
        freq = {}
        for rec in records:
            if rec["_日期对象"] is not None:
                freq.setdefault((rec["员工姓名"], rec["日期文本"][:7]), []).append(rec)
        for (name, month), items in freq.items():
            if len(items) > P["当月高频笔数"]:
                for r in items:
                    hit(r, "R07", "%s 在 %s 月共报销 %d 笔,超过 %d 笔阈值" % (name, month, len(items), int(P["当月高频笔数"])))

    # ---- R10 同额多笔 ----
    if enabled("R10"):
        same = {}
        for rec in records:
            amt = rec["_金额数值"]
            if amt is not None and amt >= 100:
                same.setdefault((rec["员工姓名"], amt), []).append(rec)
        for (name, amt), items in same.items():
            if len(items) >= P["同额多笔笔数"]:
                for r in items:
                    hit(r, "R10", "%s 有 %d 笔金额完全相同(%s 元)" % (name, len(items), _fmt_money(amt)))

    # ---- R09 大白话规则 ----
    for rule in cfg["大白话"]:
        for rec in records:
            if rule["检查"](rec):
                rec["命中"].append({"编号": rule["编号"], "名称": rule["名称"], "级别": rule["级别"],
                                    "风险分": rule["风险分"],
                                    "详情": "命中自定义规则「%s」%s" % (rule["文本"], ("——" + rule["备注"]) if rule["备注"] else "")})

    # ---- R11 要素缺失(防幻觉兜底) ----
    for rec in records:
        problems = []
        if rec["_金额数值"] is None:
            problems.append("金额无法解析")
        if not rec["员工姓名"]:
            problems.append("缺员工姓名")
        if rec["_日期对象"] is None:
            problems.append("日期无法解析")
        if not rec["发票号"]:
            problems.append("缺发票号")
        rec["_要素缺失"] = bool(problems)
        if problems and enabled("R11"):
            hit(rec, "R11", "、".join(problems) + " —— AI 不基于猜测放行,强制转人工")

    # ---- 评分与分级(信任滑杆在这里生效) ----
    T = int(settings["信任滑杆"])
    stage = "试用期" if T < 30 else ("转正期" if T < 70 else "专家期")
    low, high, sample_pct = settings["低风险线"], settings["高风险线"], settings["抽检比例"]

    for rec in records:
        score = sum(h["风险分"] for h in rec["命中"])
        blocked = any(h["级别"] == "拦截" for h in rec["命中"]) or score >= high
        rec["风险分"] = score
        rec["抽检"] = False
        if blocked:
            rec["等级"] = "拦截"
        elif rec["_要素缺失"]:
            rec["等级"] = "需人工复核"
        elif score == 0:
            rec["等级"] = "自动通过"
        else:
            if T < 30:
                ai_pass = False
            elif T < 70:
                ai_pass = score <= low
            else:
                ai_pass = score < high
            rec["等级"] = "自动通过" if ai_pass else "需人工复核"
            if ai_pass:
                rec["AI放行"] = True
                digest = int(hashlib.md5(rec["报销单号"].encode("utf-8")).hexdigest(), 16)
                rec["抽检"] = digest % 100 < sample_pct
        rec.setdefault("AI放行", False)
        rec["意见"] = _build_opinion(rec, stage)

    # ---- 汇总统计 ----
    result = {
        "文件": rel_name,
        "运行时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "信任滑杆": T, "阶段": stage,
        "解析警告": cfg["解析警告"], "缺失列": missing_cols,
        "记录": records,
        "统计": _build_stats(records),
    }
    return result


def _build_opinion(rec, stage):
    """把命中的规则组装成一段人话审核意见 —— 这就是『AI 审核意见』"""
    hits = rec["命中"]
    if not hits:
        return "未命中任何风险规则,单据要素完整,自动通过。"
    lines = []
    for i, h in enumerate(hits, 1):
        lines.append("%d. 【%s】%s" % (i, h["名称"], h["详情"]))
    body = "命中 %d 项风险:\n%s\n综合风险分 %d。" % (len(hits), "\n".join(lines), rec["风险分"])
    if rec["等级"] == "拦截":
        return body + "存在重大风险,已拦截,需报销人补充说明后由财务人工处理。"
    if rec["等级"] == "需人工复核":
        if rec["_要素缺失"]:
            return body + "关键要素缺失,按防幻觉兜底机制强制转人工,AI 不作猜测。"
        return body + "按当前信任等级(%s)转人工复核。" % stage
    extra = "已列入人工抽检名单。" if rec["抽检"] else "人工可随时抽检。"
    return body + "风险较低,按当前信任等级(%s)由 AI 放行。%s" % (stage, extra)


def _build_stats(records):
    n = len(records)
    by_level = {"自动通过": 0, "需人工复核": 0, "拦截": 0}
    ai_pass = sample = 0
    total_amt = 0.0
    rule_hits = {}
    dept = {}
    for rec in records:
        by_level[rec["等级"]] += 1
        if rec["AI放行"]:
            ai_pass += 1
        if rec["抽检"]:
            sample += 1
        if rec["_金额数值"]:
            total_amt += rec["_金额数值"]
        seen = set()
        for h in rec["命中"]:
            key = h["名称"]
            if key not in seen:
                rule_hits[key] = rule_hits.get(key, 0) + 1
                seen.add(key)
        d = rec["部门"] or "(未填部门)"
        dept.setdefault(d, {"总数": 0, "风险单": 0, "风险分": 0})
        dept[d]["总数"] += 1
        dept[d]["风险分"] += rec["风险分"]
        if rec["等级"] != "自动通过":
            dept[d]["风险单"] += 1
    top_rules = sorted(rule_hits.items(), key=lambda kv: -kv[1])[:8]
    dept_rank = sorted(
        [{"部门": k, "总数": v["总数"], "风险单": v["风险单"], "风险分": v["风险分"]} for k, v in dept.items()],
        key=lambda d: -d["风险分"])[:8]
    return {
        "总单数": n, "总金额": round(total_amt, 2),
        "分布": by_level, "AI放行": ai_pass, "抽检": sample,
        "自动化率": round(by_level["自动通过"] * 100.0 / n, 1) if n else 0.0,
        "规则Top": top_rules, "部门排行": dept_rank,
    }


# ---------------------------------------------------------------------------
# 导出审核结果 Excel
# ---------------------------------------------------------------------------

def export_result(result, out_path, reviews=None):
    reviews = reviews or {}
    rows = []
    for rec in result["记录"]:
        key = "%s::%s" % (result["文件"], rec["报销单号"])
        rv = reviews.get(key, {})
        rows.append({
            "报销单号": rec["报销单号"], "员工姓名": rec["员工姓名"], "部门": rec["部门"],
            "报销日期": rec["日期文本"] or rec["报销日期"], "费用类型": rec["费用类型"],
            "金额": rec["_金额数值"] if rec["_金额数值"] is not None else rec["金额"],
            "发票号": rec["发票号"], "事由": rec["事由"],
            "风险分": rec["风险分"], "审核结论": rec["等级"],
            "AI放行": "是" if rec["AI放行"] else "", "抽检": "是" if rec["抽检"] else "",
            "命中规则": ";".join(h["名称"] for h in rec["命中"]),
            "AI审核意见": rec["意见"].replace("\n", " "),
            "人工决定": rv.get("决定", ""), "人工备注": rv.get("备注", ""),
        })
    df = pd.DataFrame(rows)
    order = ["拦截", "需人工复核", "自动通过"]
    df["_排序"] = df["审核结论"].map({v: i for i, v in enumerate(order)})
    df = df.sort_values(["_排序", "风险分"], ascending=[True, False]).drop(columns="_排序")
    df.to_excel(out_path, index=False)
    return out_path


if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "报销明细.xlsx"
    res = run_audit(target)
    s = res["统计"]
    print("文件:", res["文件"], "| 阶段:", res["阶段"], "| 信任滑杆:", res["信任滑杆"])
    print("总单数 %d | 自动通过 %d(其中AI放行 %d) | 需人工复核 %d | 拦截 %d | 自动化率 %.1f%%"
          % (s["总单数"], s["分布"]["自动通过"], s["AI放行"], s["分布"]["需人工复核"], s["分布"]["拦截"], s["自动化率"]))
    print("规则命中Top:", s["规则Top"])
    if res["解析警告"]:
        print("大白话规则解析警告:", res["解析警告"])
    for rec in res["记录"]:
        if rec["等级"] != "自动通过":
            print("-", rec["报销单号"], rec["员工姓名"], rec["等级"], rec["风险分"],
                  "|", ";".join(h["名称"] for h in rec["命中"]))
