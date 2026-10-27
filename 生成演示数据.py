# -*- coding: utf-8 -*-
"""生成一份内容更丰富的演示数据(120笔),埋入各类问题单据,用于体验审核智能体。
运行: ../venv/bin/python 生成演示数据.py  →  生成 数据/演示数据.xlsx"""
import random
from pathlib import Path
import pandas as pd

random.seed(42)
BASE = Path(__file__).resolve().parent
OUT_DIR = BASE / "数据"
OUT_DIR.mkdir(exist_ok=True)

names = {"市场部": ["李娜", "杨杰", "赵敏"], "销售部": ["王强", "周涛", "孙丽"],
         "行政部": ["陈静", "吴磊"], "技术部": ["张伟", "刘洋", "郑爽"], "财务部": ["钱芳"]}
people = [(n, d) for d, ns in names.items() for n in ns]
types = ["餐饮费", "交通费", "办公用品", "差旅费", "招待费"]
reasons = {"餐饮费": ["项目组聚餐", "加班工作餐", "部门例会用餐"],
           "交通费": ["客户拜访打车", "机场往返", "外出办事交通"],
           "办公用品": ["购买文具", "购买打印耗材", "购买办公设备配件"],
           "差旅费": ["出差住宿", "出差高铁票", "出差市内交通"],
           "招待费": ["客户接待", "供应商洽谈用餐", "合作方来访接待"]}
rows, used_inv, no = [], set(), 1000

def inv():
    global used_inv
    while True:
        x = "FP%06d" % random.randint(100000, 999999)
        if x not in used_inv:
            used_inv.add(x)
            return x

def day(weekday_only=False, weekend=False):
    while True:
        d = pd.Timestamp("2026-06-01") + pd.Timedelta(days=random.randint(0, 29))
        if weekend and d.weekday() >= 5: return d
        if weekday_only and d.weekday() < 5: return d
        if not weekend and not weekday_only: return d

def add(name, dept, ftype, amt, d, invno=None, reason=None):
    global no
    no += 1
    rows.append({"报销单号": "BX%d" % no, "员工姓名": name, "部门": dept,
                 "报销日期": d.strftime("%Y-%m-%d"), "费用类型": ftype,
                 "金额": round(amt, 2), "发票号": inv() if invno is None else invno,
                 "事由": reason or random.choice(reasons[ftype])})

# —— 95 笔正常单据 ——
caps = {"餐饮费": 280, "交通费": 180, "办公用品": 900, "差旅费": 1800, "招待费": 950}
for _ in range(95):
    name, dept = random.choice(people)
    ftype = random.choice(types)
    add(name, dept, ftype, random.uniform(20, caps[ftype] * 0.85), day(weekday_only=(ftype in ("交通费", "餐饮费"))))

# —— 问题单据 ——
dup = inv()                                                   # R01 重复发票
add("王强", "销售部", "餐饮费", 268, day(weekday_only=True), invno=dup, reason="项目组聚餐")
add("王强", "销售部", "餐饮费", 268, day(weekday_only=True), invno=dup, reason="部门例会用餐")
for i in range(3):                                            # R02 发票连号
    add("周涛", "销售部", "交通费", random.uniform(120, 180), day(weekday_only=True), invno="FP66801%d" % i, reason="客户拜访打车")
add("刘洋", "技术部", "招待费", 2380, day(weekday_only=True), reason="客户接待")            # R03 超标
add("郑爽", "技术部", "餐饮费", 296, day(weekday_only=True), reason="加班工作餐")            # R08 贴线
d = day(weekday_only=True)                                    # R04 拆分:同日三笔均<300 合计>300
for amt in (295, 288, 292):
    add("孙丽", "销售部", "餐饮费", amt, d, reason="部门聚餐分单")
add("张伟", "技术部", "办公用品", 6000, day(weekday_only=True), reason="购买办公设备配件")   # R05 整数大额
add("赵敏", "市场部", "交通费", 95, day(weekend=True), reason="周末加班打车")                # R06+大白话周末打车
add("钱芳", "财务部", "招待费", 780, day(weekday_only=True), reason="客户接待含酒水")        # 大白话:酒水>500 拦截
for amt in (450.0, 450.0, 450.0):                             # R10 同额多笔
    add("吴磊", "行政部", "办公用品", amt, day(weekday_only=True), reason="购买打印耗材")
for _ in range(10):                                           # R07 高频(李娜当月20+笔)
    add("李娜", "市场部", "交通费", random.uniform(30, 150), day(weekday_only=True))
r = {"报销单号": "BX9998", "员工姓名": "王强", "部门": "销售部", "报销日期": "2026-06-15",
     "费用类型": "餐饮费", "金额": 156.5, "发票号": "", "事由": "工作餐"}          # R11 缺发票号
rows.append(r)
rows.append({"报销单号": "BX9999", "员工姓名": "周涛", "部门": "销售部", "报销日期": "待补",
             "费用类型": "差旅费", "金额": "见附件", "发票号": "FP000001", "事由": "出差住宿"})  # R11 金额/日期无法解析

df = pd.DataFrame(rows).sample(frac=1, random_state=7).reset_index(drop=True)
out = OUT_DIR / "演示数据.xlsx"
df.to_excel(out, index=False)
print("已生成 %s,共 %d 笔" % (out, len(df)))
