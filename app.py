# -*- coding: utf-8 -*-
"""
AI 财务审核智能体 —— 本地审核工作台
只监听本机 127.0.0.1,数据全程不出这台电脑。
启动: ../venv/bin/python app.py   (或双击 启动审核台.command)
"""
import datetime
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import (Flask, jsonify, redirect, render_template, request,
                   send_file, url_for)

import archive_lib
import drafts
import engine

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
PORT = 5017   # 避开 macOS 隔空播放占用的 5000 端口


def stage_of(settings):
    t = int(settings["信任滑杆"])
    return "试用期" if t < 30 else ("转正期" if t < 70 else "专家期")


@app.route("/")
def index():
    settings = engine.load_settings()
    return render_template("index.html", files=engine.list_data_files(),
                           settings=settings, stage=stage_of(settings),
                           rules_path=str(engine.RULES_XLSX), active="index")


@app.route("/run", methods=["POST"])
def run():
    f = request.files.get("upload")
    if f and f.filename:
        engine.DATA_DIR.mkdir(exist_ok=True)
        dest = engine.DATA_DIR / Path(f.filename).name
        f.save(dest)
        rel = str(dest.relative_to(engine.ROOT))
    else:
        rel = request.form.get("file", "")
    if not rel or engine.resolve_data_path(rel) is None:
        return redirect(url_for("index"))
    engine.save_settings({"最近文件": rel})
    return redirect(url_for("dashboard"))


@app.route("/dashboard")
def dashboard():
    settings = engine.load_settings()
    rel = settings.get("最近文件")
    if not rel or engine.resolve_data_path(rel) is None:
        return redirect(url_for("index"))
    result = engine.run_audit(rel)
    reviews = engine.load_reviews()
    for rec in result["记录"]:
        rec["人工"] = reviews.get("%s::%s" % (rel, rec["报销单号"]))
    order = {"拦截": 0, "需人工复核": 1, "自动通过": 2}
    result["记录"].sort(key=lambda r: (order[r["等级"]], -r["风险分"]))
    return render_template("dashboard.html", r=result, s=result["统计"],
                           settings=settings, active="dashboard")


@app.route("/review", methods=["POST"])
def review():
    data = request.get_json(force=True)
    settings = engine.load_settings()
    engine.save_review(settings.get("最近文件", ""), data.get("单号", ""),
                       data.get("决定", ""), data.get("备注", ""))
    return jsonify({"ok": True})


@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    if request.method == "POST":
        try:
            engine.save_settings({
                "信任滑杆": max(0, min(100, int(request.form.get("信任滑杆", 20)))),
                "低风险线": int(request.form.get("低风险线", 25)),
                "高风险线": int(request.form.get("高风险线", 60)),
                "抽检比例": max(0, min(100, int(request.form.get("抽检比例", 5)))),
            })
        except ValueError:
            pass
        return redirect(url_for("settings_page", saved=1))
    st = engine.load_settings()
    return render_template("settings.html", settings=st, stage=stage_of(st),
                           saved=request.args.get("saved"), active="settings")


@app.route("/export")
def export():
    settings = engine.load_settings()
    rel = settings.get("最近文件")
    if not rel or engine.resolve_data_path(rel) is None:
        return redirect(url_for("index"))
    result = engine.run_audit(rel)
    out = engine.BASE / "审核结果.xlsx"
    engine.export_result(result, out, engine.load_reviews())
    return send_file(out, as_attachment=True,
                     download_name="审核结果_%s.xlsx" % datetime.date.today().strftime("%Y%m%d"))


@app.route("/ledger", methods=["POST"])
def ledger():
    """把本批放行的发票登记入「已报销台账」,供跨批次重复拦截(R12)"""
    settings = engine.load_settings()
    rel = settings.get("最近文件")
    if not rel or engine.resolve_data_path(rel) is None:
        return redirect(url_for("index"))
    result = engine.run_audit(rel)
    added, total = engine.register_ledger(result, engine.load_reviews())
    return redirect(url_for("dashboard", ledger_added=added, ledger_total=total))


# ---------------------------------------------------------------------------
# 模块二:AI 填单 —— 扔凭证替代填表单
# ---------------------------------------------------------------------------

@app.route("/fill")
def fill():
    return render_template("fill.html", drafts=drafts.load_drafts(),
                           draft_path=str(drafts.DRAFT_FILE),
                           results=None, active="fill")


@app.route("/fill/upload", methods=["POST"])
def fill_upload():
    results = []
    for f in request.files.getlist("vouchers"):
        if not f or not f.filename:
            continue
        try:
            saved = drafts.save_voucher(f)
            row = drafts.add_from_file(saved)
            results.append({"文件": Path(f.filename).name, "行": row, "错误": None})
        except Exception as e:
            results.append({"文件": Path(f.filename).name, "行": None, "错误": str(e)})
    return render_template("fill.html", drafts=drafts.load_drafts(),
                           draft_path=str(drafts.DRAFT_FILE),
                           results=results, active="fill")


@app.route("/fill/delete", methods=["POST"])
def fill_delete():
    drafts.delete_draft(request.form.get("单号", ""))
    return redirect(url_for("fill"))


@app.route("/fill/clear", methods=["POST"])
def fill_clear():
    drafts.clear_drafts()
    return redirect(url_for("fill"))


@app.route("/fill/submit", methods=["POST"])
def fill_submit():
    """把草稿送去 AI 审批(草稿文件本身就是一份标准明细 Excel)"""
    if not drafts.DRAFT_FILE.exists():
        return redirect(url_for("fill"))
    engine.save_settings({"最近文件": drafts.draft_rel_path()})
    return redirect(url_for("dashboard"))


# ---------------------------------------------------------------------------
# 模块三:电子会计档案
# ---------------------------------------------------------------------------

@app.route("/archive")
def archive():
    idx = archive_lib.load_index()
    idx.sort(key=lambda e: e["档案号"], reverse=True)
    return render_template("archive.html", entries=idx, stats=archive_lib.stats(),
                           results=None, active="archive")


@app.route("/archive/upload", methods=["POST"])
def archive_upload():
    note = request.form.get("备注", "").strip()
    results = []
    for f in request.files.getlist("files"):
        if not f or not f.filename:
            continue
        with tempfile.NamedTemporaryFile(suffix=Path(f.filename).suffix, delete=False) as tmp:
            f.save(tmp.name)
            entry, err = archive_lib.archive_file(tmp.name, f.filename, note)
        results.append({"文件": Path(f.filename).name, "条目": entry, "错误": err})
    idx = archive_lib.load_index()
    idx.sort(key=lambda e: e["档案号"], reverse=True)
    return render_template("archive.html", entries=idx, stats=archive_lib.stats(),
                           results=results, active="archive")


@app.route("/archive/file/<no>")
def archive_download(no):
    e = archive_lib.find(no)
    if e is None:
        return "找不到档案", 404
    return send_file(archive_lib.file_path(e), as_attachment=True, download_name=e["原文件名"])


@app.route("/archive/verify/<no>")
def archive_verify(no):
    ok, msg = archive_lib.verify(no)
    return jsonify({"ok": ok, "msg": msg})


if __name__ == "__main__":
    engine.ensure_rules_workbook()
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:%d/" % PORT)).start()
    print("AI 财务审核智能体已启动: http://127.0.0.1:%d/  (按 Ctrl+C 退出)" % PORT)
    app.run(host="127.0.0.1", port=PORT, debug=False)
