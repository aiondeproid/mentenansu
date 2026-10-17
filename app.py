"""設備保全アプリ（プロトタイプ）

    python app.py            → http://<サーバー>:5000/
    MAINT_HOST / MAINT_PORT  で待ち受けを変更
"""
import csv
import datetime as dt
import hashlib
import io
import os

import re
from urllib.parse import quote

from flask import Flask, Response, abort, g, has_request_context, jsonify, request, send_from_directory

from db import CATEGORIES, DB_PATH, connect, init_db
from import_xls import norm
import equipment_xlsx
import pla
import pla_sheets
import parts_catalog

app = Flask(__name__, static_folder="static", static_url_path="")

RECORD_STATUSES = ["未着手", "処理中", "部品入荷待ち", "保留", "様子見", "完了"]
PART_STATUSES = ["使用中", "予備", "研磨中", "使用禁止", "廃棄"]
# 部品イベント → 変更後の状態
PART_EVENT_STATUS = {
    "装着": "使用中", "取外し": "予備", "研磨出し": "研磨中", "研磨戻り": "予備",
    "使用禁止": "使用禁止", "廃棄": "廃棄", "メモ": None,
}
FAILURE_KEYWORDS = ["ベアリング", "モーター", "オイルシール", "ベルト", "端子台|配線|漏電|ショート",
                    "リミット|センサー|スイッチ", "シャフト|軸", "チェーン|キー", "グリス|オイル漏"]
# 記入者の選択肢（この順で上に出す。過去の記録にある名前はこの後ろに並ぶ）
STAFF = [
    "矢代 顕久", "矢野 新悟", "塚本 雅行", "小森 吉貞", "髙橋 和也", "川東 忠司", "山本 浩善", "西村 憲弘",
    "筒井 祐司", "米林 威作人", "中村 徳秀", "大家 輝彦", "中野 翔吾", "浅野 冬弥", "杉本 陸", "冨永 悠士",
    "辻 隼正", "原田 翔太", "碇 大稀", "山嵜 悠", "北林 陸", "林 佑耶", "中野 倖多", "清原 憲",
    "藤本 昌樹", "大南 正人", "山本 剛",
    "秋葉 裕樹", "梶田 博幸", "山下 隆雄", "鈴川 剛史", "松本 憲二", "若井 春孝", "橋本 政幸", "中野 健志",
    "今宿 伸吾", "橋本 慎吾", "木村 真人", "木場田 智武", "西中 寿", "脇坂 翼", "大籠 恭二", "宮川 明",
    "長田 豊貴",
]
SOON_DAYS = 14
SOON_HOURS_RATIO = 0.2


def db():
    if "conn" not in g:
        g.conn = connect()
    return g.conn


@app.teardown_appcontext
def close_db(_exc):
    conn = g.pop("conn", None)
    if conn is not None:
        conn.close()


def rows(cur):
    return [dict(r) for r in cur]


def today():
    return dt.date.today()


def body():
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        abort(400, "JSON が必要です")
    return data


def blank_to_none(v):
    return None if v in ("", None) else v


# ---------------------------------------------------------------- アラート判定
def rule_state(rule, current_hours):
    """定期交換ルールの期限状況を返す。level: over / soon / ok / unknown"""
    out = {"days_left": None, "due_date": None, "hours_used": None, "hours_left": None}
    levels = []
    if rule["interval_days"] and rule["last_done_date"]:
        due = dt.date.fromisoformat(rule["last_done_date"]) + dt.timedelta(days=rule["interval_days"])
        left = (due - today()).days
        out.update(due_date=due.isoformat(), days_left=left)
        levels.append("over" if left < 0 else "soon" if left <= SOON_DAYS else "ok")
    if rule["interval_hours"] and current_hours is not None:
        used = current_hours - (rule["last_hours"] or 0)
        left = rule["interval_hours"] - used
        out.update(hours_used=round(used, 1), hours_left=round(left, 1))
        levels.append("over" if left < 0 else "soon" if left <= rule["interval_hours"] * SOON_HOURS_RATIO else "ok")
    for lv in ("over", "soon", "ok"):
        if lv in levels:
            out["level"] = lv
            break
    else:
        out["level"] = "unknown"
    return out


def part_state(part):
    if part["status"] in ("使用禁止", "廃棄") or part["grind_count"] is None:
        return "none"
    limit = part["grind_limit"] or 10
    if part["grind_count"] >= limit:
        return "over"
    if part["grind_count"] >= limit - 2:
        return "soon"
    return "ok"


def area_filter(alias="e"):
    """?area=前処理 のようなエリア指定を SQL 条件にする。"""
    area = request.args.get("area") if has_request_context() else None  # 通知バッチからも呼ばれる
    if area:
        return f" AND {alias}.area = ?", [area]
    return "", []


def load_rules(conn, equipment_id=None, include_inactive=False):
    cond, args = area_filter()
    sql = """SELECT r.*, e.name AS equipment_name, e.nickname AS equipment_nickname, e.code AS equipment_code,
                    e.area, e.current_hours,
                    e.hours_updated
             FROM rules r JOIN equipment e ON e.id = r.equipment_id
             WHERE e.status != '廃止'""" + cond
    if not include_inactive:
        sql += " AND r.active = 1"
    if equipment_id:
        sql += " AND r.equipment_id = ?"
        args.append(equipment_id)
    result = []
    for r in rows(conn.execute(sql, args)):
        r.update(rule_state(r, r["current_hours"]))
        result.append(r)
    order = {"over": 0, "soon": 1, "unknown": 2, "ok": 3}
    result.sort(key=lambda r: (order[r["level"]], r["days_left"] if r["days_left"] is not None else 99999))
    return result


# ---------------------------------------------------------------- 静的ファイル
@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.get("/api/meta")
def meta():
    conn = db()
    # 名簿の人を先に。過去の記録の名前は、名簿と同じ名前・名簿の人の姓だけのもの（例: 中村）を除いて後ろに
    staff_names = set(STAFF) | {s.split()[0] for s in STAFF}
    authors = STAFF + [r[0] for r in conn.execute(
        "SELECT author FROM records WHERE author IS NOT NULL GROUP BY author ORDER BY COUNT(*) DESC")
        if r[0] not in staff_names]
    return jsonify(
        record_statuses=RECORD_STATUSES,
        part_statuses=PART_STATUSES,
        categories=CATEGORIES,
        part_events=list(PART_EVENT_STATUS),
        authors=authors,
        areas=[r[0] for r in conn.execute(
            "SELECT area FROM equipment WHERE area IS NOT NULL GROUP BY area ORDER BY COUNT(*) DESC")],
        locations=[r[0] for r in conn.execute(
            """SELECT location FROM (SELECT location FROM parts UNION ALL SELECT location FROM part_catalog)
               WHERE location IS NOT NULL GROUP BY location ORDER BY COUNT(*) DESC, location""")],
        today=today().isoformat(),
    )


# ---------------------------------------------------------------- 設備
@app.get("/api/equipment")
def equipment_list():
    conn = db()
    eqs = rows(conn.execute(
        """SELECT e.*, v.name AS vendor_name,
                  (SELECT COUNT(*) FROM records r WHERE r.equipment_id = e.id) AS record_count,
                  (SELECT COUNT(*) FROM records r WHERE r.equipment_id = e.id AND r.kind = '不定期') AS failure_count,
                  (SELECT MAX(done_date) FROM records r WHERE r.equipment_id = e.id) AS last_date
           FROM equipment e LEFT JOIN vendors v ON v.id = e.vendor_id
           WHERE 1=1""" + area_filter()[0] + """
           ORDER BY e.status, e.area, COALESCE(e.code, 'ZZZ'), e.id""", area_filter()[1]))
    alerts = {}
    for r in load_rules(conn):
        a = alerts.setdefault(r["equipment_id"], {"over": 0, "soon": 0})
        if r["level"] in a:
            a[r["level"]] += 1
    for e in eqs:
        e["alerts"] = alerts.get(e["id"], {"over": 0, "soon": 0})
    return jsonify(eqs)


@app.get("/api/equipment/<int:eq_id>")
def equipment_detail(eq_id):
    conn = db()
    eq = conn.execute(
        """SELECT e.*, v.name AS vendor_name, v.tel AS vendor_tel, v.fax AS vendor_fax,
                  v.address AS vendor_address, v.contact AS vendor_contact, v.mobile AS vendor_mobile
           FROM equipment e LEFT JOIN vendors v ON v.id = e.vendor_id WHERE e.id = ?""", (eq_id,)).fetchone()
    if not eq:
        abort(404)
    eq = dict(eq)
    eq["rules"] = load_rules(conn, eq_id, include_inactive=True)
    eq["parts"] = rows(conn.execute("SELECT * FROM parts WHERE equipment_id = ? ORDER BY serial", (eq_id,)))
    for p in eq["parts"]:
        p["level"] = part_state(p)
    eq["catalog"] = rows(conn.execute(
        "SELECT * FROM part_catalog WHERE equipment_id = ? ORDER BY sort_order, id", (eq_id,)))
    return jsonify(eq)


EQ_FIELDS = ["code", "name", "nickname", "category", "model", "maker", "vendor_id", "supplier", "asset_no", "serial_no",
             "manufactured", "delivered", "area", "status", "note"]


@app.post("/api/equipment")
def equipment_create():
    d = body()
    if not d.get("name"):
        abort(400, "設備名は必須です")
    conn = db()
    cur = conn.execute(
        f"INSERT INTO equipment({','.join(EQ_FIELDS)}) VALUES ({','.join('?' * len(EQ_FIELDS))})",
        [blank_to_none(d.get(f)) for f in EQ_FIELDS])
    conn.commit()
    return jsonify(id=cur.lastrowid), 201


@app.put("/api/equipment/<int:eq_id>")
def equipment_update(eq_id):
    d = body()
    fields = [f for f in EQ_FIELDS if f in d]
    if not fields:
        abort(400)
    conn = db()
    conn.execute(f"UPDATE equipment SET {', '.join(f + '=?' for f in fields)} WHERE id=?",
                 [blank_to_none(d[f]) for f in fields] + [eq_id])
    conn.commit()
    return jsonify(ok=True)


@app.get("/api/equipment/export.xlsx")
def equipment_export():
    name = f"設備台帳_{today().strftime('%Y%m%d')}.xlsx"
    return Response(equipment_xlsx.export_workbook(db()),
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f"attachment; filename=equipment.xlsx; filename*=UTF-8''{quote(name)}"})


@app.post("/api/equipment/import")
def equipment_import():
    """?apply=1 のときだけ DB を変える。それ以外は変更内容の確認だけ。"""
    f = request.files.get("file")
    if not f:
        abort(400, "ファイルを選んでください")
    conn = db()
    try:
        plan = equipment_xlsx.plan_import(conn, f.read())
        if request.args.get("apply") == "1":
            return jsonify(applied=equipment_xlsx.apply_import(conn, plan))
    except ValueError as e:
        abort(400, str(e))
    return jsonify(plan)


@app.put("/api/equipment/<int:eq_id>/hours")
def equipment_hours(eq_id):
    d = body()
    try:
        hours = float(d["current_hours"])
    except (KeyError, TypeError, ValueError):
        abort(400, "アワーメーター値が不正です")
    conn = db()
    conn.execute("UPDATE equipment SET current_hours=?, hours_updated=? WHERE id=?",
                 (hours, today().isoformat(), eq_id))
    conn.commit()
    return jsonify(ok=True)


# ---------------------------------------------------------------- 作業履歴
def query_records(conn, args):
    sql = """SELECT r.*, e.name AS equipment_name, e.code AS equipment_code, ru.title AS rule_title,
                    (SELECT COUNT(*) FROM record_files f WHERE f.record_id = r.id) AS file_count,
                    (SELECT GROUP_CONCAT(f.name, ' ') FROM record_files f WHERE f.record_id = r.id) AS file_names
             FROM records r JOIN equipment e ON e.id = r.equipment_id
             LEFT JOIN rules ru ON ru.id = r.rule_id WHERE 1=1"""
    params = []
    if args.get("area"):
        sql += " AND e.area = ?"
        params.append(args["area"])
    if args.get("equipment_id"):
        sql += " AND r.equipment_id = ?"
        params.append(args["equipment_id"])
    if args.get("kind"):
        sql += " AND r.kind = ?"
        params.append(args["kind"])
    if args.get("status"):
        sts = args["status"].split(",")
        sql += f" AND r.status IN ({','.join('?' * len(sts))})"
        params += sts
    if args.get("uncertain") == "1":
        sql += " AND r.date_uncertain = 1"
    if args.get("from"):
        sql += " AND r.done_date >= ?"
        params.append(args["from"])
    if args.get("to"):
        sql += " AND r.done_date <= ?"
        params.append(args["to"])
    for word in norm(args.get("q", "")).split():
        sql += """ AND (r.content LIKE ? OR r.reason LIKE ? OR r.result LIKE ? OR r.countermeasure LIKE ?
                       OR r.author LIKE ? OR e.name LIKE ? OR e.code LIKE ? OR r.part_serial LIKE ?)"""
        params += [f"%{word}%"] * 8
    sort = args.get("sort", "done_date")
    if sort not in ("done_date", "equipment_code", "status", "kind", "downtime_min", "author"):
        sort = "done_date"
    direction = "ASC" if args.get("dir") == "asc" else "DESC"
    sql += f" ORDER BY {sort} IS NULL, {sort} {direction}, r.id DESC"
    return rows(conn.execute(sql, params))


@app.get("/api/records")
def records_list():
    return jsonify(query_records(db(), request.args))


@app.get("/api/records.csv")
def records_csv():
    recs = query_records(db(), request.args)
    cols = [("equipment_code", "設備No"), ("equipment_name", "設備名"), ("kind", "定期/不定期"),
            ("status", "作業状況"), ("done_date", "実施日"), ("date_raw", "実施日(元表記)"),
            ("downtime_min", "停止・修理時間(分)"), ("hour_meter", "稼働時間(h)"), ("content", "実施内容"),
            ("reason", "実施の理由"), ("result", "実施の結果"), ("countermeasure", "備考(今後の対策)"),
            ("author", "記入者"), ("cost", "費用(円)"), ("part_serial", "部品シリアル")]
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow([c[1] for c in cols])
    for r in recs:
        w.writerow([r.get(c[0]) if r.get(c[0]) is not None else "" for c in cols])
    return Response("﻿" + buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": "attachment; filename=maintenance_records.csv"})


@app.get("/api/records/<int:rec_id>")
def record_get(rec_id):
    conn = db()
    r = conn.execute("SELECT * FROM records WHERE id=?", (rec_id,)).fetchone()
    if not r:
        abort(404)
    files = rows(conn.execute("SELECT * FROM record_files WHERE record_id=? ORDER BY id", (rec_id,)))
    return jsonify(dict(r) | {"files": files})


REC_FIELDS = ["equipment_id", "kind", "status", "done_date", "downtime_min", "hour_meter", "content",
              "reason", "result", "countermeasure", "author", "cost", "part_serial", "rule_id"]


def clean_record(d):
    v = {f: blank_to_none(d.get(f)) for f in REC_FIELDS if f in d}
    for f in ("downtime_min", "cost", "rule_id", "equipment_id"):
        if v.get(f) is not None:
            v[f] = int(v[f])
    if v.get("hour_meter") is not None:
        v["hour_meter"] = float(v["hour_meter"])
    if v.get("done_date"):
        dt.date.fromisoformat(v["done_date"])  # 形式チェック
        v["date_uncertain"] = 0
    if v.get("part_serial"):
        v["part_serial"] = norm(v["part_serial"]).upper()
    return v


def after_record_saved(conn, rec_id, v):
    """完了した作業をルールの前回実施・部品の履歴に反映する。"""
    if v.get("status") != "完了" or not v.get("done_date"):
        return
    if v.get("rule_id"):
        eq = conn.execute("SELECT current_hours FROM equipment WHERE id=?", (v["equipment_id"],)).fetchone()
        conn.execute(
            """UPDATE rules SET last_done_date = MAX(COALESCE(last_done_date, ''), ?),
                   last_hours = CASE WHEN interval_hours IS NOT NULL THEN ? ELSE last_hours END
               WHERE id = ?""",
            (v["done_date"], eq["current_hours"] if eq and eq["current_hours"] is not None else 0, v["rule_id"]))
    if v.get("part_serial"):
        for serial in [s.strip() for s in v["part_serial"].split(",") if s.strip()]:
            p = conn.execute("SELECT id FROM parts WHERE serial=?", (serial,)).fetchone()
            if not p:
                continue
            exists = conn.execute("SELECT 1 FROM part_events WHERE part_id=? AND record_id=?", (p["id"], rec_id)).fetchone()
            if not exists:
                conn.execute(
                    "INSERT INTO part_events(part_id,event_date,event,hour_meter,record_id,note) VALUES (?,?,?,?,?,?)",
                    (p["id"], v["done_date"], "交換記録", v.get("hour_meter"), rec_id, "作業履歴より"))


@app.post("/api/records")
def record_create():
    try:
        v = clean_record(body())
    except ValueError as e:
        abort(400, f"入力値が不正です: {e}")
    if not v.get("equipment_id") or not v.get("content"):
        abort(400, "設備と実施内容は必須です")
    v.setdefault("status", "完了")
    v["source"] = "app"
    conn = db()
    cur = conn.execute(f"INSERT INTO records({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
    after_record_saved(conn, cur.lastrowid, v)
    conn.commit()
    return jsonify(id=cur.lastrowid), 201


@app.put("/api/records/<int:rec_id>")
def record_update(rec_id):
    try:
        v = clean_record(body())
    except ValueError as e:
        abort(400, f"入力値が不正です: {e}")
    conn = db()
    old = conn.execute("SELECT * FROM records WHERE id=?", (rec_id,)).fetchone()
    if not old:
        abort(404)
    v["updated_at"] = dt.datetime.now().isoformat(timespec="seconds")
    conn.execute(f"UPDATE records SET {', '.join(k + '=?' for k in v)} WHERE id=?", list(v.values()) + [rec_id])
    merged = dict(old) | v
    after_record_saved(conn, rec_id, merged)
    conn.commit()
    return jsonify(ok=True)


@app.delete("/api/records/<int:rec_id>")
def record_delete(rec_id):
    conn = db()
    conn.execute("DELETE FROM part_events WHERE record_id=?", (rec_id,))
    conn.execute("DELETE FROM record_files WHERE record_id=?", (rec_id,))  # ファイル本体はバックアップから戻せるよう残す
    conn.execute("DELETE FROM records WHERE id=?", (rec_id,))
    conn.commit()
    return jsonify(ok=True)


# ---- 添付ファイル（PDF・写真）
RECORD_FILE_DIR = os.path.join(os.path.dirname(DB_PATH), "record_files")
RECORD_FILE_TYPES = {"application/pdf": "pdf", "image/jpeg": "jpg", "image/png": "png"}
RECORD_FILE_MAX = 30 * 1024 * 1024


@app.get("/record-files/<path:name>")
def record_file(name):
    """inline で返す（PDF・写真はブラウザでそのまま開く）。?dl=1 なら元のファイル名でダウンロード。"""
    if request.args.get("dl"):
        r = db().execute("SELECT name FROM record_files WHERE file=? LIMIT 1", (name,)).fetchone()
        return send_from_directory(RECORD_FILE_DIR, name, as_attachment=True,
                                   download_name=(r and r["name"]) or name)
    return send_from_directory(RECORD_FILE_DIR, name, max_age=86400 * 30)


@app.post("/api/records/<int:rec_id>/files")
def record_file_upload(rec_id):
    conn = db()
    if not conn.execute("SELECT 1 FROM records WHERE id=?", (rec_id,)).fetchone():
        abort(404)
    files = request.files.getlist("file")
    if not files:
        abort(400, "ファイルを選んでください")
    saved = []
    for f in files:
        ext = RECORD_FILE_TYPES.get(f.mimetype) or {"pdf": "pdf", "jpg": "jpg", "jpeg": "jpg", "png": "png"}.get(
            os.path.splitext(f.filename or "")[1].lower().lstrip("."))
        if not ext:
            abort(400, f"PDF か JPEG・PNG のファイルを選んでください（{f.filename}）")
        data = f.read()
        if len(data) > RECORD_FILE_MAX:
            abort(400, f"ファイルが大きすぎます（30MB まで）: {f.filename}")
        os.makedirs(RECORD_FILE_DIR, exist_ok=True)
        name = hashlib.sha1(data).hexdigest()[:20] + "." + ext  # 同じ中身は 1 つのファイルを使い回す
        path = os.path.join(RECORD_FILE_DIR, name)
        if not os.path.exists(path):
            with open(path, "wb") as out:
                out.write(data)
        saved.append((rec_id, name, os.path.basename(f.filename or "") or name, len(data),
                      dt.datetime.now().isoformat(timespec="seconds")))
    conn.executemany("INSERT INTO record_files(record_id, file, name, size, created_at) VALUES (?,?,?,?,?)", saved)
    conn.commit()
    return jsonify(ok=True, count=len(saved)), 201


@app.delete("/api/record-files/<int:file_id>")
def record_file_delete(file_id):
    conn = db()
    conn.execute("DELETE FROM record_files WHERE id=?", (file_id,))  # 同じ中身をほかの記録でも使っていることがあるので本体は残す
    conn.commit()
    return jsonify(ok=True)


# ---------------------------------------------------------------- 定期交換ルール
RULE_FIELDS = ["equipment_id", "title", "interval_days", "interval_hours", "last_done_date", "last_hours",
               "note", "active"]


def clean_rule(d):
    v = {f: blank_to_none(d.get(f)) for f in RULE_FIELDS if f in d}
    for f in ("interval_days", "equipment_id", "active"):
        if v.get(f) is not None:
            v[f] = int(v[f])
    for f in ("interval_hours", "last_hours"):
        if v.get(f) is not None:
            v[f] = float(v[f])
    return v


@app.get("/api/rules")
def rules_list():
    return jsonify(load_rules(db(), request.args.get("equipment_id"), include_inactive=True))


@app.post("/api/rules")
def rule_create():
    v = clean_rule(body())
    if not v.get("equipment_id") or not v.get("title"):
        abort(400, "設備と項目名は必須です")
    if not v.get("interval_days") and not v.get("interval_hours"):
        abort(400, "間隔（日数か稼働時間）を入力してください")
    conn = db()
    cur = conn.execute(f"INSERT INTO rules({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
    conn.commit()
    return jsonify(id=cur.lastrowid), 201


@app.put("/api/rules/<int:rule_id>")
def rule_update(rule_id):
    v = clean_rule(body())
    conn = db()
    conn.execute(f"UPDATE rules SET {', '.join(k + '=?' for k in v)} WHERE id=?", list(v.values()) + [rule_id])
    conn.commit()
    return jsonify(ok=True)


@app.delete("/api/rules/<int:rule_id>")
def rule_delete(rule_id):
    conn = db()
    conn.execute("UPDATE records SET rule_id=NULL WHERE rule_id=?", (rule_id,))
    conn.execute("DELETE FROM rules WHERE id=?", (rule_id,))
    conn.commit()
    return jsonify(ok=True)


# ---------------------------------------------------------------- 点検計画（カレンダー）
SCHEDULE_MAX_DAYS = 400


@app.get("/api/schedule")
def schedule():
    """期間内の予定・実績をカレンダー用の一覧にして返す。

    type: planned（未完了の作業 = 予定日が決まっているもの） / due（定期項目の次回期限）
          projected（期限どおりに実施した場合の、その次以降の見込み） / done（完了した作業）
    期限を過ぎた定期項目は日付に置かず overdue に、稼働時間だけで管理する項目は hours_only に入れる。
    """
    try:
        start = dt.date.fromisoformat(request.args["from"])
        end = dt.date.fromisoformat(request.args["to"])
    except (KeyError, ValueError):
        abort(400, "from と to を YYYY-MM-DD で指定してください")
    if not start <= end or (end - start).days > SCHEDULE_MAX_DAYS:
        abort(400, f"期間は {SCHEDULE_MAX_DAYS} 日以内で指定してください")
    conn = db()
    cond, args = area_filter()
    rules = load_rules(conn)
    s, e = start.isoformat(), end.isoformat()

    eq_cols = "e.name AS equipment_name, e.nickname AS equipment_nickname, e.code AS equipment_code, e.area"
    open_recs = rows(conn.execute(
        f"""SELECT r.id, r.equipment_id, r.rule_id, r.kind, r.status, r.done_date, r.content, r.author, {eq_cols}
            FROM records r JOIN equipment e ON e.id = r.equipment_id
            WHERE r.status != '完了' AND r.done_date IS NOT NULL""" + cond + " ORDER BY r.done_date", args))
    done_recs = rows(conn.execute(
        f"""SELECT r.id, r.equipment_id, r.rule_id, r.kind, r.status, r.done_date, r.content, r.author, {eq_cols}
            FROM records r JOIN equipment e ON e.id = r.equipment_id
            WHERE r.status = '完了' AND r.done_date BETWEEN ? AND ?""" + cond + " ORDER BY r.done_date",
        [s, e] + args))

    events = [dict(r, date=r["done_date"], type="planned", title=r["content"]) for r in open_recs if s <= r["done_date"] <= e]
    events += [dict(r, date=r["done_date"], type="done", title=r["content"]) for r in done_recs]

    # 予定日が入っている定期項目は、その予定日を起点に次回以降を見込む
    planned_by_rule = {}
    for r in open_recs:
        planned_by_rule.setdefault(r["rule_id"], r["done_date"])
    overdue, hours_only = [], []
    for r in rules:
        base = {k: r[k] for k in ("equipment_id", "equipment_name", "equipment_nickname", "equipment_code", "area",
                                  "title", "level", "interval_days", "interval_hours")} | {"rule_id": r["id"]}
        if not r["interval_days"]:
            hours_only.append(base | {"hours_left": r["hours_left"]})
            continue
        planned = planned_by_rule.get(r["id"])
        if planned:
            nxt = dt.date.fromisoformat(planned)
        elif r["due_date"] and r["days_left"] >= 0:
            nxt = dt.date.fromisoformat(r["due_date"])
            if nxt >= start:
                events.append(base | {"date": r["due_date"], "type": "due"})
        else:
            overdue.append(base | {"due_date": r["due_date"], "days_left": r["days_left"],
                                   "last_done_date": r["last_done_date"]})
            continue
        step = dt.timedelta(days=r["interval_days"])
        nxt += step
        while nxt <= end:
            if nxt >= start:
                events.append(base | {"date": nxt.isoformat(), "type": "projected", "level": "ok"})
            nxt += step

    order = {"planned": 0, "due": 1, "projected": 2, "done": 3}
    events.sort(key=lambda x: (x["date"], order[x["type"]], x.get("equipment_code") or "", x["title"] or ""))
    overdue.sort(key=lambda x: (x["due_date"] or "", x["equipment_code"] or ""))
    return jsonify(events=events, overdue=overdue, hours_only=hours_only, today=today().isoformat(),
                   rules=[{k: r[k] for k in ("id", "equipment_id", "equipment_name", "equipment_nickname", "equipment_code",
                                             "title", "level", "interval_days", "interval_hours", "last_done_date",
                                             "due_date")} for r in rules])


# ---------------------------------------------------------------- 部品・研磨
@app.get("/api/parts")
def parts_list():
    conn = db()
    parts = rows(conn.execute(
        """SELECT p.*, e.name AS equipment_name, e.code AS equipment_code,
                  (SELECT MAX(event_date) FROM part_events pe WHERE pe.part_id = p.id) AS last_event_date,
                  (SELECT COUNT(*) FROM part_events pe WHERE pe.part_id = p.id) AS event_count
           FROM parts p LEFT JOIN equipment e ON e.id = p.equipment_id
           ORDER BY CASE p.status WHEN '使用中' THEN 0 WHEN '予備' THEN 1 WHEN '研磨中' THEN 2
                                  WHEN '使用禁止' THEN 3 ELSE 4 END, p.serial"""))
    for p in parts:
        p["level"] = part_state(p)
    return jsonify(parts)


@app.get("/api/parts/<int:part_id>")
def part_detail(part_id):
    conn = db()
    p = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
    if not p:
        abort(404)
    p = dict(p)
    p["level"] = part_state(p)
    p["events"] = rows(conn.execute(
        """SELECT pe.*, e.name AS equipment_name FROM part_events pe
           LEFT JOIN records r ON r.id = pe.record_id LEFT JOIN equipment e ON e.id = r.equipment_id
           WHERE pe.part_id=? ORDER BY pe.event_date DESC, pe.id DESC""", (part_id,)))
    return jsonify(p)


PART_FIELDS = ["part_type", "serial", "equipment_id", "grind_count", "grind_limit", "status", "location", "note"]


@app.post("/api/parts")
def part_create():
    d = body()
    v = {f: blank_to_none(d.get(f)) for f in PART_FIELDS if f in d}
    if not v.get("serial") or not v.get("part_type"):
        abort(400, "部品種別とシリアルは必須です")
    v["serial"] = norm(v["serial"]).upper()
    conn = db()
    if conn.execute("SELECT 1 FROM parts WHERE serial=?", (v["serial"],)).fetchone():
        abort(409, "同じシリアルが既に登録されています")
    cur = conn.execute(f"INSERT INTO parts({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
    conn.commit()
    return jsonify(id=cur.lastrowid), 201


@app.put("/api/parts/<int:part_id>")
def part_update(part_id):
    d = body()
    v = {f: blank_to_none(d.get(f)) for f in PART_FIELDS if f in d}
    conn = db()
    conn.execute(f"UPDATE parts SET {', '.join(k + '=?' for k in v)} WHERE id=?", list(v.values()) + [part_id])
    conn.commit()
    return jsonify(ok=True)


@app.post("/api/parts/<int:part_id>/events")
def part_event(part_id):
    d = body()
    event = d.get("event")
    if event not in PART_EVENT_STATUS:
        abort(400, "イベント種別が不正です")
    conn = db()
    p = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
    if not p:
        abort(404)
    conn.execute(
        "INSERT INTO part_events(part_id,event_date,event,hour_meter,note) VALUES (?,?,?,?,?)",
        (part_id, d.get("event_date") or today().isoformat(), event,
         blank_to_none(d.get("hour_meter")), blank_to_none(d.get("note"))))
    status = PART_EVENT_STATUS[event]
    grind = p["grind_count"]
    message = None
    if event == "研磨出し":
        grind = (grind or 0) + 1
        if grind >= (p["grind_limit"] or 10):
            message = f"研磨回数が上限（{p['grind_limit'] or 10}回）に達しました。戻り後の使用可否を確認してください。"
    if event == "装着" and p["equipment_id"]:
        # 同じ設備で使用中だった部品は取外し扱いにする
        conn.execute("UPDATE parts SET status='予備' WHERE equipment_id=? AND status='使用中' AND id!=?",
                     (p["equipment_id"], part_id))
    conn.execute("UPDATE parts SET grind_count=?, status=COALESCE(?, status) WHERE id=?", (grind, status, part_id))
    conn.commit()
    return jsonify(ok=True, message=message)


# ---------------------------------------------------------------- 部品一覧（価格表）・Excel 入出力
CATALOG_FIELDS = ["group_name", "machine", "name", "unit_price", "qty", "note", "memo", "location", "equipment_id", "area"]


@app.get("/api/catalog")
def catalog_list():
    area = request.args.get("area")
    sql = """SELECT c.*, e.name AS equipment_name, e.nickname AS equipment_nickname, e.code AS equipment_code
             FROM part_catalog c LEFT JOIN equipment e ON e.id = c.equipment_id"""
    args = []
    if area:
        sql += " WHERE COALESCE(e.area, c.area) = ?"
        args.append(area)
    return jsonify(rows(db().execute(sql + " ORDER BY c.sort_order, c.id", args)))


def catalog_values(d):
    v = {f: blank_to_none(d.get(f)) for f in CATALOG_FIELDS if f in d}
    if "name" in v and not (v["name"] or "").strip():
        abort(400, "部品名は必須です")
    v["updated_at"] = dt.datetime.now().isoformat(timespec="seconds")
    return v


@app.post("/api/catalog")
def catalog_create():
    v = catalog_values(body())
    if not v.get("name"):
        abort(400, "部品名は必須です")
    conn = db()
    v["sort_order"] = conn.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM part_catalog").fetchone()[0]
    v["source"] = "app"
    cur = conn.execute(f"INSERT INTO part_catalog({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
    conn.commit()
    return jsonify(id=cur.lastrowid), 201


@app.put("/api/catalog/<int:item_id>")
def catalog_update(item_id):
    v = catalog_values(body())
    conn = db()
    conn.execute(f"UPDATE part_catalog SET {', '.join(k + '=?' for k in v)} WHERE id=?", list(v.values()) + [item_id])
    conn.commit()
    return jsonify(ok=True)


@app.delete("/api/catalog/<int:item_id>")
def catalog_delete(item_id):
    conn = db()
    conn.execute("DELETE FROM part_catalog WHERE id=?", (item_id,))
    conn.commit()
    return jsonify(ok=True)


@app.get("/api/parts/export.xlsx")
def parts_export():
    data = parts_catalog.export_workbook(db())
    name = f"部品一覧_{today().strftime('%Y%m%d')}.xlsx"
    return Response(data, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f"attachment; filename=parts.xlsx; filename*=UTF-8''{quote(name)}"})


@app.post("/api/parts/import")
def parts_import():
    """?apply=1 のときだけ DB を変える。それ以外は変更内容の確認だけ。"""
    f = request.files.get("file")
    if not f:
        abort(400, "ファイルを選んでください")
    conn = db()
    try:
        plan = parts_catalog.plan_import(conn, f.read())
        if request.args.get("apply") == "1":
            return jsonify(applied=parts_catalog.apply_import(conn, plan))
    except ValueError as e:
        abort(400, str(e))
    return jsonify(plan)


# ---------------------------------------------------------------- プラ点検
# プラリスト（pla_items）= 危害度分析の一覧。チェック表（pla_sheets / pla_points）= 現場で毎日つける紙の表。
# 点検はチェック表の 1 行ごとに記録し（pla_checks）、行はプラリストの行と紐づける（pla_point_items）。
PLA_FIELDS = ["process", "equipment", "item", "color", "material", "material_pct", "material_maker", "hazard",
              "methods", "is_new", "frequency", "severity", "possibility", "evaluation", "note", "active"]
PHOTO_TYPES = {"image/jpeg": "jpg", "image/png": "png", "image/webp": "webp", "image/gif": "gif"}


def pla_process():
    p = request.args.get("process") or "前工程"
    if p not in pla.PROCESS_AREA:
        abort(400, "工程が不正です")
    return p


@app.get("/api/pla/meta")
def pla_meta():
    return jsonify(processes=list(pla.PROCESS_AREA), process_area=pla.PROCESS_AREA, timings=pla_sheets.TIMINGS,
                   hazards=pla.HAZARDS, methods=[n for _, n in pla.METHOD_PATTERNS])


@app.get("/pla-photos/<path:name>")
def pla_photo(name):
    return send_from_directory(pla_sheets.PHOTO_DIR, name, max_age=86400 * 30)


# ---- プラリスト
@app.get("/api/pla/items")
def pla_items_list():
    conn = db()
    items = rows(conn.execute("SELECT * FROM pla_items WHERE process = ? ORDER BY sort_order, id", (pla_process(),)))
    links = {}
    for r in conn.execute(
            """SELECT x.item_id, s.id AS sheet_id, s.title, p.id AS point_id FROM pla_point_items x
               JOIN pla_points p ON p.id = x.point_id JOIN pla_sheets s ON s.id = p.sheet_id
               WHERE s.active = 1 AND p.active = 1 ORDER BY s.sort_order"""):
        links.setdefault(r["item_id"], []).append({"sheet_id": r["sheet_id"], "title": r["title"], "point_id": r["point_id"]})
    for i in items:
        i["daily"] = pla.is_daily(i)
        i["sheets"] = links.get(i["id"], [])
    return jsonify(items)


def pla_values(d):
    v = {f: blank_to_none(d.get(f)) for f in PLA_FIELDS if f in d}
    for f in ("is_new", "active"):
        if f in v:
            v[f] = 1 if v[f] in (1, "1", True, "on", "true") else 0
    if "item" in v and not (v["item"] or "").strip():
        abort(400, "点検箇所（機械・器具）は必須です")
    if "process" in v and v["process"] not in pla.PROCESS_AREA:
        abort(400, "工程が不正です")
    v["updated_at"] = dt.datetime.now().isoformat(timespec="seconds")
    return v


@app.post("/api/pla/items")
def pla_item_create():
    v = pla_values(body())
    if not v.get("item") or not v.get("equipment") or not v.get("process"):
        abort(400, "工程・設備名・点検箇所は必須です")
    v.setdefault("active", 1)
    conn = db()
    last = conn.execute("SELECT MAX(sort_order) FROM pla_items WHERE process=? AND equipment=?",
                        (v["process"], v["equipment"])).fetchone()[0]
    if last is None:
        last = conn.execute("SELECT COALESCE(MAX(sort_order), 0) FROM pla_items WHERE process=?", (v["process"],)).fetchone()[0]
    conn.execute("UPDATE pla_items SET sort_order = sort_order + 1 WHERE process=? AND sort_order > ?", (v["process"], last))
    v["sort_order"], v["source"] = last + 1, "app"
    cur = conn.execute(f"INSERT INTO pla_items({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
    conn.commit()
    return jsonify(id=cur.lastrowid), 201


@app.put("/api/pla/items/<int:item_id>")
def pla_item_update(item_id):
    v = pla_values(body())
    conn = db()
    conn.execute(f"UPDATE pla_items SET {', '.join(k + '=?' for k in v)} WHERE id=?", list(v.values()) + [item_id])
    conn.commit()
    return jsonify(ok=True)


@app.delete("/api/pla/items/<int:item_id>")
def pla_item_delete(item_id):
    conn = db()
    conn.execute("DELETE FROM pla_point_items WHERE item_id=?", (item_id,))
    conn.execute("DELETE FROM pla_items WHERE id=?", (item_id,))
    conn.commit()
    return jsonify(ok=True)


# ---- チェック表
def pla_sheet_rows(conn, process, active_only=True):
    """チェック表と行・写真・紐づくプラリストの行をまとめて返す。"""
    cond = " AND active = 1" if active_only else ""
    sheets = rows(conn.execute(f"SELECT * FROM pla_sheets WHERE process = ?{cond} ORDER BY sort_order, id", (process,)))
    if not sheets:
        return []
    ids = ",".join(str(s["id"]) for s in sheets)
    points = rows(conn.execute(f"SELECT * FROM pla_points WHERE sheet_id IN ({ids}){cond} ORDER BY sort_order, id"))
    photos = rows(conn.execute(f"SELECT * FROM pla_photos WHERE sheet_id IN ({ids}) ORDER BY sort_order, id"))
    links = rows(conn.execute(
        f"""SELECT x.point_id, i.id, i.equipment, i.item, i.hazard, i.material, i.material_maker, i.color, i.active
            FROM pla_point_items x JOIN pla_items i ON i.id = x.item_id
            WHERE x.point_id IN (SELECT id FROM pla_points WHERE sheet_id IN ({ids})) ORDER BY i.sort_order"""))
    by_sheet = {s["id"]: s for s in sheets}
    for s in sheets:
        s["points"], s["photos"] = [], []
    by_point = {}
    for p in points:
        p["photos"], p["items"] = [], []
        by_point[p["id"]] = p
        by_sheet[p["sheet_id"]]["points"].append(p)
    for ph in photos:
        (by_point[ph["point_id"]]["photos"] if ph["point_id"] in by_point else by_sheet[ph["sheet_id"]]["photos"]).append(ph)
    for l in links:
        if l["point_id"] in by_point:
            by_point[l["point_id"]]["items"].append(l)
    return [s for s in sheets if s["points"] or not active_only]


@app.get("/api/pla/sheets")
def pla_sheets_list():
    return jsonify(pla_sheet_rows(db(), pla_process(), active_only=request.args.get("all") != "1"))


@app.put("/api/pla/sheets/<int:sheet_id>")
def pla_sheet_update(sheet_id):
    d = body()
    v = {k: blank_to_none(d.get(k)) for k in ("title", "note") if k in d}
    if "active" in d:
        v["active"] = 1 if d["active"] in (1, "1", True) else 0
    if "title" in v and not v["title"]:
        abort(400, "チェック表の名前は必須です")
    v["updated_at"] = dt.datetime.now().isoformat(timespec="seconds")
    conn = db()
    conn.execute(f"UPDATE pla_sheets SET {', '.join(k + '=?' for k in v)} WHERE id=?", list(v.values()) + [sheet_id])
    conn.commit()
    return jsonify(ok=True)


@app.post("/api/pla/sheets")
def pla_sheet_create():
    d = body()
    if not (d.get("title") or "").strip():
        abort(400, "チェック表の名前は必須です")
    process = d.get("process")
    if process not in pla.PROCESS_AREA:
        abort(400, "工程が不正です")
    conn = db()
    order = conn.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM pla_sheets").fetchone()[0]
    cur = conn.execute("INSERT INTO pla_sheets(process, title, active, sort_order, source, updated_at) VALUES (?,?,1,?,?,?)",
                       (process, d["title"].strip(), order, "app", dt.datetime.now().isoformat(timespec="seconds")))
    conn.commit()
    return jsonify(id=cur.lastrowid), 201


@app.post("/api/pla/points")
def pla_point_create():
    d = body()
    if not (d.get("item_text") or "").strip() or not d.get("sheet_id"):
        abort(400, "項目は必須です")
    conn = db()
    order = conn.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM pla_points WHERE sheet_id=?", (d["sheet_id"],)).fetchone()[0]
    cur = conn.execute("INSERT INTO pla_points(sheet_id, location, item_text, sort_order) VALUES (?,?,?,?)",
                       (d["sheet_id"], blank_to_none(d.get("location")), d["item_text"].strip(), order))
    conn.commit()
    return jsonify(id=cur.lastrowid), 201


@app.put("/api/pla/points/<int:point_id>")
def pla_point_update(point_id):
    d = body()
    conn = db()
    v = {k: blank_to_none(d.get(k)) for k in ("location", "item_text") if k in d}
    if "active" in d:
        v["active"] = 1 if d["active"] in (1, "1", True) else 0
    if "item_text" in v and not v["item_text"]:
        abort(400, "項目は必須です")
    if v:
        conn.execute(f"UPDATE pla_points SET {', '.join(k + '=?' for k in v)} WHERE id=?", list(v.values()) + [point_id])
    if "item_ids" in d:  # プラリストとの紐づけをこの内容に置き換える
        conn.execute("DELETE FROM pla_point_items WHERE point_id=?", (point_id,))
        for iid in d["item_ids"] or []:
            conn.execute("INSERT OR IGNORE INTO pla_point_items(point_id, item_id) VALUES (?,?)", (point_id, int(iid)))
    if "move" in d:  # 1 つ上 / 下へ
        p = conn.execute("SELECT * FROM pla_points WHERE id=?", (point_id,)).fetchone()
        op, order = ("<", "DESC") if d["move"] == "up" else (">", "ASC")
        other = conn.execute(f"SELECT * FROM pla_points WHERE sheet_id=? AND sort_order {op} ? ORDER BY sort_order {order} LIMIT 1",
                             (p["sheet_id"], p["sort_order"])).fetchone()
        if other:
            conn.execute("UPDATE pla_points SET sort_order=? WHERE id=?", (other["sort_order"], p["id"]))
            conn.execute("UPDATE pla_points SET sort_order=? WHERE id=?", (p["sort_order"], other["id"]))
    conn.commit()
    return jsonify(ok=True)


@app.delete("/api/pla/points/<int:point_id>")
def pla_point_delete(point_id):
    conn = db()
    if conn.execute("SELECT 1 FROM pla_checks WHERE point_id=?", (point_id,)).fetchone():
        abort(409, "点検記録がある行は削除できません。「使わない」にしてください")
    conn.execute("DELETE FROM pla_point_items WHERE point_id=?", (point_id,))
    conn.execute("UPDATE pla_photos SET point_id=NULL WHERE point_id=?", (point_id,))
    conn.execute("DELETE FROM pla_points WHERE id=?", (point_id,))
    conn.commit()
    return jsonify(ok=True)


@app.post("/api/pla/photos")
def pla_photo_upload():
    """写真を足す（point_id があればその行、なければチェック表の「場所の写真」）。"""
    f = request.files.get("file")
    sheet_id = request.form.get("sheet_id", type=int)
    point_id = request.form.get("point_id", type=int) or None
    if not f or not sheet_id:
        abort(400, "写真を選んでください")
    ext = PHOTO_TYPES.get(f.mimetype)
    if not ext:
        abort(400, "JPEG か PNG の写真を選んでください")
    data = f.read()
    if len(data) > 15 * 1024 * 1024:
        abort(400, "写真が大きすぎます（15MB まで）")
    name = pla_sheets._save_photo(data, ext)
    conn = db()
    order = conn.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM pla_photos WHERE sheet_id=?", (sheet_id,)).fetchone()[0]
    conn.execute("INSERT INTO pla_photos(sheet_id, point_id, file, sort_order, caption, source, created_at) VALUES (?,?,?,?,?,?,?)",
                 (sheet_id, point_id, name, order, blank_to_none(request.form.get("caption")), "app",
                  dt.datetime.now().isoformat(timespec="seconds")))
    conn.commit()
    return jsonify(ok=True, file=name), 201


@app.delete("/api/pla/photos/<int:photo_id>")
def pla_photo_delete(photo_id):
    conn = db()
    conn.execute("DELETE FROM pla_photos WHERE id=?", (photo_id,))  # 写真ファイルはほかの行でも使っていることがあるので残す
    conn.commit()
    return jsonify(ok=True)


# ---- 点検の記録と確認
def pla_date_timing(d):
    date, timing = d.get("date"), d.get("timing")
    try:
        dt.date.fromisoformat(date or "")
    except ValueError:
        abort(400, "日付が不正です")
    if timing not in pla_sheets.TIMINGS:
        abort(400, "始業時・終業時のどちらかを選んでください")
    return date, timing


@app.get("/api/pla/check")
def pla_check_sheet():
    """点検する画面のデータ。使っているチェック表と、その日・その時間の結果。"""
    process = pla_process()
    date, timing = pla_date_timing(request.args)
    conn = db()
    results = {r["point_id"]: r for r in rows(conn.execute(
        """SELECT c.* FROM pla_checks c JOIN pla_points p ON p.id = c.point_id JOIN pla_sheets s ON s.id = p.sheet_id
           WHERE s.process = ? AND c.check_date = ? AND c.timing = ?""", (process, date, timing)))}
    review = conn.execute("SELECT * FROM pla_reviews WHERE process=? AND check_date=? AND timing=?",
                          (process, date, timing)).fetchone()
    return jsonify(sheets=pla_sheet_rows(conn, process), results=results, review=dict(review) if review else None)


@app.put("/api/pla/results")
def pla_results_save():
    """結果をまとめて保存する。result が空なら取り消し。"""
    d = body()
    date, timing = pla_date_timing(d)
    who = (d.get("checked_by") or "").strip()
    if not who:
        abort(400, "右上の「記入者」に名前を入れてください")
    now = dt.datetime.now().isoformat(timespec="seconds")
    conn = db()
    for r in d.get("results") or []:
        result = r.get("result")
        if result not in ("ok", "ng", "na", None, ""):
            abort(400, "結果が不正です")
        if result == "ng" and not (r.get("note") or "").strip():
            abort(400, "異常のときは内容を書いてください")
        if not result:
            conn.execute("DELETE FROM pla_checks WHERE point_id=? AND check_date=? AND timing=?", (r["point_id"], date, timing))
            continue
        conn.execute(
            """INSERT INTO pla_checks(point_id, check_date, timing, result, note, checked_by, checked_at)
               VALUES (?,?,?,?,?,?,?)
               ON CONFLICT(point_id, check_date, timing) DO UPDATE SET
                 result=excluded.result, note=excluded.note, checked_by=excluded.checked_by, checked_at=excluded.checked_at""",
            (r["point_id"], date, timing, result, blank_to_none(r.get("note")) if result == "ng" else None, who, now))
    conn.commit()
    return jsonify(ok=True)


@app.get("/api/pla/summary")
def pla_summary():
    """確認画面のデータ。月の日ごと・時間ごとの点検数、異常、確認者。"""
    process = pla_process()
    ym = request.args.get("month") or today().strftime("%Y-%m")
    try:
        first = dt.date.fromisoformat(ym + "-01")
    except ValueError:
        abort(400, "月が不正です")
    last = (first.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
    conn = db()
    total = conn.execute("""SELECT COUNT(*) FROM pla_points p JOIN pla_sheets s ON s.id = p.sheet_id
                            WHERE s.process = ? AND s.active = 1 AND p.active = 1""", (process,)).fetchone()[0]
    days = {}
    for r in conn.execute(
            """SELECT c.check_date, c.timing, COUNT(*) AS done, SUM(c.result = 'ng') AS ng,
                      GROUP_CONCAT(DISTINCT c.checked_by) AS checkers
               FROM pla_checks c JOIN pla_points p ON p.id = c.point_id JOIN pla_sheets s ON s.id = p.sheet_id
               WHERE s.process = ? AND c.check_date BETWEEN ? AND ? GROUP BY c.check_date, c.timing""",
            (process, first.isoformat(), last.isoformat())):
        days.setdefault(r["check_date"], {})[r["timing"]] = {"done": r["done"], "ng": r["ng"], "checkers": r["checkers"]}
    for r in conn.execute("SELECT * FROM pla_reviews WHERE process=? AND check_date BETWEEN ? AND ?",
                          (process, first.isoformat(), last.isoformat())):
        days.setdefault(r["check_date"], {}).setdefault(r["timing"], {"done": 0, "ng": 0, "checkers": None})["review"] = dict(r)
    ng = rows(conn.execute(
        """SELECT c.*, s.title AS sheet_title, s.id AS sheet_id, p.location, p.item_text FROM pla_checks c
           JOIN pla_points p ON p.id = c.point_id JOIN pla_sheets s ON s.id = p.sheet_id
           WHERE s.process = ? AND c.result = 'ng' AND c.check_date BETWEEN ? AND ?
           ORDER BY c.check_date DESC, c.timing DESC, s.sort_order, p.sort_order""",
        (process, first.isoformat(), last.isoformat())))
    return jsonify(total=total, days=days, ng=ng, month=ym)


@app.put("/api/pla/reviews")
def pla_review_save():
    d = body()
    date, timing = pla_date_timing(d)
    process = d.get("process")
    if process not in pla.PROCESS_AREA:
        abort(400, "工程が不正です")
    reviewer = (d.get("reviewer") or "").strip()
    conn = db()
    if d.get("cancel"):
        conn.execute("DELETE FROM pla_reviews WHERE process=? AND check_date=? AND timing=?", (process, date, timing))
    else:
        if not reviewer:
            abort(400, "右上の「記入者」に確認者の名前を入れてください")
        conn.execute(
            """INSERT INTO pla_reviews(process, check_date, timing, reviewer, reviewed_at, comment) VALUES (?,?,?,?,?,?)
               ON CONFLICT(process, check_date, timing) DO UPDATE SET
                 reviewer=excluded.reviewer, reviewed_at=excluded.reviewed_at, comment=excluded.comment""",
            (process, date, timing, reviewer, dt.datetime.now().isoformat(timespec="seconds"), blank_to_none(d.get("comment"))))
    conn.commit()
    return jsonify(ok=True)


@app.get("/api/pla/export.xlsx")
def pla_export():
    process = pla_process()
    ym = request.args.get("month") or today().strftime("%Y-%m")
    name = f"プラ点検記録_{process}_{ym}.xlsx"
    return Response(pla_sheets.export_month(db(), process, ym),
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f"attachment; filename=pla.xlsx; filename*=UTF-8''{quote(name)}"})


# ---------------------------------------------------------------- 取引業者
@app.get("/api/vendors")
def vendors_list():
    return jsonify(rows(db().execute(
        """SELECT v.*, GROUP_CONCAT(e.name, '、') AS equipment_names
           FROM vendors v LEFT JOIN equipment e ON e.vendor_id = v.id GROUP BY v.id ORDER BY v.id""")))


@app.put("/api/vendors/<int:vendor_id>")
def vendor_update(vendor_id):
    d = body()
    fields = [f for f in ("name", "tel", "fax", "address", "contact", "mobile") if f in d]
    conn = db()
    conn.execute(f"UPDATE vendors SET {', '.join(f + '=?' for f in fields)} WHERE id=?",
                 [d[f] for f in fields] + [vendor_id])
    conn.commit()
    return jsonify(ok=True)


# ---------------------------------------------------------------- ダッシュボード・分析
@app.get("/api/dashboard")
def dashboard():
    conn = db()
    rules = load_rules(conn)
    cond, args = area_filter()
    parts = [dict(p) | {"level": part_state(p)} for p in conn.execute(
        "SELECT p.*, e.name AS equipment_name FROM parts p LEFT JOIN equipment e ON e.id = p.equipment_id"
        + (" WHERE e.area = ?" if args else ""), args)]
    open_records = rows(conn.execute(
        """SELECT r.*, e.name AS equipment_name, e.code AS equipment_code FROM records r
           JOIN equipment e ON e.id = r.equipment_id WHERE r.status != '完了'""" + cond + """
           ORDER BY r.done_date DESC""", args))
    recent = rows(conn.execute(
        """SELECT r.*, e.name AS equipment_name, e.code AS equipment_code FROM records r
           JOIN equipment e ON e.id = r.equipment_id WHERE r.done_date IS NOT NULL""" + cond + """
           ORDER BY r.done_date DESC, r.id DESC LIMIT 8""", args))
    uncertain = conn.execute(
        "SELECT COUNT(*) FROM records r JOIN equipment e ON e.id = r.equipment_id WHERE r.date_uncertain=1" + cond,
        args).fetchone()[0]
    return jsonify(
        rules_over=[r for r in rules if r["level"] == "over"],
        rules_soon=[r for r in rules if r["level"] == "soon"],
        rules_unknown=[r for r in rules if r["level"] == "unknown"],
        parts_alert=[p for p in parts if p["level"] in ("over", "soon")],
        open_records=open_records,
        recent=recent,
        uncertain_count=uncertain,
        equipment_count=conn.execute("SELECT COUNT(*) FROM equipment e WHERE status!='廃止'" + cond, args).fetchone()[0],
    )


@app.get("/api/analytics")
def analytics():
    conn = db()
    years = request.args.get("years", type=int) or 5
    since = f"{today().year - years + 1}-01-01"
    cond, args = area_filter()
    by_year = rows(conn.execute(
        """SELECT substr(done_date,1,4) AS year,
                  SUM(CASE WHEN kind='定期' THEN 1 ELSE 0 END) AS teiki,
                  SUM(CASE WHEN kind='不定期' THEN 1 ELSE 0 END) AS futeiki,
                  SUM(CASE WHEN kind IS NULL THEN 1 ELSE 0 END) AS other,
                  COALESCE(SUM(downtime_min),0) AS downtime, COALESCE(SUM(cost),0) AS cost
           FROM records r JOIN equipment e ON e.id = r.equipment_id
           WHERE done_date IS NOT NULL""" + cond + " GROUP BY year ORDER BY year", args))
    by_equipment = rows(conn.execute(
        """SELECT e.id, e.code, e.name, COUNT(r.id) AS count,
                  SUM(CASE WHEN r.kind='不定期' THEN 1 ELSE 0 END) AS futeiki,
                  COALESCE(SUM(r.downtime_min),0) AS downtime,
                  COALESCE(SUM(r.cost),0) AS cost
           FROM equipment e JOIN records r ON r.equipment_id = e.id
           WHERE r.done_date >= ?""" + cond + " GROUP BY e.id ORDER BY downtime DESC, futeiki DESC",
        [since] + args))
    texts =[f"{r['content'] or ''} {r['reason'] or ''}" for r in conn.execute(
        """SELECT content, reason FROM records r JOIN equipment e ON e.id = r.equipment_id
           WHERE done_date >= ? AND (kind='不定期' OR kind IS NULL)""" + cond, [since] + args)]
    keywords = []
    for kw in FAILURE_KEYWORDS:
        pat = re.compile(kw)
        keywords.append({"keyword": kw.replace("|", "・"), "count": sum(1 for t in texts if pat.search(t))})
    keywords.sort(key=lambda k: -k["count"])
    return jsonify(since=since, by_year=by_year, by_equipment=by_equipment, keywords=keywords)


# ---------------------------------------------------------------- 通知設定
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+$")


@app.get("/api/notify")
def notify_settings():
    import notify
    conn = db()
    recips = notify.recipients_by_area(conn)
    return jsonify(
        status=notify.config_status(notify.load_config()),
        recipients=recips,
        sent_count=conn.execute("SELECT COUNT(*) FROM notify_log").fetchone()[0],
        last_sent=conn.execute("SELECT MAX(sent_at) FROM notify_log").fetchone()[0],
    )


@app.put("/api/notify/recipients")
def notify_recipients_update():
    d = body()  # {"*": "a@x, b@y", "前処理": "..."}
    rows_in = []
    for area, text in d.items():
        for email in re.split(r"[\s,;、]+", text or ""):
            if not email:
                continue
            if not EMAIL_RE.match(email):
                abort(400, f"メールアドレスの形式が正しくありません: {email}")
            rows_in.append((area, email))
    conn = db()
    conn.execute("DELETE FROM notify_recipients")
    conn.executemany("INSERT OR IGNORE INTO notify_recipients(area, email) VALUES (?, ?)", rows_in)
    conn.commit()
    return jsonify(ok=True, count=len(rows_in))


@app.get("/api/notify/preview")
def notify_preview():
    import notify
    messages, warnings = notify.build_messages(db(), notify.load_config(), include_all=request.args.get("all") == "1")
    return jsonify(messages=messages, warnings=warnings)


@app.post("/api/notify/test")
def notify_test():
    import notify
    email = (body().get("email") or "").strip()
    if not EMAIL_RE.match(email):
        abort(400, "メールアドレスを入力してください")
    try:
        notify.send_mail(notify.load_config(), [email], "【設備保全】テストメール",
                         "設備保全アプリからのテストメールです。このメールが届いていれば送信設定は正しく動いています。")
    except Exception as e:
        abort(400, f"送信できませんでした: {e}")
    return jsonify(ok=True)


@app.errorhandler(400)
@app.errorhandler(404)
@app.errorhandler(409)
def handle_error(e):
    if request.path.startswith("/api/"):
        return jsonify(error=e.description), e.code
    return e


def ensure_db():
    conn = connect()
    init_db(conn)
    if not conn.execute("SELECT COUNT(*) FROM equipment").fetchone()[0]:
        print("※ データが空です。先に `python import_xls.py` を実行してください。")


if __name__ == "__main__":
    ensure_db()
    app.run(host=os.environ.get("MAINT_HOST", "0.0.0.0"), port=int(os.environ.get("MAINT_PORT", "5000")),
            debug=os.environ.get("MAINT_DEBUG") == "1")
