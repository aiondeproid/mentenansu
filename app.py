"""前処理 設備保全アプリ（プロトタイプ）

    python app.py            → http://<サーバー>:5000/
    MAINT_HOST / MAINT_PORT  で待ち受けを変更
"""
import csv
import datetime as dt
import io
import os

import re

from flask import Flask, Response, abort, g, jsonify, request, send_from_directory

from db import connect, init_db
from import_xls import norm

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


def load_rules(conn, equipment_id=None, include_inactive=False):
    sql = """SELECT r.*, e.name AS equipment_name, e.code AS equipment_code, e.current_hours,
                    e.hours_updated
             FROM rules r JOIN equipment e ON e.id = r.equipment_id
             WHERE e.status != '廃止'"""
    args = []
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
    authors = [r[0] for r in conn.execute(
        "SELECT author FROM records WHERE author IS NOT NULL GROUP BY author ORDER BY COUNT(*) DESC")]
    return jsonify(
        record_statuses=RECORD_STATUSES,
        part_statuses=PART_STATUSES,
        part_events=list(PART_EVENT_STATUS),
        authors=authors,
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
           ORDER BY e.status, COALESCE(e.code, 'ZZZ'), e.id"""))
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
    return jsonify(eq)


EQ_FIELDS = ["code", "name", "model", "maker", "vendor_id", "supplier", "asset_no", "serial_no",
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
    sql = """SELECT r.*, e.name AS equipment_name, e.code AS equipment_code, ru.title AS rule_title
             FROM records r JOIN equipment e ON e.id = r.equipment_id
             LEFT JOIN rules ru ON ru.id = r.rule_id WHERE 1=1"""
    params = []
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
            ("downtime_min", "停止時間(分)"), ("hour_meter", "稼働時間(h)"), ("content", "実施内容"),
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
    r = db().execute("SELECT * FROM records WHERE id=?", (rec_id,)).fetchone()
    if not r:
        abort(404)
    return jsonify(dict(r))


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
    conn.execute("DELETE FROM records WHERE id=?", (rec_id,))
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


PART_FIELDS = ["part_type", "serial", "equipment_id", "grind_count", "grind_limit", "status", "note"]


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
    parts = [dict(p) | {"level": part_state(p)} for p in conn.execute(
        "SELECT p.*, e.name AS equipment_name FROM parts p LEFT JOIN equipment e ON e.id = p.equipment_id")]
    open_records = rows(conn.execute(
        """SELECT r.*, e.name AS equipment_name, e.code AS equipment_code FROM records r
           JOIN equipment e ON e.id = r.equipment_id WHERE r.status != '完了'
           ORDER BY r.done_date DESC"""))
    recent = rows(conn.execute(
        """SELECT r.*, e.name AS equipment_name, e.code AS equipment_code FROM records r
           JOIN equipment e ON e.id = r.equipment_id WHERE r.done_date IS NOT NULL
           ORDER BY r.done_date DESC, r.id DESC LIMIT 8"""))
    uncertain = conn.execute("SELECT COUNT(*) FROM records WHERE date_uncertain=1").fetchone()[0]
    return jsonify(
        rules_over=[r for r in rules if r["level"] == "over"],
        rules_soon=[r for r in rules if r["level"] == "soon"],
        rules_unknown=[r for r in rules if r["level"] == "unknown"],
        parts_alert=[p for p in parts if p["level"] in ("over", "soon")],
        open_records=open_records,
        recent=recent,
        uncertain_count=uncertain,
        equipment_count=conn.execute("SELECT COUNT(*) FROM equipment WHERE status!='廃止'").fetchone()[0],
    )


@app.get("/api/analytics")
def analytics():
    conn = db()
    years = request.args.get("years", type=int) or 5
    since = f"{today().year - years + 1}-01-01"
    by_year = rows(conn.execute(
        """SELECT substr(done_date,1,4) AS year,
                  SUM(CASE WHEN kind='定期' THEN 1 ELSE 0 END) AS teiki,
                  SUM(CASE WHEN kind='不定期' THEN 1 ELSE 0 END) AS futeiki,
                  SUM(CASE WHEN kind IS NULL THEN 1 ELSE 0 END) AS other,
                  COALESCE(SUM(downtime_min),0) AS downtime, COALESCE(SUM(cost),0) AS cost
           FROM records WHERE done_date IS NOT NULL GROUP BY year ORDER BY year"""))
    by_equipment = rows(conn.execute(
        """SELECT e.id, e.code, e.name, COUNT(r.id) AS count,
                  SUM(CASE WHEN r.kind='不定期' THEN 1 ELSE 0 END) AS futeiki,
                  COALESCE(SUM(r.downtime_min),0) AS downtime,
                  COALESCE(SUM(r.cost),0) AS cost
           FROM equipment e JOIN records r ON r.equipment_id = e.id
           WHERE r.done_date >= ? GROUP BY e.id ORDER BY downtime DESC, futeiki DESC""", (since,)))
    texts =[f"{r['content'] or ''} {r['reason'] or ''}" for r in conn.execute(
        "SELECT content, reason FROM records WHERE done_date >= ? AND (kind='不定期' OR kind IS NULL)", (since,))]
    keywords = []
    for kw in FAILURE_KEYWORDS:
        pat = re.compile(kw)
        keywords.append({"keyword": kw.replace("|", "・"), "count": sum(1 for t in texts if pat.search(t))})
    keywords.sort(key=lambda k: -k["count"])
    return jsonify(since=since, by_year=by_year, by_equipment=by_equipment, keywords=keywords)


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
