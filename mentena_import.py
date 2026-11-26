"""MENTENA の「作業計画・履歴一覧」(.xlsx) を定期交換・点検項目（rules）と作業記録（records）に取り込む。

使い方:
    python mentena_import.py [ファイル.xlsx ...] [--dry-run]
    （ファイル省略時は 予防保全/作業計画・履歴一覧_*.xlsx をすべて読む）

MENTENA の 1 行は「繰り返し計画の 1 回分」。同じ整理番号（W-xxxxx）が 1 つの計画なので、
整理番号ごとに 1 項目（rules）にまとめ、完了した回を作業記録（records, source='mentena'）にする。
何度流しても重複しないよう、rules.mentena_no（整理番号）と records.mentena_id（更新用ID）で突き合わせる。
"""
import argparse
import datetime as dt
import glob
import os
import shutil
import sys

import openpyxl

from db import DB_PATH, connect, guess_category, init_db, log_equipment

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE = "mentena"
BY = "Mentena取込"

# 列位置（作業計画・履歴シート。見出しが 4 行あり、5 行目からデータ）
C_UID, C_NO, C_TITLE, C_LOC, C_SUMMARY, C_STATUS, C_RESULT = 1, 3, 6, 7, 8, 9, 10
C_DONE, C_DONE_BY, C_TYPE, C_REPEAT_END, C_EVERY, C_UNIT, C_UPDATED = 11, 12, 14, 28, 29, 30, 46

# MENTENA のロケーション・設備番号 → 設備台帳の id
EQUIPMENT_MAP = {
    "A-00561": 165,  # Aプロパティック745成型機 → A成型機745
    "A-00567": 72,   # C予冷コンベア → C予冷搬送コンベア
    "A-00568": 107, "A-00588": 108, "A-00590": 105, "A-00591": 106,  # C/D/A/B フリーザー
    "A-00573": 114, "A-00574": 115, "A-00575": 116, "A-00576": 117,  # A〜D バケットコンベア
    "A-00577": 118, "A-00578": 119, "A-00579": 120, "A-00580": 121,  # A〜D 包装機
    "A-00581": 122, "A-00582": 123, "A-00583": 124,  # ムルチ 1・2・4 号機
    "A-00584": 125, "A-00585": 126,  # ポリスター 1 号機（小）・2 号機（大）
    "A-00589": 192,  # D予冷コンベア → 予冷D
    "A-00592": 36, "A-00593": 33,  # ソース小袋充填機・小袋用殺菌機（台帳では廃止）
    "A-00594": 30, "A-00595": 31, "A-00596": 32,  # ニーダー 1・2・4 号機（32 は元シート名がニーダー4号機）
    "A-00597": 61,   # Bプロパティック745成型機 → B成型機745
    "A-00598": 59,   # Cハントマン・ヴェルブッファ → ハントマン
    "A-00600": 64, "A-00601": 65,  # 包餡機 火星人・メガフォーマー
    "A-00602": 66, "A-00604": 67,  # A・B 二次成型機
    "A-00605": 77, "A-00606": 170,  # A・B シャトルコンベア
    "A-00609": 80, "A-00610": 81,  # スイットル小・大
    "A-00613": 74, "A-00952": 73, "A-00953": 166,  # D・A・B 焼き機
    "A-00617": 50,   # D配合ミキサー700L → 700L オープンミキサー
    "A-00619": 37,   # 玉ねぎダイサー
    "A-00620": 42, "A-01537": 43,  # フローズンカッター 1・2 号機
    "A-00623": 38,   # サイレントカッター1号機 → サイレントカッター①
    "A-00974": 51,   # サンr2号
    "A-01310": 158,  # チラー冷却水装置
}
# 台帳にないので新しく追加する設備（MENTENA の番号 → 正式名）
NEW_EQUIPMENT = {
    "A-00599": "REX充填・成型機",
    "A-01321": "前工程ストレーナー",
    "A-01467": "金検コンベア(フローズン1号機用)",
    "A-01468": "金検コンベア(フローズン2号機用)",
    "L-00047": "前工程（設備未指定）",
}
AREA_BY_PROCESS = {"前工程": "前処理", "後工程": "後処理", "タレ工程": "タレ工程"}


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def interval_days(every, unit):
    if not every:
        return None
    return int(every) * {"月": 30, "年": 365, "週": 7, "日": 1}.get(unit, 30)


def read_rows(files):
    rows = {}
    for f in files:
        ws = openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0]
        for r in list(ws.iter_rows(values_only=True))[4:]:
            if r[C_UID] and r[C_NO] and not r[0]:  # 先頭列は「削除」フラグ
                rows[r[C_UID]] = r  # 同じ回が複数ファイルにあっても 1 回だけ
    return rows


def ensure_columns(conn):
    for table, col in (("rules", "mentena_no"), ("records", "mentena_id")):
        if col not in {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} TEXT")


def equipment_for(conn, loc, created):
    code = loc.split(" - ")[0].strip()
    if code in EQUIPMENT_MAP:
        return EQUIPMENT_MAP[code]
    if code not in NEW_EQUIPMENT:
        raise SystemExit(f"設備の対応が決まっていません: {loc}")
    name = NEW_EQUIPMENT[code]
    row = conn.execute("SELECT id FROM equipment WHERE name = ? AND source_file = 'Mentena'", (name,)).fetchone()
    if row:
        return row["id"]
    process = loc.split(" - ")[1].split("/")[-1].strip() if " - " in loc else ""
    area = AREA_BY_PROCESS.get(process, "前処理")
    cur = conn.execute(
        "INSERT INTO equipment(name, area, status, category, source_file, note) VALUES (?,?,?,?,?,?)",
        (name, area, "稼働", guess_category(name), "Mentena", f"MENTENA {code} から追加"))
    log_equipment(conn, cur.lastrowid, "追加", by=BY, note=f"MENTENA {loc}")
    created.append(name)
    return cur.lastrowid


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*")
    ap.add_argument("--dry-run", action="store_true", help="書き込まずに件数だけ表示する")
    args = ap.parse_args()
    files = args.files or [f for f in sorted(glob.glob(os.path.join(BASE_DIR, "予防保全", "作業計画・履歴一覧_*.xlsx")))
                           if not os.path.basename(f).startswith("~$")]
    rows = read_rows(files)
    plans = {}
    for r in rows.values():
        plans.setdefault(r[C_NO], []).append(r)

    if not args.dry_run:
        backup = os.path.join(os.path.dirname(DB_PATH), f"maintenance_Mentena取込前_{dt.date.today():%Y%m%d}.db")
        if not os.path.exists(backup):
            shutil.copy2(DB_PATH, backup)
            print("バックアップ:", backup)

    conn = connect()
    init_db(conn)
    ensure_columns(conn)
    today = dt.date.today()
    created, stats = [], {"rule_new": 0, "rule_upd": 0, "rec_new": 0, "rec_upd": 0, "inactive": 0}

    for no, occ in sorted(plans.items()):
        latest = max(occ, key=lambda r: (r[C_UPDATED] or dt.datetime.min, r[C_UID]))
        eq_id = equipment_for(conn, latest[C_LOC], created)
        done = [r for r in occ if r[C_STATUS] == "完了" and to_date(r[C_DONE])]
        last_done = max((to_date(r[C_DONE]) for r in done), default=None)
        end = to_date(latest[C_REPEAT_END])
        active = 0 if end and end < today else 1
        stats["inactive"] += not active
        vals = dict(equipment_id=eq_id, title=latest[C_TITLE].strip(),
                    interval_days=interval_days(latest[C_EVERY], latest[C_UNIT]),
                    last_done_date=last_done.isoformat() if last_done else None,
                    note=(latest[C_SUMMARY] or "").strip() or None, active=active)
        row = conn.execute("SELECT id FROM rules WHERE mentena_no = ?", (no,)).fetchone()
        if row:
            rule_id = row["id"]
            conn.execute("""UPDATE rules SET equipment_id=:equipment_id, title=:title, interval_days=:interval_days,
                                last_done_date=MAX(COALESCE(last_done_date, ''), COALESCE(:last_done_date, '')),
                                note=:note, active=:active WHERE id=:id""", vals | {"id": rule_id})
            conn.execute("UPDATE rules SET last_done_date = NULL WHERE id = ? AND last_done_date = ''", (rule_id,))
            stats["rule_upd"] += 1
        else:
            rule_id = conn.execute(
                """INSERT INTO rules(equipment_id, title, interval_days, last_done_date, note, active, mentena_no)
                   VALUES (:equipment_id, :title, :interval_days, :last_done_date, :note, :active, :no)""",
                vals | {"no": no}).lastrowid
            stats["rule_new"] += 1

        for r in done:
            d = r[C_DONE]
            rec = dict(equipment_id=eq_id, kind="定期" if r[C_TYPE] == "予防保全" else "不定期",
                       done_date=to_date(d).isoformat(), content=r[C_TITLE].strip(),
                       result=(r[C_RESULT] or "").strip() or None, author=(r[C_DONE_BY] or "").strip() or None,
                       rule_id=rule_id, created_at=d.strftime("%Y-%m-%d %H:%M:%S"), mid=str(r[C_UID]))
            if conn.execute("SELECT 1 FROM records WHERE mentena_id = ?", (rec["mid"],)).fetchone():
                conn.execute("""UPDATE records SET equipment_id=:equipment_id, kind=:kind, done_date=:done_date, content=:content,
                                    result=:result, author=:author, rule_id=:rule_id WHERE mentena_id=:mid""", rec)
                stats["rec_upd"] += 1
            else:
                conn.execute("""INSERT INTO records(equipment_id, kind, status, done_date, content, result, author, rule_id,
                                    source, created_at, mentena_id)
                                VALUES (:equipment_id, :kind, '完了', :done_date, :content, :result, :author, :rule_id,
                                    'mentena', :created_at, :mid)""", rec)
                stats["rec_new"] += 1

    if not args.dry_run and (stats["rule_new"] or stats["rec_new"]):
        conn.execute(
            "INSERT INTO activity(at, by, kind, action, target, detail, ref_type) VALUES (?,?,?,?,?,?,?)",
            (dt.datetime.now().isoformat(timespec="seconds"), BY, "交換アラート", "取込", "MENTENA 作業計画・履歴",
             f"点検項目 {stats['rule_new']} 件・作業記録 {stats['rec_new']} 件を追加", "rule"))
    print(f"読込 {len(rows)} 行 / 計画 {len(plans)} 件（{len(files)} ファイル）")
    print(f"点検項目: 追加 {stats['rule_new']}・更新 {stats['rule_upd']}（うち計画終了で停止 {stats['inactive']}）")
    print(f"作業記録: 追加 {stats['rec_new']}・更新 {stats['rec_upd']}")
    if created:
        print("設備を追加:", "、".join(created))
    if args.dry_run:
        conn.rollback()
        print("（--dry-run のため保存していません）")
    else:
        conn.commit()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
