"""機械メンテナンス履歴（前処理1.xls など）を SQLite に取り込む。

使い方:
    python import_xls.py [ファイル.xls ...] [--reset]
    （ファイル省略時は、このフォルダの *.xls をすべて取り込む）

1 シート = 1 設備。「実施年月日」の見出し行と「設備名」の行を探して位置を合わせるので、
取引業者欄が 2 段あるなどで行がずれたシートも読める。
読めない日付・日付なし・前後関係がおかしい日付は date_uncertain=1（要確認）にして元の表記を残す。
"""
import argparse
import glob
import datetime as dt
import os
import re
import sys
import unicodedata

import xlrd

from db import DB_PATH, connect, guess_category, init_db, make_nickname, norm

# 履歴部分の列位置（全ファイル共通。行位置はシートごとに見出しを探す）
COL_TEIKI, COL_FUTEIKI, COL_DATE, COL_TIME = 0, 2, 4, 8
COL_CONTENT, COL_REASON, COL_RESULT, COL_NOTE, COL_AUTHOR = 11, 16, 20, 27, 34

# ファイル名 → エリア（職場）
AREA_BY_FILE = [
    ("前処理", "前処理"), ("後処理", "後処理"), ("タレ", "タレ工程"), ("原動機", "原動機"),
    ("排水", "排水施設"), ("洗浄", "洗浄関係"), ("X線", "X線・金検"),
]
# 定期項目・部品の初期値は前処理1.xls の内容から作っている
SEED_FILE = "前処理1.xls"

# 履歴と備考から読み取った定期交換・点検ルール（間隔は現場確認前の仮設定を含む）
RULE_SEEDS = [
    # (シート名, タイトル, 間隔日数, 間隔時間h, 前回実施を拾うキーワード, 根拠)
    ("ﾐｰﾄｾﾊﾟﾚｰﾀｰ1号", "フィルタースクリュー交換", 90, 200, "フィルター", "交換実績 150〜450h（平均約200h）から仮設定"),
    ("ﾐｰﾄｾﾊﾟﾚｰﾀｰ2号", "フィルタースクリュー交換", 90, 200, "フィルター", "交換実績 150〜450h（平均約200h）から仮設定"),
    ("ﾐｰﾄｾﾊﾟﾚｰﾀｰ3号", "フィルタースクリュー交換", 90, 200, "フィルター", "交換実績 150〜450h（平均約200h）から仮設定"),
    ("ﾐｰﾄｾﾊﾟﾚｰﾀｰ2号", "オイルシール点検", 180, None, "オイルシール", "グリス漏れによる交換が再発（2024）"),
    ("ﾐｰﾄｾﾊﾟﾚｰﾀｰ3号", "オイルシール点検", 180, None, "オイルシール", "グリス漏れによる交換が再発（2022・2024）"),
    ("ｷｬﾘｱｸﾞﾗｲﾝﾀﾞｰ1号", "キャリアプレート交換", 90, None, "プレート", "2〜4か月ごとの交換実績より"),
    ("ｷｬﾘｱｸﾞﾗｲﾝﾀﾞｰ2号", "キャリアプレート交換", 90, None, "プレート", "2〜4か月ごとの交換実績より"),
    ("ｷｬﾘｱｸﾞﾗｲﾝﾀﾞｰ2号", "メーカー年次点検", 365, None, "^点検", "年1点検推奨（メーカー供給終了部品あり）"),
    ("ヒガシモト配合機", "油圧オイル交換", 365, None, "オイル交換", "メーカー推奨1年毎。製造課オペレーターでも可"),
    ("ヒガシモト配合機", "業者年次点検", 365, None, "点検", "平日は無料とのこと"),
    ("配合機①400kg", "定期点検保全", 730, None, "フルハシ", "ベアリング破損が再発 → 2年ごとの点検保全へ"),
    ("配合機①400kg", "電装BOX乾燥材交換", 180, None, "乾燥材|圧着端子", "漏電対策（間隔は仮設定）"),
    ("配合機②400kg", "下スプロケット交換", 1095, None, "メンテナンス・OH|スプロケ", "2020.6 OH 時に「3年後くらい」を推奨"),
    ("配合機②400kg", "モーターグリスアップ", 90, None, "グリス", "予防保全が出来ていなかった（2021.10）。間隔は仮設定"),
    ("ﾌﾛｰｽﾞﾝｶｯﾀｰ", "駆動・従動ベアリング交換", 730, None, "ベアリング", "「2年くらいが目安？」（23.10.19 備考）"),
    ("玉葱ﾀﾞｲｻｰ", "樹脂ギア交換", 900, None, "O/H|ギア", "交換目安は2〜3年"),
    ("玉葱ﾀﾞｲｻｰ", "プーリー固定ボルト増し締め", 180, None, "イモネジ|増し締め", "半年毎の増し締めを提案（2024.6）"),
    ("ｻｲﾚﾝﾄカッター②", "ナイフカバー樹脂ストライプ交換", 730, None, "ストライ|ナイフカバー", "2023.2 → 2025.2 の実績から仮設定"),
    ("ｻｲﾚﾝﾄｶｯﾀｰ③", "ナイフカバー樹脂ストライプ交換", 730, None, "ストライ|ベルト", "2023.2 → 2025.2 の実績から仮設定"),
]

# 「ﾐｰﾄｾﾊﾟﾚｰﾀｰｽｸﾘｭｰﾌｨﾙﾀｰ」シートと各号機の欄外メモから
PART_SEEDS = [
    # (シリアル, 専用機シート, 研磨回数, 状態, メモ)
    ("7697.1", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ2号", 8, "使用禁止", "肉盛り研磨後に削れるので2度返品。研磨限界にて使用不可"),
    ("10853.1", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ3号", 4, "予備", "交換時 162.5H"),
    ("M1005.1", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ3号", 0, "予備", "21.5.6 戻ってきたフィルターの内側が穴つながり、21.7.14 タカミへ送り返し"),
    ("M1005.2", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ2号", 0, "予備", "MS2号機専用"),
    ("M1005.3", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ2号", 0, "予備", "MS2号機専用"),
    ("M1051.1", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ3号", 0, "予備", "22.2.23 スクリューとフィルターが削れたため M1005.1 に戻す"),
    ("40853.1", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ2号", None, "使用禁止", "25.5.21 フィルター研磨不可のため使用やめ"),
    ("M10653", None, 0, "予備", "23.6.13 予備として購入"),
    ("7961.1", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ1号", 11, "使用禁止", "スクリューの穴径が小さくシャフトが入らない。22.2.15 返却、穴3mmで研磨不可"),
    ("5974.1", None, 13, "使用禁止", "使用は限界（2mmの穴が3mm）"),
    ("6021.2", "ﾐｰﾄｾﾊﾟﾚｰﾀｰ3号", 12, "廃棄", "22.6.8 研磨限界で返品され処分"),
]
PART_TYPE = "フィルタースクリュー"
SERIAL_RE = re.compile(r"M?\d{4,5}\.\d", re.IGNORECASE)


def cell_text(sheet, r, c):
    if r >= sheet.nrows or c >= sheet.ncols:
        return ""
    v = sheet.cell_value(r, c)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    return norm(v)


def parse_date(sheet, r, c, datemode):
    """戻り値: (YYYY-MM-DD or None, 元の表記, 要確認フラグ)"""
    if r >= sheet.nrows or c >= sheet.ncols:
        return None, "", 0
    v = sheet.cell_value(r, c)
    if v in ("", None):
        return None, "", 0
    if isinstance(v, float):
        if 20000 < v < 80000:  # Excel のシリアル値
            y, m, d, *_ = xlrd.xldate_as_tuple(v, datemode)
            return f"{y:04d}-{m:02d}-{d:02d}", str(v), 0
        v = str(v)
    raw = norm(v)
    s = re.sub(r"[/,、]", ".", raw.replace(" ", ""))
    # 年.月.日（後ろに「~15」「・14」など期間の続きがあれば最初の日付を採用）
    m = re.match(r"(\d{2}|\d{4})\.(\d{1,2})\.(\d{1,2})(?!\d)(.*)", s)
    if m and (not m.group(4) or not m.group(4)[0].isdigit()):
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return dt.date(y, mo, d).isoformat(), raw, 0
        except ValueError:
            return None, raw, 1
    # 年.月 のみ（「2020.5月」「2018/5/某」など）→ 月初として並べ、要確認にする
    m = re.fullmatch(r"(\d{4})\.(\d{1,2})(?:月|\.某)?", s)
    if m and 1 <= int(m.group(2)) <= 12:
        return f"{m.group(1)}-{int(m.group(2)):02d}-01", raw, 1
    m = re.fullmatch(r"(\d{4})年", s)
    if m:
        return f"{m.group(1)}-01-01", raw, 1  # 年のみ
    return None, raw, 1


def parse_ym(text):
    """'2011.8' / '2022.11.19' / '2005' → ISO 風文字列"""
    s = norm(text)
    m = re.fullmatch(r"(\d{4})(?:\.(\d{1,2}))?(?:\.(\d{1,2}))?", s)
    if not m:
        return s or None
    parts = [m.group(1)] + [f"{int(g):02d}" for g in m.groups()[1:] if g]
    return "-".join(parts)


def parse_time(raw):
    """実施稼働時間欄 → (停止分, アワーメーターh)"""
    s = norm(raw)
    if not s:
        return None, None
    if re.search(r"停止(は)?(な|ゼロ)", s) or s == "なし":
        return 0, None
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*[hH]", s)
    if m:
        return None, float(m.group(1))
    m = re.search(r"停止(\d+)(?!時間)", s)  # 「660(停止305)」
    if m:
        return int(m.group(1)), None
    m = re.fullmatch(r"(?:停止)?(\d+(?:\.\d+)?)時間", s)  # 「停止4時間」
    if m:
        return round(float(m.group(1)) * 60), None
    ranges = re.findall(r"(\d{1,2}):(\d{2})[~\-](\d{1,2}):(\d{2})", s)  # 「10:00~14:00」（複数は合計）
    if ranges:
        return sum(((int(c) * 60 + int(d)) - (int(a) * 60 + int(b))) % (24 * 60) for a, b, c, d in ranges), None
    m = re.match(r"(\d+(?:\.\d+)?)\s*[(\n]", s)  # 「100(実際の修理時間は690分)」は先頭の数値
    if m:
        return int(float(m.group(1))), None
    minutes = re.findall(r"(\d+)分", s)
    if minutes:
        return sum(int(x) for x in minutes), None
    m = re.fullmatch(r"(\d+(?:\.\d+)?)(?:停止)?", s)  # 数値のみは分とみなす
    if m:
        return int(float(m.group(1))), None
    return None, None


def parse_cost(text):
    m = re.search(r"(\d+)万(?:(\d+)円)?", text or "")
    if m:
        return int(m.group(1)) * 10000 + int(m.group(2) or 0)
    if "百万円" in (text or ""):
        return 1_000_000
    return None


def guess_status(content, result, note):
    joined = f"{content} {result} {note}"
    if "様子" in result:
        return "様子見"
    if "調査待ち" in joined:
        return "保留"
    if "見積もり中" in content:
        return "処理中"
    return "完了"


def vendor_key(name):
    s = unicodedata.normalize("NFKC", name)
    for w in ("株式会社", "有限会社", "(株)", "(有)", " ", "・"):
        s = s.replace(w, "")
    return s


def upsert_vendor(conn, cache, info):
    if not info["name"]:
        return None
    key = vendor_key(info["name"])
    for k, vid in cache.items():
        if k.startswith(key) or key.startswith(k):
            # 情報が多い方で埋める
            cols = [c for c in ("tel", "fax", "address", "contact", "mobile") if info[c]]
            for c in cols:
                conn.execute(f"UPDATE vendors SET {c}=COALESCE(NULLIF({c},''),?) WHERE id=?", (info[c], vid))
            if len(key) > len(k):
                conn.execute("UPDATE vendors SET name=? WHERE id=?", (info["name"], vid))
            return vid
    cur = conn.execute(
        "INSERT INTO vendors(name,tel,fax,address,contact,mobile) VALUES (?,?,?,?,?,?)",
        (info["name"], info["tel"], info["fax"], info["address"], info["contact"], info["mobile"]),
    )
    cache[key] = cur.lastrowid
    return cur.lastrowid


def area_of(path):
    name = os.path.basename(path)
    for key, area in AREA_BY_FILE:
        if key in name:
            return area
    return os.path.splitext(name)[0]


def find_layout(sheet):
    """見出し行（実施年月日）と設備名の行を探す。記録簿の書式でなければ None。"""
    name_row = None
    for r in range(min(sheet.nrows, 30)):
        if cell_text(sheet, r, 1) == "設備名":
            name_row = r
        if cell_text(sheet, r, COL_DATE) == "実施年月日":
            return (r, name_row) if name_row is not None else None
    return None


def has_records(sheet, header):
    return any(cell_text(sheet, r, c) for r in range(header + 1, sheet.nrows)
               for c in (COL_CONTENT, COL_REASON, COL_RESULT))


def skip_reason(sheet, layout):
    if layout is None:
        return "記録簿の書式ではない"
    if sheet.name.startswith("原紙"):
        return "原紙（ひな形）"
    if re.search(r" \(\d+\)$", sheet.name) and not has_records(sheet, layout[0]):
        return "複製シート（履歴なし）"
    return None


def import_sheet(conn, sheet, layout, datemode, vendor_cache, area, source_file):
    header, nr = layout
    t = lambda r, c: cell_text(sheet, r, c)

    # 取引業者（2 段ある場合は 1 つ目を主とし、残りはメモへ）
    vendor_ids, extra_info = [], []
    for r in range(1, nr):
        label = t(r, 14)
        if label == "取引業者名":
            vid = upsert_vendor(conn, vendor_cache, {
                "name": t(r, 16), "tel": t(r, 28), "fax": t(r, 33), "address": "", "contact": "", "mobile": ""})
            if vid:
                vendor_ids.append(vid)
            filled = False
        elif label == "住所" and vendor_ids:
            info = {"address": t(r, 16), "contact": t(r, 28), "mobile": t(r, 33)}
            if not any(info.values()):
                continue
            if not filled:
                for k, v in info.items():
                    if v:
                        conn.execute(f"UPDATE vendors SET {k}=COALESCE(NULLIF({k},''),?) WHERE id=?", (v, vendor_ids[-1]))
                filled = True
            else:
                extra_info.append("連絡先: " + " / ".join(v for v in info.values() if v))
    vendor_id = vendor_ids[0] if vendor_ids else None
    if len(vendor_ids) > 1:
        names = [r["name"] for r in conn.execute(
            f"SELECT name FROM vendors WHERE id IN ({','.join('?' * (len(vendor_ids) - 1))})", vendor_ids[1:])]
        extra_info.append("他の取引業者: " + "、".join(names))

    # 欄外メモ（左上・右側）
    notes = []
    for r in range(0, header):
        for c in [1] + list(range(36, sheet.ncols)):
            v = t(r, c)
            if v and not re.fullmatch(r"\d+(\.0)?", v) and v not in ("設備名", "設備No", "資産№", "メーカー名"):
                notes.append(v)
    notes += extra_info

    name = t(nr, 4)
    asset_no = t(nr + 2, 4)
    if re.fullmatch(r"\d{9,}", name):  # 設備名欄に資産番号が書かれている
        asset_no, name = asset_no or name, ""
    name = name.replace("\n", " ") or sheet.name
    retired = sheet.name.startswith("旧")
    code = t(nr + 1, 4).upper() or None
    cur = conn.execute(
        """INSERT INTO equipment(code,name,model,maker,vendor_id,supplier,asset_no,serial_no,
               manufactured,delivered,area,status,note,sheet_name,source_file,nickname,category)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            code, name + ("（旧機）" if retired and "旧" not in name else ""), t(nr, 13), t(nr + 3, 4),
            vendor_id, t(nr, 24), asset_no, t(nr + 2, 13), parse_ym(t(nr + 1, 13)), parse_ym(t(nr + 3, 13)),
            area, "廃止" if retired else "稼働", "\n".join(notes) or None, sheet.name, source_file,
            make_nickname(sheet.name, name), guess_category(sheet.name, name),
        ),
    )
    eq_id = cur.lastrowid

    count = 0
    latest = None
    for r in range(header + 1, sheet.nrows):
        content, reason, result = t(r, COL_CONTENT), t(r, COL_REASON), t(r, COL_RESULT)
        note, author = t(r, COL_NOTE), t(r, COL_AUTHOR)
        extras = [t(r, c) for c in range(COL_AUTHOR + 1, sheet.ncols) if t(r, c)]
        done_date, date_raw, uncertain = parse_date(sheet, r, COL_DATE, datemode)
        time_raw = t(r, COL_TIME)
        # 本文がない行（例: 「そしておよそ2年後・・・」や日付だけの行）は取り込まない
        if not (content or reason or result):
            continue
        if not done_date:
            uncertain = 1  # 日付なし・読めない日付
        # 直前の行より明らかに古い日付は記入ミスの可能性
        if done_date and latest and done_date < latest and not uncertain:
            if (dt.date.fromisoformat(latest) - dt.date.fromisoformat(done_date)).days > 60:
                uncertain = 1
        if done_date and (not latest or done_date > latest):
            latest = done_date

        kind = "定期" if t(r, COL_TEIKI) else ("不定期" if t(r, COL_FUTEIKI) else None)
        downtime, hours = parse_time(time_raw)
        if extras:
            note = (note + "\n" if note else "") + "\n".join(extras)
        serials = sorted({s.upper() for s in SERIAL_RE.findall(note)}) if "セパレーター" in name else []
        cost = parse_cost(note)
        cur = conn.execute(
            """INSERT INTO records(equipment_id,kind,status,done_date,date_raw,date_uncertain,
                   downtime_min,downtime_raw,hour_meter,content,reason,result,countermeasure,
                   author,cost,part_serial,source)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'xls')""",
            (
                eq_id, kind, guess_status(content, result, note), done_date, date_raw, uncertain,
                downtime, time_raw or None, hours, content, reason, result, note or None,
                author.replace("\n", " ") or None, cost, ",".join(serials) or None,
            ),
        )
        rec_id = cur.lastrowid
        for s in serials:
            pid = ensure_part(conn, s, eq_id)
            conn.execute(
                "INSERT INTO part_events(part_id,event_date,event,hour_meter,record_id,note) VALUES (?,?,?,?,?,?)",
                (pid, done_date, "交換記録", hours, rec_id, "Excel 履歴より"),
            )
        count += 1
    return eq_id, count


def ensure_part(conn, serial, eq_id):
    row = conn.execute("SELECT id FROM parts WHERE serial=?", (serial,)).fetchone()
    if row:
        return row["id"]
    cur = conn.execute(
        "INSERT INTO parts(part_type,serial,equipment_id,grind_count,status,note) VALUES (?,?,?,?,?,?)",
        (PART_TYPE, serial, eq_id, None, "予備", "履歴から自動登録（シリアル表記の確認が必要）"),
    )
    return cur.lastrowid


def seed_parts(conn, eq_by_sheet):
    """研磨管理シートの情報で部品マスタを作る（履歴で自動登録済みなら上書き）。"""
    for serial, sheet, grind, status, note in PART_SEEDS:
        pid = ensure_part(conn, serial, eq_by_sheet.get(sheet))
        conn.execute(
            "UPDATE parts SET equipment_id=?, grind_count=?, status=?, note=? WHERE id=?",
            (eq_by_sheet.get(sheet), grind, status, note, pid),
        )


def seed_rules(conn, eq_by_sheet):
    for sheet, title, days, hours, keyword, note in RULE_SEEDS:
        eq_id = eq_by_sheet.get(sheet)
        if not eq_id:
            print(f"  ! ルールの設備が見つかりません: {sheet}", file=sys.stderr)
            continue
        pattern = re.compile(keyword)
        last = None
        for rec in conn.execute(
            "SELECT done_date, content, reason FROM records WHERE equipment_id=? AND done_date IS NOT NULL AND date_uncertain=0",
            (eq_id,),
        ):
            if pattern.search(rec["content"] or "") or pattern.search(rec["reason"] or ""):
                if not last or rec["done_date"] > last:
                    last = rec["done_date"]
        conn.execute(
            """INSERT INTO rules(equipment_id,title,interval_days,interval_hours,last_done_date,last_hours,keyword,note)
               VALUES (?,?,?,?,?,?,?,?)""",
            (eq_id, title, days, hours, last, 0 if hours else None, keyword, note),
        )


def run(paths, reset):
    if reset and os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    conn = connect()
    init_db(conn)
    if conn.execute("SELECT COUNT(*) FROM equipment").fetchone()[0]:
        print("既にデータがあります。作り直す場合は --reset を付けてください。")
        return
    # 部品の価格表（〇〇部品一覧.xls）は履歴ではないので別に取り込む
    price_lists = [p for p in paths if "部品一覧" in os.path.basename(p)]
    pla_lists = [p for p in paths if "プラ樹脂点検リスト" in os.path.basename(p)]  # プラ点検の点検箇所
    paths = [p for p in paths if p not in price_lists and p not in pla_lists]
    vendor_cache, seed_eq = {}, {}
    total_eq = total_rec = 0
    for path in paths:
        wb = xlrd.open_workbook(path)
        fname, area = os.path.basename(path), area_of(path)
        print(f"■ {fname}（エリア: {area}）")
        for s in wb.sheets():
            layout = find_layout(s)
            why = skip_reason(s, layout)
            if why:
                if s.nrows:
                    print(f"  - {s.name}: 取り込まない（{why}）")
                continue
            eq_id, n = import_sheet(conn, s, layout, wb.datemode, vendor_cache, area, fname)
            if fname == SEED_FILE:
                seed_eq[s.name] = eq_id
            total_eq += 1
            total_rec += n
            print(f"  {s.name}: {n} 件")
    if seed_eq:
        seed_parts(conn, seed_eq)
        seed_rules(conn, seed_eq)
    conn.commit()
    uncertain = conn.execute("SELECT COUNT(*) FROM records WHERE date_uncertain=1").fetchone()[0]
    print(f"設備 {total_eq} 台 / 履歴 {total_rec} 件（日付要確認 {uncertain} 件）を {DB_PATH} に取り込みました")
    from parts_catalog import import_price_list  # 設備を入れたあとでないと紐づけられない
    for path in price_lists:
        items, _ = import_price_list(conn, path)
        print(f"■ {os.path.basename(path)}: 部品一覧に {len(items)} 件")
    from pla import import_list
    for path in pla_lists:
        items, _ = import_list(conn, path)
        print(f"■ {os.path.basename(path)}: プラ点検の点検箇所に {len(items)} 件")


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser()
    ap.add_argument("xls", nargs="*", help="取り込むファイル（省略時はフォルダ内の *.xls すべて）")
    ap.add_argument("--reset", action="store_true", help="既存の DB を削除して取り込み直す")
    args = ap.parse_args()
    run(args.xls or sorted(glob.glob(os.path.join(here, "*.xls"))), args.reset)
