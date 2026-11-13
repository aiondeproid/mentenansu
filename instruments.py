"""測定機器毎の管理台帳.xls の各シート（秤量器・ノギス・温度計・加熱装置付属温度計）を測定機器（instruments）に取り込む。

    python instruments.py                          # 測定機器毎の管理台帳.xls の全シートを取り込む
    python instruments.py --kind ノギス            # 1 シートだけ
    python instruments.py 別のファイル.xls
    python instruments.py --dry-run                # 取り込まずに読み取った内容を表示

- 1 行 = 1 台。エリアの決め方はシートごと（SHEETS の area）
  - 秤量器: № の頭が「前」→前処理、「後」→後処理。末尾の数字だけの番号は予備機（「予備1」など）
  - ノギス・温度計・加熱装置付属温度計: 使用エリアの言葉で決める（area_by_place）
- 右側の欄（経緯メモ）は日付ごとに区切って履歴（instrument_events）にする。原文は xls_memo に残す
- 何度流しても同じ（№で突き合わせて上書き。履歴は取り込み分だけ入れ直し、アプリで足した履歴は残す）
"""
import argparse
import datetime as dt
import os
import re
import sys
import xlrd

from db import connect, init_db, norm

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FILE = "測定機器毎の管理台帳.xls"

STATUSES = ["使用中", "故障中", "運休", "予備", "廃棄", "欠番"]
EVENTS = ["故障", "修理", "新規・購入", "変更・交換", "運休", "廃棄", "校正", "メモ"]

AREA_BY_PREFIX = {"前": "前処理", "後": "後処理"}


def area_by_place(plant):
    """ノギス・温度計などの「使用エリア」からエリアを決める（上から順に見る）。"""
    plant = plant or ""
    if re.search(r"後処理|凍結", plant):  # 「後処理凍結後C」があるので凍結後は後処理
        return "後処理"
    if re.search(r"試作|品質|品管|クレーム|細菌", plant):
        return "品管"
    if re.search(r"タレ|たれ", plant):
        return "タレ工程"
    if re.search(r"成型|蒸煮|二次|火星人|ライン|肉処理|前処理|玉葱|包餡|受入|解凍|配合", plant):
        return "前処理"
    return None


def split_model_serial(rec):
    """ノギスの型式欄「80409676⏎530-108」は 製造番号 + 型式。数字だけなら製造番号。"""
    v = rec.pop("model", None)
    if not v:
        return
    lines = [x.strip() for x in v.splitlines() if x.strip()]
    if len(lines) == 2 and re.fullmatch(r"\d{6,}", lines[0]):
        rec["serial_no"], rec["model"] = lines
    elif re.fullmatch(r"\d{6,}", v):
        rec["serial_no"] = v
    else:
        rec["model"] = " ".join(lines)


def dash_to_none(rec):
    """「-」だけの欄は空欄にする（温度計 18 のランク・使用エリアなど）。"""
    for k, v in rec.items():
        if v in ("-", "－", "ー"):
            rec[k] = None


# 経緯メモの明らかな書き間違い（日付として読めないもの）
MEMO_FIXES = {"18011.2校正": "18.11.2校正"}


# シートごとの列（他のシートを足すときはここに追加する）。列番号は 0 始まり
#   header: № の列の見出し / numeric_no: 数字だけの番号の付け方 / memo_from: この列から右は経緯メモ
#   remark: 摘要の列（日付があれば経緯メモ。なければ remark_to_type なら type に、そうでなければ経緯メモに） / labels: 画面・Excel での項目名の読み替え
SHEETS = {
    "秤量器": {
        "sheet": "6秤量器", "header": "№", "numeric_no": "予備{}",
        "cols": {"no": 0, "rank": 1, "ledger_no": 2, "plant": 3, "process": 4, "purpose": 5, "maker": 6, "type": 7,
                 "model": 8, "serial_no": 9, "capacity": 10, "division": 11, "accuracy_class": 12, "trade_cert": 13,
                 "purchased_raw": 14, "asset_no": 15},
        "memo_from": 16,
        "area": lambda rec: AREA_BY_PREFIX.get(rec["no"][0]),
        "labels": {"type": "秤の種類", "model": "型式番号", "serial_no": "秤量器番号", "capacity": "最大秤量", "division": "表示目量"},
    },
    "ノギス": {
        "sheet": "1ノギス", "header": "管理番号", "numeric_no": "{}",
        "cols": {"no": 1, "plant": 2, "model": 3, "maker": 4, "purpose": 5, "capacity": 6, "division": 7, "purchased_raw": 8},
        "remark": 9, "remark_to_type": True,
        "memo_from": 10,
        "area": lambda rec: area_by_place(rec.get("plant")),
        "post": split_model_serial,
        "labels": {"plant": "使用エリア", "type": "種別", "model": "型式", "serial_no": "製造番号", "capacity": "使用範囲", "division": "最小表示"},
    },
    "温度計": {
        "sheet": "8温度計", "header": "管理番号", "numeric_no": "{}",
        "cols": {"no": 1, "rank": 2, "plant": 3, "model": 4, "maker": 5, "type": 6, "capacity": 7, "division": 8,
                 "purchased_raw": 9},
        "remark": 10,
        "memo_from": 11,
        "area": lambda rec: area_by_place(rec.get("plant")),
        "post": dash_to_none,
        "labels": {"plant": "使用エリア", "type": "機種名", "model": "型式", "capacity": "使用範囲", "division": "最小表示"},
    },
    "加熱装置付属温度計": {
        "sheet": "9加熱装置付属の温度計", "header": "管理番号", "numeric_no": "{}",
        "cols": {"no": 1, "plant": 2, "type": 3, "model": 4, "maker": 5, "capacity": 6, "division": 7, "purchased_raw": 8},
        "remark": 9,
        "memo_from": 10,
        "area": lambda rec: area_by_place(rec.get("plant")),  # 「Aライン蒸煮」→前処理、「別添たれ」→タレ工程
        "post": dash_to_none,
        "labels": {"plant": "使用エリア", "type": "装置名", "model": "型式", "capacity": "使用範囲", "division": "最小表示"},
    },
}
SPEC_FIELDS = ["rank", "ledger_no", "plant", "process", "purpose", "maker", "type", "model", "serial_no",
               "capacity", "division", "accuracy_class", "trade_cert", "purchased_raw", "asset_no"]


def clean(v):
    if v is None:
        return None
    if isinstance(v, float):
        v = int(v) if v.is_integer() else v
    s = norm(v)  # 全角英数・半角カナだけ直す（①や㈱はそのまま）
    return s or None


# ---------------------------------------------------------------- 日付
# 18.11.8 / 2018.6.26 / 24.08.07 / 20.0513（YY.MMDD）/ 19.720（YY.MDD、要確認）/ 190823（YYMMDD）
DATE_RE = re.compile(r"(?<![\d.])(?:(?:20)?(\d{2})\.(\d{1,2})\.(\d{1,2})|(\d{2})\.(\d{3,4})|(\d{6}))(?![\d])")


def _date(y, m, d):
    try:
        return dt.date(2000 + int(y) % 100, int(m), int(d)).isoformat()
    except ValueError:
        return None


def parse_date_match(m):
    """(ISO 日付, 要確認か) を返す。日付として読めなければ (None, None)。"""
    if m.group(1):
        return _date(m.group(1), m.group(2), m.group(3)), False
    if m.group(4):
        y, md = m.group(4), m.group(5)
        if len(md) == 4:
            return _date(y, md[:2], md[2:]), False
        return _date(y, md[0], md[1:]), True  # 19.720 → 2019-07-20 と読むが要確認
    s = m.group(6)
    return _date(s[:2], s[2:4], s[4:]), False


def first_date(text):
    for m in DATE_RE.finditer(text or ""):
        iso, _ = parse_date_match(m)
        if iso:
            return iso
    return None


# ---------------------------------------------------------------- 経緯メモ → 履歴
EVENT_KEYWORDS = [
    ("廃棄", r"廃棄"),
    ("故障", r"故障|破壊|壊れ|破損|断線|逸脱|誤差|不安定|戻らない|電源が(?:つかない|入らない)|割れ|折れ|取れ|外れ|不適|おかしい"),
    ("修理", r"修理"),
    ("新規・購入", r"新規|新品|購入|発注|入荷|追加"),
    ("運休", r"運休|使用なし"),
    ("変更・交換", r"変更|交換|取替|移動"),
]
SEP = "、,，/／→⇒ 　。"


def classify(text):
    """いちばん先に出てくるキーワードで種類を決める（「故障の為、新規」は故障）。"""
    if "廃棄" in text:  # 「故障で廃棄」は廃棄
        return "廃棄"
    best = None
    for event, pat in EVENT_KEYWORDS:
        m = re.search(pat, text)
        if m and (best is None or m.start() < best[0]):
            best = (m.start(), event)
    return best[1] if best else "メモ"


def split_cells(cells):
    """経緯メモのセルを 1 つずつ split_memo して合わせる（セルの区切りは文の区切りなので、まとめてから読むより正確）。"""
    events, rests = [], []
    for cell in cells:
        ev, rest = split_memo(cell)
        events += ev
        if rest:
            rests.append(rest)
    return events, "、".join(rests) or None


def split_memo(memo):
    """経緯メモを日付ごとに区切る。返り値は (履歴のリスト, 日付のない残りの文)。

    「18.11.8故障18.11.20新規」のように日付が前に付く書き方が多いが、
    「新規19.7.25故障26.10.30」「…取替え190729、…」のように後ろに付く書き方も混じるので、
      - 文が日付以外で始まり日付で終わる → すべて「文 + 日付」とみなす
      - それ以外 → 日付の直後が区切り（、や空白）で、その前に未使用の文があれば「文 + 日付」、なければ「日付 + 文」
    """
    if not memo:
        return [], None
    memo = re.sub(r"\s*\n\s*", " ", memo)
    parts, pos = [], 0  # [文, 日付, 文, 日付, …, 文]
    for m in DATE_RE.finditer(memo):
        iso, unsure = parse_date_match(m)
        if not iso:
            continue
        parts += [memo[pos:m.start()], (iso, unsure, m.group(0))]
        pos = m.end()
    parts.append(memo[pos:])
    texts = parts[0::2]
    dates = parts[1::2]
    strip = lambda s: s.strip(SEP)
    if not dates:
        return [], strip(memo)

    events, used = [], [False] * len(texts)
    suffix_style = strip(texts[0]) != "" and strip(texts[-1]) == ""
    for i, d in enumerate(dates):
        before, after = texts[i], texts[i + 1]
        if suffix_style:
            use = i
        elif (strip(after) == "" or after[0] in SEP) and not used[i] and strip(before):
            use = i
        else:
            use = i + 1
        used[use] = True
        text = strip(texts[use])
        if i == 0 and use == 1 and strip(texts[0]):  # 「故障のため買い替え17.10.31新品使用」→ 前の文も同じ日の記録にする
            used[0] = True
            text = "、".join(filter(None, [strip(texts[0]), text]))
        if text:
            events.append({"event_date": d[0], "date_uncertain": int(d[1]), "date_raw": d[2], "note": text, "event": classify(text)})
    rest = "、".join(strip(t) for t, u in zip(texts, used) if not u and strip(t))
    return events, rest or None


def decide_status(rec, events, rest):
    has_spec = any(rec.get(f) for f in ("maker", "model", "capacity", "plant"))
    memo = rec.get("xls_memo") or ""
    if not has_spec and not memo:
        return "欠番"
    last = max(events, key=lambda e: e["event_date"]) if events else None
    if last and last["event"] == "廃棄":
        return "廃棄"
    if not has_spec and "廃棄" in memo:
        return "廃棄"
    if last and last["event"] == "故障" and not re.search(r"新規|新品|新機種|購入|修理|交換|使用|変更|渡す|貸し出し|→|⇒|を.*に", last["note"]):
        return "故障中"
    if not events and rest and "運休" in rest:
        return "運休"
    if last and last["event"] == "運休" and "使用なし" in last["note"]:  # 温度計 20・21「対策の為、使用なし」
        return "運休"
    if rec["no"].startswith("予備") or "予備" in (rec.get("plant") or ""):  # 「品管（予備）」も予備
        return "予備"
    return "使用中"


# ---------------------------------------------------------------- 読み込み
def read_sheet(path, kind):
    cfg = SHEETS[kind]
    book = xlrd.open_workbook(path)
    sh = book.sheet_by_name(cfg["sheet"])
    no_col = cfg["cols"]["no"]
    header_row = next(r for r in range(sh.nrows) if clean(sh.cell_value(r, no_col)) == cfg["header"])
    out = []
    for r in range(header_row + 1, sh.nrows):
        no = clean(sh.cell_value(r, no_col))
        if not no or no == "例":
            continue
        if re.fullmatch(r"\d+", no):
            no = cfg["numeric_no"].format(no)
        rec = {"kind": kind, "no": no, "sort_order": len(out) + 1}
        for field, c in cfg["cols"].items():
            if field == "no":
                continue
            cell = sh.cell(r, c)
            if cell.ctype == xlrd.XL_CELL_DATE:
                rec[field] = xlrd.xldate_as_datetime(cell.value, book.datemode).date().isoformat()
            else:
                rec[field] = clean(cell.value)
        memo = [clean(sh.cell_value(r, c)) for c in range(cfg["memo_from"], sh.ncols)]
        if "remark" in cfg:  # 摘要: 「TRUSCO F1974-Q1に変更（19.2.25）」は経緯、「デプスゲージ」は種別
            remark = clean(sh.cell_value(r, cfg["remark"]))
            if remark and first_date(remark):
                memo.insert(0, re.sub(r"[(（]\s*([\d.]+)\s*[)）]", r" \1", remark))
            elif remark and cfg.get("remark_to_type"):
                rec["type"] = remark
            elif remark:
                memo.insert(0, remark)
        if cfg.get("post"):
            cfg["post"](rec)
        # 資産№の欄に経緯が書かれている行がある（数字でなければメモの先頭に回す）
        if rec.get("asset_no") and not re.fullmatch(r"[\d\-]+", rec["asset_no"]):
            memo.insert(0, rec.pop("asset_no"))
            rec["asset_no"] = None
        memo = [m for m in memo if m]
        for bad, good in MEMO_FIXES.items():
            memo = [m.replace(bad, good) for m in memo]
        rec["xls_memo"] = "、".join(memo) or None
        rec["purchased"] = first_date(rec.get("purchased_raw"))
        rec["area"] = cfg["area"](rec)
        events, rest = split_cells(memo)
        rec["note"] = rest
        rec["status"] = decide_status(rec, events, rest)
        rec["events"] = events
        out.append(rec)
    return out


COLUMNS = ["kind", "no", "sort_order", *SPEC_FIELDS, "purchased", "area", "status", "note", "xls_memo"]


def save(conn, recs, source):
    now = dt.datetime.now().isoformat(timespec="seconds")
    added = updated = 0
    for rec in recs:
        row = conn.execute("SELECT id FROM instruments WHERE kind=? AND no=?", (rec["kind"], rec["no"])).fetchone()
        vals = [rec.get(c) for c in COLUMNS]
        if row:
            inst_id = row["id"]
            conn.execute(f"UPDATE instruments SET {', '.join(f'{c}=?' for c in COLUMNS)}, source=?, updated_at=? WHERE id=?",
                         (*vals, source, now, inst_id))
            conn.execute("DELETE FROM instrument_events WHERE instrument_id=? AND source='xls'", (inst_id,))
            updated += 1
        else:
            inst_id = conn.execute(f"INSERT INTO instruments({', '.join(COLUMNS)}, source, updated_at) VALUES ({', '.join('?' * len(COLUMNS))}, ?, ?)",
                                   (*vals, source, now)).lastrowid
            added += 1
        conn.executemany(
            "INSERT INTO instrument_events(instrument_id, event_date, date_raw, date_uncertain, event, note, source) VALUES (?,?,?,?,?,?,'xls')",
            [(inst_id, e["event_date"], e["date_raw"], e["date_uncertain"], e["event"], e["note"]) for e in rec["events"]])
    conn.commit()
    return added, updated


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xls", nargs="?", default=os.path.join(HERE, DEFAULT_FILE))
    ap.add_argument("--kind", choices=list(SHEETS), help="取り込むシート（省略するとすべて）")
    ap.add_argument("--dry-run", action="store_true", help="取り込まずに内容を表示する")
    args = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    kinds = [args.kind] if args.kind else list(SHEETS)
    if args.dry_run:
        for r in (r for k in kinds for r in read_sheet(args.xls, k)):
            print(f"{r['no']:6} {r['status']:4} {r.get('area') or '-':4} {r.get('plant') or ''} / {r.get('maker') or ''} {r.get('model') or ''} {r.get('capacity') or ''}")
            for e in r["events"]:
                print(f"         {e['event_date']}{'?' if e['date_uncertain'] else ' '} [{e['event']}] {e['note']}")
            if r["note"]:
                print(f"         メモ: {r['note']}")
        return
    conn = connect()
    init_db(conn)
    for kind in kinds:
        recs = read_sheet(args.xls, kind)
        added, updated = save(conn, recs, os.path.basename(args.xls))
        n_events = sum(len(r["events"]) for r in recs)
        counts = {s: sum(1 for r in recs if r["status"] == s) for s in STATUSES}
        print(f"{kind}: 追加 {added} 台・更新 {updated} 台、履歴 {n_events} 件")
        print("  " + "・".join(f"{s} {n}" for s, n in counts.items() if n))


if __name__ == "__main__":
    main()
