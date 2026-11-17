"""ロス調査シート（076- ロス調査シート（関西2026）.xls）をロス（losses）に取り込む・同じ様式で書き出す。

    python losses.py                       # 076- ロス調査シート*.xls を取り込む
    python losses.py 別のファイル.xls
    python losses.py --dry-run             # 取り込まずに読み取った内容と設備の候補を表示

- 取り込むのは月のシート（「11月」「2025年3　月」など）だけ。まとめ（例）・記載すべきロスのシートは読まない
- 1 行 = 1 件。「例」の行と、№・工場しか入っていない行は飛ばす
- 対象設備名とラインから設備台帳の設備を推定して「候補」として入れる（equipment_confirmed = 0）。
  画面の「設備の確認」で確定したときに、設備があればその設備の作業記録（不定期）を作る
- 何度流しても同じ（シート名と№で突き合わせて上書き。確定済みの設備・作業記録へのリンクはそのまま）
"""
import argparse
import datetime as dt
import difflib
import glob
import io
import os
import re
import sys

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from db import connect, init_db, norm

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_GLOB = "076- ロス調査シート*.xls"

PLANT = "関西"
DEFAULT_RATE = 1800  # ロス金額の時間単価（円/h）。シートの式「ロス時間 × 1800」
LOSS_TYPES = ["ライン停止", "検品・手直し", "配合ミス", "廃棄ロス"]
CAUSES = ["設備由来", "原材料由来", "作業者由来", "方法・手順由来", "その他"]
RECURRENCES = ["初発", "再発"]
YES_NO = ["有", "無"]

# シートの列（0 始まり）→ 項目。順番は様式どおり（Excel 出力でもこの順に並べる）
#   (項目, 見出し, 列幅, 種類)  種類: text / date / time / num / int
XLS_COLS = [
    ("no", "№", 5, "int"), ("plant", "工場", 7, "text"), ("occurred_date", "製造日", 11, "date"), ("occurred_time", "発生時刻", 8, "time"),
    ("loss_type", "ロス分類", 12, "text"), ("cause_type", "ロス原因", 13, "text"), ("recurrence", "初発・再発", 8, "text"),
    ("product_class", "分類", 9, "text"), ("line", "ライン", 10, "text"), ("process", "工程", 9, "text"), ("equipment_text", "対象設備名", 18, "text"),
    ("product_code", "コード", 9, "text"), ("product_name", "品名", 24, "text"),
    ("situation", "発生状況 ", 32, "text"), ("direct_cause", "直接の原因", 28, "text"), ("temp_action", "暫定対策", 28, "text"),
    ("perm_action", "恒久対策\n（再発防止対策）", 30, "text"), ("responsible", "実施責任者", 10, "text"), ("confirmer", "確認者\n（L・課長）", 10, "text"),
    ("stop_min", "ライン停止時間\n（分）", 10, "num"), ("stop_people", "ライン配置人数\n（人）", 10, "num"), ("stop_hours", "ライン停止\nロス時間\n（Ｈ）", 10, "num"),
    ("inspect_min", "検品時間\n（分）", 9, "num"), ("inspect_people", "検品配置人数\n（人）", 9, "num"), ("inspect_hours", "検品ロス時間\n（Ｈ）", 9, "num"),
    ("loss_amount", "ロス金額", 11, "int"),
    ("waste_code", "コード（廃棄）", 9, "text"), ("waste_name", "品名(廃棄）", 20, "text"), ("waste_qty", "廃棄または処分数量", 10, "num"),
    ("waste_unit_cost", "原価＠または差損＠", 10, "num"), ("waste_amount", "廃棄または\n差損額（円）", 12, "int"),
    ("note", "", 30, "text"),
    ("gd_sheet", "有無", 6, "text"), ("gd_issued", "発行日", 11, "date"), ("gd_valid", "妥当性\n確認日", 11, "date"), ("gd_effective", "有効性\n確認日", 11, "date"),
    ("ct_sheet", "有無", 6, "text"), ("ct_issued", "発行日", 11, "date"), ("ct_valid", "妥当性\n確認日", 11, "date"), ("ct_effective", "有効性\n確認日", 11, "date"),
]
FIELDS = [c[0] for c in XLS_COLS if c[0] != "no"]
NUM_FIELDS = [c[0] for c in XLS_COLS if c[3] in ("num", "int") and c[0] != "no"]
DATE_FIELDS = [c[0] for c in XLS_COLS if c[3] == "date"]
# 3 行目の見出し（まとまりの名前）: (最初の列, 最後の列, 名前)
GROUPS = [(1, 1, "入力必須"), (2, 2, "入力必須"), (3, 3, "入力必須"), (4, 4, "入力必須"), (5, 5, "入力必須"), (6, 6, "入力必須"),
          (7, 10, "不具合の発生場所"), (11, 12, "対象製品"), (13, 16, "不具合の内容"), (17, 18, "担当者"),
          (19, 21, "ラインの停止"), (22, 24, "検品・手直し"), (25, 25, "ロス時間"), (26, 29, "廃棄または差損の発生"),
          (30, 30, "廃棄ロス"), (31, 31, "備考"), (32, 35, "GDシート"), (36, 39, "出荷停止・トラブル是正処置")]
TYPE_ROW = ["数値", None, "日付", "時刻", "文字(選択）", "文字(選択）", "文字(選択）", "文字", "文字", "文字", None, "コード", "文字",
            "文字", "文字", "文字", "文字", "文字", "文字", "数値", "数値", "計算式", "数値", "数値", "計算式", "計算式", "コード", "文字",
            "数値", "数値", "計算式", "文字", "文字（選択）", "日付", "日付", "日付", "文字（選択）", "日付", "日付", "日付"]
EXAMPLE_ROW = ["例", "寒河江", dt.date(2020, 10, 27), dt.time(10, 5), "ライン停止", "設備由来", "初発", "トンカツ", "ロースライン", "箱詰", "包装機",
               "94510", "ベストデリ　ロースカツ１００Ｐ", "包装機の不具合でチョコ停発生\n念のため、30ｃ/ｓを遡及して包装状態を確認", "センサーの故障",
               "センサー交換", "定期点検の実施", "日東", "太郎L", 20, 15, 5, 3, 30, 1.5, 11700, "94510", "ベストデリ　ロースカツ１００Ｐ", 2, 3000, 6000,
               "遡及した30ｃ/ｓ中、包装異常が2C/S発見された。", "有", dt.date(2021, 6, 1), dt.date(2021, 6, 7), dt.date(2021, 6, 30),
               "有", dt.date(2021, 6, 1), dt.date(2021, 6, 7), dt.date(2021, 6, 30)]


def clean(v):
    if v is None:
        return None
    if isinstance(v, float):
        v = int(v) if v.is_integer() else v
    s = norm(v)
    return s or None


# ---------------------------------------------------------------- 計算
def compute(v, rate):
    """ロス時間（停止分 × 人数 ÷ 60）・ロス金額・廃棄額を入れる。入っていない欄は計算しない。"""
    def hours(m, p):
        return round(v[m] * v[p] / 60, 2) if v.get(m) is not None and v.get(p) is not None else None
    v["stop_hours"] = hours("stop_min", "stop_people")
    v["inspect_hours"] = hours("inspect_min", "inspect_people")
    total = (v["stop_hours"] or 0) + (v["inspect_hours"] or 0)
    v["loss_amount"] = round(total * rate) if v["stop_hours"] is not None or v["inspect_hours"] is not None else None
    v["hourly_rate"] = rate
    if v.get("waste_qty") is not None and v.get("waste_unit_cost") is not None:
        v["waste_amount"] = round(v["waste_qty"] * v["waste_unit_cost"])
    elif "waste_amount" not in v:
        v["waste_amount"] = None
    return v


def get_rate(conn):
    r = conn.execute("SELECT value FROM app_settings WHERE key='loss_hourly_rate'").fetchone()
    return int(r["value"]) if r and str(r["value"]).isdigit() else DEFAULT_RATE


def set_rate(conn, rate):
    conn.execute("INSERT INTO app_settings(key, value) VALUES ('loss_hourly_rate', ?) "
                 "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (str(int(rate)),))


# ---------------------------------------------------------------- 設備の推定
ALIASES = [  # 表記ゆれ（左 → 右にそろえる）
    (r"メガ[ホフ]ォ?ーマー?", "メガフォーマ"), (r"ボックス", "BOX"), (r"号機", "号"), (r"二次|２次", "2次"),
    (r"玉ねぎ|玉ネギ|たまねぎ|タマネギ", "玉葱"), (r"ウェイト", "ウエイト"), (r"コンベアー", "コンベア"),
    (r"ライン", ""), (r"[\s・()（）]", ""),
]
CIRCLED = str.maketrans("①②③④⑤⑥⑦⑧⑨", "123456789")


def eq_key(s):
    s = norm(s).upper().translate(CIRCLED)
    for pat, rep in ALIASES:
        s = re.sub(pat, rep, s)
    return s


def line_letter(text_key, line):
    """設備名の頭（「B成型機」）か、ラインの欄（「Bライン」）から A〜D を拾う。"""
    m = re.match(r"([A-D])(?![A-Z])", text_key)
    if m:
        return m.group(1)
    letters = re.findall(r"[A-D]", eq_key(line or ""))
    return letters[0] if len(set(letters)) == 1 else None


def candidates(equipment, text, line, limit=5):
    """対象設備名に近い設備を (点数, 設備) で返す（点数の高い順）。equipment は台帳の行（dict）のリスト。"""
    first = (text or "").splitlines()[0] if text else ""  # 「A配合機⏎A成型機⏎シャトル」は最初の設備で探す
    base = eq_key(first)
    if not base:
        return []
    letter = line_letter(base, line)
    with_line = base if not letter or base.startswith(letter) else letter + base
    core = base[1:] if letter and base.startswith(letter) else base  # ラインの文字を除いた名前
    out = []
    for e in equipment:
        best = 0.0
        for name in {eq_key(e.get("nickname")), eq_key(e.get("name"))} - {""}:
            s = difflib.SequenceMatcher(None, with_line, name).ratio()
            if name == with_line or name == base:
                s = 1.0
            elif name.startswith(with_line):  # 「B成型機」→「B成型機745」
                s = max(s, 0.92 + 0.05 * len(with_line) / len(name))
            elif len(with_line) >= 3 and with_line in name:
                s = max(s, 0.88 + 0.05 * len(with_line) / len(name))
            elif len(core) >= 3 and core in name:  # 「メガフォーマ」⊂「D成型機メガフォーマ」
                s = max(s, 0.75 + 0.1 * len(core) / len(name))
            elif len(name) >= 3 and name in with_line:  # 「フローズンカッター」⊂「フローズンカッター1号」
                s = max(s, 0.75 + 0.1 * len(name) / len(with_line))
            if ("2次" in name) != ("2次" in with_line):  # 成型機と二次成型機は別の設備
                s -= 0.15
            m = re.match(r"([A-D])(?![A-Z])", name) or re.search(r"(?<![A-Z])([A-D])$", name)  # 「B成型機745」「シャトルコンベアB」
            if letter and m:
                s += 0.05 if m.group(1) == letter else -0.3
            best = max(best, s)
        if e.get("status") == "廃止":
            best -= 0.15
        if best >= 0.45:
            out.append((round(min(best, 1.0), 3), e))
    out.sort(key=lambda x: -x[0])
    return out[:limit]


AUTO_SCORE = 0.7  # これ以上なら候補として最初から選んでおく（確定は画面で人が行う）


def needs_check(v):
    """設備の確認が要るか（対象設備名があるか、設備由来なのに設備がない）。"""
    return bool(v.get("equipment_text")) or v.get("cause_type") == "設備由来"


# ---------------------------------------------------------------- 作業記録との同期
def record_values(v):
    """ロスから作る作業記録（不定期）の中身。"""
    reason = "\n".join(filter(None, [v.get("situation"), f"原因: {v['direct_cause']}" if v.get("direct_cause") else None]))
    return {
        "equipment_id": v.get("equipment_id"), "kind": "不定期", "status": "完了",
        "done_date": v.get("occurred_date"), "downtime_min": int(round(v["stop_min"])) if v.get("stop_min") is not None else None,
        "content": v.get("temp_action") or v.get("situation") or "ロス調査より",
        "reason": reason or None, "countermeasure": v.get("perm_action"), "author": v.get("responsible"),
    }


def create_record(conn, loss):
    vals = record_values(loss) | {"source": "loss"}
    rid = conn.execute(f"INSERT INTO records({','.join(vals)}) VALUES ({','.join('?' * len(vals))})", list(vals.values())).lastrowid
    conn.execute("UPDATE losses SET record_id=? WHERE id=?", (rid, loss["id"]))
    return rid


def sync_record(conn, old, new):
    """ロスを直したら、つながっている作業記録のうち変わった項目だけ書き換える（作業記録の側で直した項目は残す）。"""
    if not new.get("record_id"):
        return
    a, b = record_values(old), record_values(new)
    changed = {k: b[k] for k in b if a[k] != b[k] and not (k == "equipment_id" and not b[k])}
    if changed:
        changed["updated_at"] = dt.datetime.now().isoformat(timespec="seconds")
        conn.execute(f"UPDATE records SET {', '.join(k + '=?' for k in changed)} WHERE id=?", [*changed.values(), new["record_id"]])


# ---------------------------------------------------------------- 読み込み
MONTH_SHEET = re.compile(r"^(?:\d{4}\s*年)?\s*(\d{1,2})\s*月$")


def _cell(book, sh, r, c, kind):
    if c >= sh.ncols:
        return None
    cell = sh.cell(r, c)
    if cell.ctype == xlrd.XL_CELL_EMPTY or (cell.ctype == xlrd.XL_CELL_TEXT and not cell.value.strip()):
        return None
    if kind == "date":
        if cell.ctype in (xlrd.XL_CELL_DATE, xlrd.XL_CELL_NUMBER) and cell.value > 1000:
            return xlrd.xldate_as_datetime(cell.value, book.datemode).date().isoformat()
        return clean(cell.value)  # 日付として読めない表記はそのまま（画面で直してもらう）
    if kind == "time":
        if cell.ctype in (xlrd.XL_CELL_DATE, xlrd.XL_CELL_NUMBER) and 0 <= cell.value < 1:
            mins = round(cell.value * 24 * 60)
            return f"{mins // 60 % 24:02d}:{mins % 60:02d}"
        m = re.match(r"(\d{1,2})[:：時](\d{2})", norm(cell.value))
        return f"{int(m.group(1)):02d}:{m.group(2)}" if m else None
    if kind in ("num", "int"):
        if cell.ctype == xlrd.XL_CELL_NUMBER:
            return cell.value
        try:
            return float(norm(cell.value).replace(",", ""))
        except ValueError:
            return None
    return clean(cell.value)


def read_book(path):
    book = xlrd.open_workbook(path)
    out = []
    for sh in book.sheets():
        m = MONTH_SHEET.match(norm(sh.name))
        if not m or sh.nrows < 5 or clean(sh.cell_value(3, 0)) != "№":
            continue
        for r in range(4, sh.nrows):
            if clean(sh.cell_value(r, 0)) == "例":
                continue
            rec = {f: _cell(book, sh, r, c, kind) for c, (f, _, _, kind) in enumerate(XLS_COLS)}
            if not any(rec.get(f) for f in ("occurred_date", "loss_type", "situation", "product_name")):
                continue  # №と工場だけの行
            no = rec.pop("no")
            rec["source_no"] = int(no) if isinstance(no, (int, float)) else None
            rec["source_sheet"] = sh.name
            rec["source_row"] = r + 1
            for f in ("product_code", "waste_code"):  # 91833.0 → 91833
                if isinstance(rec.get(f), float):
                    rec[f] = str(int(rec[f]))
            for f in ("stop_hours", "inspect_hours", "loss_amount", "waste_amount"):  # 計算式の欄の 0 は「なし」
                if not rec.get(f):
                    rec[f] = None
            for f in ("loss_amount", "waste_amount"):
                if rec.get(f) is not None:
                    rec[f] = int(round(rec[f]))
            rec["plant"] = rec.get("plant") or PLANT
            out.append(rec)
    return out


COLUMNS = FIELDS + ["source_sheet", "source_no", "source_row"]


def save_import(conn, recs, source):
    equipment = [dict(e) for e in conn.execute("SELECT id, code, name, nickname, area, status FROM equipment")]
    now = dt.datetime.now().isoformat(timespec="seconds")
    added = updated = 0
    for rec in recs:
        row = conn.execute("SELECT * FROM losses WHERE source='xls' AND source_sheet=? AND source_row=?",
                           (rec["source_sheet"], rec["source_row"])).fetchone()
        vals = {c: rec.get(c) for c in COLUMNS}
        if not row or not row["equipment_confirmed"]:  # 確定済みの設備は取り込み直しても変えない
            cands = candidates(equipment, rec.get("equipment_text"), rec.get("line"))
            top = next((c for c in cands if c[1]["status"] != "廃止"), None)  # 廃止した設備は最初から選ばない
            vals["equipment_id"] = top[1]["id"] if top and top[0] >= AUTO_SCORE and top is cands[0] else None
            vals["equipment_confirmed"] = 0 if needs_check(rec) else 1
        if row:
            conn.execute(f"UPDATE losses SET {', '.join(k + '=?' for k in vals)}, source=?, updated_at=? WHERE id=?",
                         (*vals.values(), "xls", now, row["id"]))
            sync_record(conn, dict(row), dict(row) | vals)
            updated += 1
        else:
            conn.execute(f"INSERT INTO losses({', '.join(vals)}, source, source_file, created_at) VALUES ({', '.join('?' * len(vals))}, 'xls', ?, ?)",
                         (*vals.values(), source, now))
            added += 1
    conn.commit()
    return added, updated


# ---------------------------------------------------------------- Excel 出力（ロス調査シートと同じ並び。月ごとに 1 シート、新しい月が左）
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
REQ_FILL = PatternFill("solid", fgColor="FCE4D6")
CALC_FILL = PatternFill("solid", fgColor="F2F2F2")


def _month_title(ym):
    return f"{int(ym[5:7])}月"


def export_workbook(conn, from_ym, to_ym):
    rows = [dict(r) for r in conn.execute(
        # 対象設備名が空なら設備台帳の名前（通称）を出す
        """SELECT l.*, COALESCE(l.equipment_text, e.nickname, e.name) AS equipment_text
           FROM losses l LEFT JOIN equipment e ON e.id = l.equipment_id
           WHERE substr(l.occurred_date, 1, 7) BETWEEN ? AND ?
           ORDER BY l.occurred_date, l.occurred_time IS NULL, l.occurred_time, l.id""", (from_ym, to_ym))]
    months = sorted({r["occurred_date"][:7] for r in rows if r.get("occurred_date")}, reverse=True)
    wb = Workbook()
    wb.remove(wb.active)
    if not months:
        ws = wb.create_sheet("ロスなし")
        ws.append([f"{from_ym}〜{to_ym} のロスはありません"])
    titles = [_month_title(ym) for ym in months]
    for ym in months:
        title = _month_title(ym)
        if titles.count(title) > 1:  # 年をまたいで同じ月があるときは年を付ける
            title = f"{ym[:4]}年{title}"
        ws = wb.create_sheet(title)
        _write_month(ws, [r for r in rows if (r.get("occurred_date") or "").startswith(ym)])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _xl_value(field, kind, v):
    if v is None:
        return None
    if kind == "date":
        try:
            return dt.date.fromisoformat(v)
        except ValueError:
            return v
    if kind == "time":
        try:
            return dt.time.fromisoformat(v)
        except ValueError:
            return v
    return v


def _write_month(ws, rows):
    ncol = len(XLS_COLS)
    ws.append(TYPE_ROW)
    first, last = 6, 5 + max(len(rows), 1)
    sums = {21: "V", 24: "Y", 25: "Z", 30: "AE"}
    row2 = [None] * ncol
    row2[1] = row2[7] = row2[10] = row2[18] = "追加"
    row2[20] = row2[23] = "ロス時間計"
    for c, letter in sums.items():
        row2[c] = f"=SUM({letter}{first}:{letter}{last})"
    ws.append(row2)
    ws.append([None] * ncol)
    for c1, c2, label in GROUPS:
        ws.cell(3, c1 + 1, label)
        if c2 > c1:
            ws.merge_cells(start_row=3, start_column=c1 + 1, end_row=3, end_column=c2 + 1)
    ws.append([c[1].rstrip() if c[0] != "note" else None for c in XLS_COLS])
    ws.append(EXAMPLE_ROW)
    for n, r in enumerate(rows, start=1):
        ws.append([n] + [_xl_value(f, kind, r.get(f)) for f, _, _, kind in XLS_COLS[1:]])
    # 書式
    for c, (f, _, width, kind) in enumerate(XLS_COLS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = width
        for r in range(3, ws.max_row + 1):
            cell = ws.cell(r, c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center" if r <= 4 else "top", horizontal="center" if r <= 4 else None, wrap_text=True)
            if r <= 4:
                cell.fill = REQ_FILL if 1 <= c - 1 <= 6 else HEAD_FILL
                cell.font = Font(bold=True, size=10)
            elif kind == "date":
                cell.number_format = "yyyy/m/d"
            elif kind == "time":
                cell.number_format = "h:mm"
            elif f in ("loss_amount", "waste_amount", "waste_unit_cost"):
                cell.number_format = "#,##0"
            if r == 5:
                cell.font = Font(color="808080")
            if f in ("stop_hours", "inspect_hours", "loss_amount", "waste_amount") and r >= 5:
                cell.fill = CALC_FILL
    for c in sums:
        ws.cell(2, c + 1).number_format = "#,##0.##"
        ws.cell(2, c + 1).font = Font(bold=True)
    ws.row_dimensions[4].height = 48
    ws.freeze_panes = "D5"


# ---------------------------------------------------------------- コマンド
def find_default():
    files = sorted(glob.glob(os.path.join(HERE, DEFAULT_GLOB)))
    return files[0] if files else None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xls", nargs="?", default=None)
    ap.add_argument("--dry-run", action="store_true", help="取り込まずに内容を表示する")
    args = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    path = args.xls or find_default()
    if not path:
        sys.exit(f"{DEFAULT_GLOB} が見つかりません")
    recs = read_book(path)
    conn = connect()
    init_db(conn)
    if args.dry_run:
        equipment = [dict(e) for e in conn.execute("SELECT id, code, name, nickname, area, status FROM equipment")]
        for r in recs:
            cands = candidates(equipment, r.get("equipment_text"), r.get("line"), 3)
            c = " / ".join(f"{e.get('nickname') or e['name']}({s:.2f})" for s, e in cands)
            print(f"{r['source_sheet']:>8} {r.get('occurred_date') or '':10} {r.get('loss_type') or '':6} {r.get('cause_type') or '':7} "
                  f"{r.get('line') or '':6} {(r.get('equipment_text') or '').replace(chr(10), '/'):14} → {c}")
        print(f"{len(recs)} 件")
        return
    added, updated = save_import(conn, recs, os.path.basename(path))
    pending = conn.execute("SELECT COUNT(*) FROM losses WHERE equipment_confirmed=0").fetchone()[0]
    print(f"ロス: 追加 {added} 件・更新 {updated} 件（設備の確認待ち {pending} 件）")


if __name__ == "__main__":
    main()
