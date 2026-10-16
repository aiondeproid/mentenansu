"""プラ・樹脂箇所の点検（プラ点検）。

    python pla.py                       フォルダ内の「プラ樹脂点検リスト*.xls」を点検箇所として取り込む
    python pla.py ファイル.xls --replace  同じファイルから入れた分を入れ直す（点検記録が付いた箇所は残す）

点検表一覧（プラ樹脂点検リスト 前工程.xls など）の読み方:
    「機械・器具」がある行を見出しとして、列の意味は見出しの文字で決める（ファイルごとに列の位置が違うため）。
    設備名・工程はセル結合や空欄で省略されているので、上の行の値を引き継ぐ。
    「項目から外す」「撤去」と書かれた行や、点検方法・頻度が空の行は「点検しない（停止中）」として残す。
"""
import argparse
import datetime as dt
import glob
import io
import os
import re

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from db import connect, init_db, norm

TIMINGS = ["始業時", "終業時"]
HAZARDS = ["A", "B", "C", "D"]
# 日々の点検に入る点検方法（自主保全基準書は年次）
DAILY_METHODS = ["配合チェック表", "他記録", "始業前点検表", "プラリスク点検表", "洗浄殺菌記録", "タレ記録表", "ライン上部点検表"]
METHOD_PATTERNS = [  # (見出しの正規表現, 保存する名前)
    (r"配合チェック", "配合チェック表"), (r"他記録", "他記録"), (r"始業前", "始業前点検表"),
    (r"プラ.*リスク", "プラリスク点検表"), (r"洗浄殺菌", "洗浄殺菌記録"), (r"ﾀﾚ記録|タレ記録", "タレ記録表"),
    (r"ライン.*上部", "ライン上部点検表"), (r"自主.*保全", "自主保全基準書"),
]
PROCESS_AREA = {"前工程": "前処理", "後工程": "後処理", "タレ工程": "タレ工程"}


def process_of(path):
    name = os.path.basename(path)
    for key in PROCESS_AREA:
        if key[:-2] in name:  # 前 / 後 / タレ
            return key
    return os.path.splitext(name)[0]


# ---------------------------------------------------------------- 点検表一覧（.xls）の読み込み
def _merged_value(sheet):
    """結合セルは左上の値を全体に広げる。"""
    fill = {}
    for r0, r1, c0, c1 in sheet.merged_cells:
        v = sheet.cell_value(r0, c0)
        for r in range(r0, r1):
            for c in range(c0, c1):
                fill[(r, c)] = v
    return lambda r, c: fill.get((r, c), sheet.cell_value(r, c) if c < sheet.ncols else "")


def _clean(v):
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    s = norm(v).replace("\n", " ")
    return re.sub(r" {2,}", " ", s).strip()


def _columns(get, head_row, ncols):
    """見出しの文字から列の意味を決める。"""
    cols, methods = {}, {}
    for c in range(ncols):
        low = _clean(get(head_row, c)).replace(" ", "")
        up = _clean(get(head_row - 1, c)).replace(" ", "")
        label = low or up
        if not label:
            continue
        if label == "工程" or up == "工程":
            cols.setdefault("process", c)
        elif up == "設備名" or label == "設備名":
            cols.setdefault("equipment", c)
        elif label == "機械・器具":
            cols["item"] = c
        elif label == "色":
            cols["color"] = c
        elif label.startswith("材質大括り"):
            cols["material"] = c
        elif label.startswith("％") or label.startswith("%"):
            cols["material_pct"] = c
        elif label.startswith("材質(メーカー)") or label.startswith("材質（メーカー）"):
            cols["material_maker"] = c
        elif label == "材質":
            cols["material_maker2"] = c
        elif label == "危害度" or up == "危害度":
            cols["hazard"] = c
        elif label == "新規" or up == "新規":
            cols["is_new"] = c
        elif label in ("重大性", "可能性", "頻度", "評価"):
            cols[{"重大性": "severity", "可能性": "possibility", "頻度": "frequency", "評価": "evaluation"}[label]] = c
        elif label.startswith("備考"):
            cols["note"] = c
        else:
            for pat, name in METHOD_PATTERNS:
                if re.search(pat, label):
                    methods[c] = name
                    break
    return cols, methods


def parse_list(path):
    """点検表一覧を行のリストにする。warnings は内容確認で気づいた点。"""
    wb = xlrd.open_workbook(path, formatting_info=True)
    sh = wb.sheet_by_index(0)  # 「〇〇継続進捗中」シートが本体（ほかはコピー用・入力フォーム）
    get = _merged_value(sh)
    head = next((r for r in range(min(sh.nrows, 60)) for c in range(sh.ncols) if _clean(get(r, c)) == "機械・器具"), None)
    if head is None:
        raise ValueError(f"{os.path.basename(path)}: 「機械・器具」の見出しが見つかりません")
    cols, methods = _columns(get, head, sh.ncols)
    note_extra = [c for c in range(cols.get("note", sh.ncols) + 1, sh.ncols)]  # 備考の右（撤去の記録など）
    process = process_of(path)
    items, warnings = [], []
    equipment, blank_run = None, 0
    for r in range(head + 1, sh.nrows):
        val = lambda key: _clean(get(r, cols[key])) if key in cols else ""
        eq = val("equipment")
        item = val("item")
        if not item:
            blank_run = blank_run + 1 if not eq else 0
            if eq:
                equipment = eq
            if blank_run > 20 and items:  # 表の下にある別の表（チェック表への追加の記録など）は読まない
                break
            continue
        blank_run = 0
        if eq:
            equipment = eq
        flags = [name for c, name in methods.items() if _clean(get(r, c)) in ("○", "〇", "◯", "o", "O")]
        note = val("note")
        extra = " ".join(x for x in (_clean(get(r, c)) for c in note_extra) if x)
        full_note = " / ".join(x for x in (note, extra) if x) or None
        freq = val("frequency") or None
        maker = val("material_maker") or val("material_maker2") or None
        removed = bool(re.search(r"外す|撤去", note) and not flags) or bool(re.search(r"撤去", extra))
        active = 0 if removed or not (flags or freq) else 1
        items.append({
            "process": process, "equipment": equipment or "（設備名なし）", "item": item,
            "color": val("color") or None, "material": val("material") or None,
            "material_pct": val("material_pct") or None, "material_maker": maker,
            "hazard": val("hazard") or None, "methods": ",".join(dict.fromkeys(flags)) or None,
            "is_new": 1 if val("is_new") in ("○", "〇", "◯") else 0,
            "frequency": freq, "severity": val("severity") or None, "possibility": val("possibility") or None,
            "evaluation": val("evaluation") or None, "note": full_note, "active": active,
            "sort_order": len(items) + 1,
        })
        if not equipment:
            warnings.append(f"{r + 1}行目「{item}」: 設備名がありません")
        if active and not val("hazard"):
            warnings.append(f"{r + 1}行目 {equipment} /「{item}」: 危害度が空欄")
    return items, warnings


def import_list(conn, path, replace=False):
    source = os.path.basename(path)
    items, warnings = parse_list(path)
    if conn.execute("SELECT 1 FROM pla_items WHERE source = ?", (source,)).fetchone():
        if not replace:
            raise SystemExit(f"{source} は取り込み済みです。入れ直すときは --replace を付けてください。")
        # 点検記録のある箇所は消さずに停止中にする（記録を残すため）
        conn.execute("""UPDATE pla_items SET active = 0 WHERE source = ?
                        AND id IN (SELECT item_id FROM pla_results)""", (source,))
        conn.execute("DELETE FROM pla_items WHERE source = ? AND id NOT IN (SELECT item_id FROM pla_results)", (source,))
    now = dt.datetime.now().isoformat(timespec="seconds")
    keys = ["process", "equipment", "item", "color", "material", "material_pct", "material_maker", "hazard", "methods",
            "is_new", "frequency", "severity", "possibility", "evaluation", "note", "active", "sort_order"]
    for it in items:
        conn.execute(f"INSERT INTO pla_items({','.join(keys)}, source, updated_at) VALUES ({','.join('?' * (len(keys) + 2))})",
                     [it[k] for k in keys] + [source, now])
    conn.commit()
    return items, warnings


def is_daily(item):
    """日々（始業時・終業時）の点検に入る箇所か。"""
    if not item["active"]:
        return False
    methods = (item["methods"] or "").split(",")
    return bool(item["frequency"]) or any(m in DAILY_METHODS for m in methods)


# ---------------------------------------------------------------- 点検記録の Excel（月ごと。紙の点検表と同じ形）
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def export_month(conn, process, ym):
    """設備・箇所 × 日付（始・終）の表。○ / × と、× の内容は別シートに出す。"""
    first = dt.date.fromisoformat(ym + "-01")
    last = (first.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
    items = [dict(r) for r in conn.execute(
        "SELECT * FROM pla_items WHERE process = ? ORDER BY sort_order, id", (process,))]
    items = [i for i in items if is_daily(i)]
    res = {(r["item_id"], r["check_date"], r["timing"]): dict(r) for r in conn.execute(
        """SELECT r.* FROM pla_results r JOIN pla_items i ON i.id = r.item_id
           WHERE i.process = ? AND r.check_date BETWEEN ? AND ?""", (process, first.isoformat(), last.isoformat()))}
    reviews = {(r["check_date"], r["timing"]): dict(r) for r in conn.execute(
        "SELECT * FROM pla_reviews WHERE process = ? AND check_date BETWEEN ? AND ?",
        (process, first.isoformat(), last.isoformat()))}
    days = [first + dt.timedelta(days=i) for i in range((last - first).days + 1)]

    wb = Workbook()
    ws = wb.active
    ws.title = f"{first.month}月"
    ws.append([f"プラ・樹脂箇所 点検記録　{process}　{first.year}年{first.month}月"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    head1 = ["設備名", "点検箇所", "危害度"] + sum(([f"{d.day}", ""] for d in days), [])
    head2 = ["", "", ""] + ["始", "終"] * len(days)
    ws.append(head1)
    ws.append(head2)
    for i in range(len(days)):
        c = 4 + i * 2
        ws.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c + 1)
    for c in (1, 2, 3):
        ws.merge_cells(start_row=3, start_column=c, end_row=4, end_column=c)
    fill_head = PatternFill("solid", fgColor="B5001F")
    fill_sun = PatternFill("solid", fgColor="FDECEA")
    fill_ng = PatternFill("solid", fgColor="F8C9C9")
    for row in ws.iter_rows(min_row=3, max_row=4):
        for cell in row:
            cell.fill, cell.font, cell.border = fill_head, Font(bold=True, color="FFFFFF"), BOX
            cell.alignment = Alignment(horizontal="center", vertical="center")
    prev = None
    for it in items:
        row = [it["equipment"] if it["equipment"] != prev else "", it["item"], it["hazard"] or ""]
        prev = it["equipment"]
        for d in days:
            for t in TIMINGS:
                r = res.get((it["id"], d.isoformat(), t))
                row.append({"ok": "○", "ng": "×", "na": "－"}.get(r["result"]) if r else "")
        ws.append(row)
        for i, cell in enumerate(ws[ws.max_row], start=1):
            cell.border = BOX
            if i > 3:
                cell.alignment = Alignment(horizontal="center")
                if cell.value == "×":
                    cell.fill = fill_ng
                    cell.font = Font(bold=True, color="C62828")
                elif days[(i - 4) // 2].weekday() == 6:
                    cell.fill = fill_sun
    # 点検者・確認者の行
    for label, key in (("点検者", "checked_by"), ("確認者", "reviewer")):
        row = ["", label, ""]
        for d in days:
            for t in TIMINGS:
                if key == "reviewer":
                    rv = reviews.get((d.isoformat(), t))
                    row.append(rv["reviewer"] if rv else "")
                else:
                    names = {r["checked_by"] for (iid, cd, tt), r in res.items() if cd == d.isoformat() and tt == t and r["checked_by"]}
                    row.append("、".join(sorted(names)))
        ws.append(row)
        for i, cell in enumerate(ws[ws.max_row], start=1):
            cell.border = BOX
            cell.font = Font(bold=True) if i == 2 else Font(size=8)
            if i > 3:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True, text_rotation=255)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 6
    for i in range(len(days) * 2):
        ws.column_dimensions[get_column_letter(4 + i)].width = 3.2
    ws.freeze_panes = "D5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "3:4"

    ws2 = wb.create_sheet("異常の内容")
    ws2.append(["日付", "時", "設備名", "点検箇所", "危害度", "結果", "内容", "点検者", "記入日時"])
    by_id = {i["id"]: i for i in items}
    for (iid, cd, t), r in sorted(res.items(), key=lambda x: (x[0][1], x[0][2])):
        if r["result"] != "ng" or iid not in by_id:
            continue
        it = by_id[iid]
        ws2.append([cd, t, it["equipment"], it["item"], it["hazard"], "×", r["note"], r["checked_by"], r["checked_at"]])
    for i, w in enumerate((11, 7, 22, 34, 6, 5, 40, 10, 18), start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    for cell in ws2[1]:
        cell.fill, cell.font = fill_head, Font(bold=True, color="FFFFFF")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser()
    ap.add_argument("xls", nargs="*", help="点検表一覧（省略時はフォルダ内の「プラ樹脂点検リスト*.xls」すべて）")
    ap.add_argument("--replace", action="store_true", help="同じファイルから取り込んだ分を入れ直す")
    args = ap.parse_args()
    conn = connect()
    init_db(conn)
    for path in args.xls or sorted(glob.glob(os.path.join(here, "プラ樹脂点検リスト*.xls"))):
        items, warnings = import_list(conn, path, args.replace)
        daily = sum(1 for i in items if is_daily(i))
        print(f"{os.path.basename(path)}: {len(items)} 箇所（日々の点検 {daily}・停止中 {sum(1 for i in items if not i['active'])}）")
        for w in warnings[:30]:
            print("  要確認:", w)
        if len(warnings) > 30:
            print(f"  ほか {len(warnings) - 30} 件")
