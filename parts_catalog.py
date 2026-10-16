"""部品一覧（価格表）の取り込みと、部品の Excel 出力・入力。

    python parts_catalog.py 前工程ライン部品一覧.xls          価格表を部品一覧に追加する
    python parts_catalog.py 前工程ライン部品一覧.xls --replace  同じファイルから入れた分を入れ直す

価格表の書式（前工程ライン部品一覧.xls）:
    1 シート = 1 区分（Aライン・コンベアネット など）。A 列に機械の見出し、
    B〜D 列に「部品名 / 値段（単価） / 備考」。E 列は数字なら数量、文字なら補足として読む。

アプリからの Excel 出力・入力は openpyxl の .xlsx。
    「部品一覧」シート     … 価格表。ID 空欄は新規、「削除」列に何か書くと削除
    「研磨管理部品」シート … シリアル管理の部品。シリアルで照合し、なければ新規登録
"""
import argparse
import datetime as dt
import io
import os
import re

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from db import connect, init_db, norm

PART_STATUSES = ["使用中", "予備", "研磨中", "使用禁止", "廃棄"]

# 価格表の見出し → 設備台帳の設備（確実に対応が分かるものだけ。ほかは画面で紐づける）
# (シート名の正規表現, 見出しの正規表現, 設備No または設備名)
EQUIPMENT_LINKS = [
    (r"^Aライン$", r"DeJong成型機", "K020"),
    (r"^Bライン$", r"^745成型機", "K021"),
    (r"^Bライン$", r"^火星人成型機", "K024"),
    (r"^Cライン$", r"ハントマン", "ハントマン"),
    (r"^Dライン$", r"DeJong成型機", "K023"),
    (r"^Dライン$", r"^2次成型機", "K029"),
    (r"^Dライン$", r"スイットル ?小", "スイットル小"),
    (r"^Dライン$", r"スイットル ?大", "スイットル大"),
    (r"^Dライン$", r"^メガフォーマー成型機", "K026"),
    (r"^Dライン$", r"^メガフォーマー ?2次成型機", "K030"),
]

CATALOG_COLS = [  # (キー, Excel の見出し, 列幅)
    ("id", "ID", 7), ("group_name", "区分", 14), ("machine", "機械・場所", 24), ("name", "部品名", 40),
    ("unit_price", "単価(円)", 11), ("qty", "数量", 7), ("note", "備考", 40), ("memo", "補足", 30), ("location", "保管場所", 20),
    ("equipment", "紐づけ設備", 30), ("area", "エリア", 10), ("delete", "削除", 6),
]
SERIAL_COLS = [
    ("id", "ID", 7), ("serial", "シリアル", 14), ("part_type", "種別", 18), ("equipment", "専用機", 30),
    ("status", "状態", 10), ("grind_count", "研磨回数", 9), ("grind_limit", "研磨上限", 9), ("location", "保管場所", 20), ("note", "メモ", 50),
]


# ---------------------------------------------------------------- 価格表（.xls）の読み込み
def _text(sheet, r, c):
    return norm(sheet.cell_value(r, c)) if c < sheet.ncols else ""


def parse_price_list(path):
    """価格表を行のリストにする。warnings は内容確認で気づいた点。"""
    wb = xlrd.open_workbook(path)
    items, warnings = [], []
    sheet_names = {norm(s) for s in wb.sheet_names()}
    for sh in wb.sheets():
        group = norm(sh.name)
        title, machine, section_items = None, None, {}
        for r in range(sh.nrows):
            a, b = _text(sh, r, 0), _text(sh, r, 1)
            if a == "部品値段一覧":
                continue
            if a and not b:
                if title is None:
                    title = a
                    if title != group and title in sheet_names:  # 別シートからの写し忘れ
                        warnings.append(f"「{sh.name}」シートの表題が「{title}」になっている（シート名と違う）")
                    continue
                machine = a
                section_items.setdefault(machine, 0)
                continue
            if not b or b == "部品名":
                continue
            price = sh.cell_value(r, 2) if sh.ncols > 2 else ""
            extra = sh.cell_value(r, 4) if sh.ncols > 4 else ""
            m = machine or (title if title and title != group else None)
            item = {
                "group_name": group, "machine": m, "name": b,
                "unit_price": int(price) if isinstance(price, float) else None,
                "qty": extra if isinstance(extra, float) else None,
                "note": _text(sh, r, 3) or None,
                "memo": norm(extra) if isinstance(extra, str) and extra.strip() else None,
                "sort_order": len(items) + 1,
            }
            items.append(item)
            if machine:
                section_items[machine] += 1
            if item["unit_price"] is None:
                warnings.append(f"{group} / {m or '—'} / {b}: 単価が空欄")
        for sec, n in section_items.items():
            if n == 0:
                warnings.append(f"{group} / {sec}: 見出しだけで部品が入っていない")
    return items, warnings


def link_equipment(conn, items):
    eqs = conn.execute("SELECT id, code, name, nickname, area FROM equipment WHERE status != '廃止'").fetchall()
    for it in items:
        for sheet_pat, machine_pat, key in EQUIPMENT_LINKS:
            if re.search(sheet_pat, it["group_name"]) and it["machine"] and re.search(machine_pat, it["machine"]):
                hit = [e for e in eqs if key in (e["code"], e["name"], e["nickname"])]
                if len(hit) == 1:
                    it["equipment_id"] = hit[0]["id"]
                    it["area"] = hit[0]["area"]
                break
    return items


def import_price_list(conn, path, area="前処理", replace=False):
    source = os.path.basename(path)
    items, warnings = parse_price_list(path)
    link_equipment(conn, items)
    if replace:
        conn.execute("DELETE FROM part_catalog WHERE source = ?", (source,))
    elif conn.execute("SELECT 1 FROM part_catalog WHERE source = ?", (source,)).fetchone():
        raise SystemExit(f"{source} は取り込み済みです。入れ直すときは --replace を付けてください。")
    now = dt.datetime.now().isoformat(timespec="seconds")
    for it in items:
        conn.execute(
            """INSERT INTO part_catalog(group_name, machine, name, unit_price, qty, note, memo, equipment_id,
                                        area, sort_order, source, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (it["group_name"], it["machine"], it["name"], it["unit_price"], it["qty"], it["note"], it["memo"],
             it.get("equipment_id"), it.get("area") or area, it["sort_order"], source, now))
    conn.commit()
    return items, warnings


# ---------------------------------------------------------------- 設備の表記 ⇔ ID
def equipment_label(e):
    return " ".join(x for x in (e["code"], e["nickname"] or e["name"]) if x)


class EquipmentResolver:
    """Excel に書かれた「K020 A成型機ディヨン」「ハントマン」などを設備 ID に戻す。"""

    def __init__(self, conn):
        self.eqs = [dict(e) for e in conn.execute("SELECT id, code, name, nickname, area, status FROM equipment")]
        self.labels = {norm(equipment_label(e)): e for e in self.eqs}

    def __call__(self, text):
        t = norm(text)
        if not t:
            return None, None
        if t in self.labels:
            return self.labels[t], None
        live = [e for e in self.eqs if e["status"] != "廃止"]
        by_code = [e for e in live if e["code"] and norm(e["code"]) == t.split(" ")[0]]
        by_name = [e for e in live if t in (norm(e["nickname"]), norm(e["name"]))]
        for hit in (by_code, by_name):
            if len(hit) == 1:
                return hit[0], None
        if by_code or by_name:
            return None, f"「{text}」に当たる設備が複数あります（設備リストシートの名前をそのまま使ってください）"
        return None, f"「{text}」という設備が見つかりません"


# ---------------------------------------------------------------- Excel 出力
def col_letter(cols, key):
    """列定義（cols）の中で key が何列目か（"A", "B" …）。"""
    return get_column_letter([c[0] for c in cols].index(key) + 1)


HEAD_FILL = PatternFill("solid", fgColor="B5001F")
HEAD_FONT = Font(bold=True, color="FFFFFF")


def _sheet(ws, cols, rows):
    ws.append([c[1] for c in cols])
    for cell in ws[1]:
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(vertical="center")
    for r in rows:
        ws.append([r.get(c[0]) for c in cols])
    for i, (_, _, width) in enumerate(cols, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = width
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions


def export_workbook(conn):
    eqs = conn.execute("SELECT * FROM equipment WHERE status != '廃止' ORDER BY area, code IS NULL, code, name").fetchall()
    eq_by_id = {e["id"]: e for e in conn.execute("SELECT * FROM equipment")}
    lab = lambda eid: equipment_label(eq_by_id[eid]) if eid in eq_by_id else None

    wb = Workbook()
    ws = wb.active
    ws.title = "部品一覧"
    cat = [dict(r) for r in conn.execute("SELECT * FROM part_catalog ORDER BY sort_order, id")]
    for r in cat:
        r["equipment"] = lab(r["equipment_id"])
    _sheet(ws, CATALOG_COLS, cat)
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):  # 単価
        row[0].number_format = "#,##0"

    ws2 = wb.create_sheet("研磨管理部品")
    ser = [dict(r) for r in conn.execute("SELECT * FROM parts ORDER BY serial")]
    for r in ser:
        r["equipment"] = lab(r["equipment_id"])
    _sheet(ws2, SERIAL_COLS, ser)

    # 入力しやすいように、設備と状態はプルダウンにする
    ws3 = wb.create_sheet("設備リスト")
    ws3.append(["紐づけ設備・専用機に書く名前", "エリア"])
    for e in eqs:
        ws3.append([equipment_label(e), e["area"]])
    ws3.column_dimensions["A"].width = 40
    eq_range = f"設備リスト!$A$2:$A${len(eqs) + 1}"
    for target, col in ((ws, col_letter(CATALOG_COLS, "equipment")), (ws2, col_letter(SERIAL_COLS, "equipment"))):
        dv = DataValidation(type="list", formula1=eq_range, allow_blank=True, showErrorMessage=False)
        target.add_data_validation(dv)
        dv.add(f"{col}2:{col}2000")
    dv = DataValidation(type="list", formula1='"' + ",".join(PART_STATUSES) + '"', allow_blank=True)
    ws2.add_data_validation(dv)
    st = col_letter(SERIAL_COLS, "status")
    dv.add(f"{st}2:{st}2000")

    ws4 = wb.create_sheet("使い方")
    for line in [
        "部品の Excel 入力のしかた",
        "",
        "■ 部品一覧（価格表）",
        "・行を足すと新しい部品として登録されます。ID は空欄のままにしてください。",
        "・ID がある行は、その部品の内容を書き換えます。ID は変えないでください。",
        "・部品を消すときは「削除」列に ○ を入れます（行を消しただけでは削除されません）。",
        "・紐づけ設備はプルダウン（設備リストシート）から選びます。空欄でもかまいません。",
        "・保管場所には、予備をしまってある場所（例: 部品庫 棚A-2）を書きます。",
        "",
        "■ 研磨管理部品（フィルタースクリューなど）",
        "・シリアルで照合します。新しいシリアルの行は新規登録、ある行は内容を書き換えます。",
        "・使わなくなった部品は、行を消さずに状態を「廃棄」か「使用禁止」にしてください。",
        "",
        "書き終わったら、アプリの「部品・研磨管理」→「Excel取込」でこのファイルを選びます。",
        "取り込む前に、追加・変更・削除される内容が表示されます。",
    ]:
        ws4.append([line])
    ws4["A1"].font = Font(bold=True, size=14)
    ws4.column_dimensions["A"].width = 90
    wb.move_sheet("使い方", offset=-3)
    wb.active = 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- Excel 入力
def _num(v, field, errors, where, integer=False):
    if v in (None, ""):
        return None
    if isinstance(v, str):
        v = norm(v).replace(",", "").replace("¥", "").replace("円", "")
    try:
        f = float(v)
    except ValueError:
        errors.append(f"{where}: {field}「{v}」が数字ではありません")
        return None
    return int(round(f)) if integer else (int(f) if f == int(f) else f)


def _str(v):
    """文字はそのまま使う（全角カッコなどを変えると、触っていない行まで「変更」に見えるため）。"""
    if v is None:
        return None
    if isinstance(v, float) and v == int(v):
        v = int(v)
    s = str(v).strip()
    return s or None


def _read(ws, cols):
    head = [norm(c.value) for c in ws[1]]
    idx = {}
    for key, label, _ in cols:
        if label in head:
            idx[key] = head.index(label)
    out = []
    for n, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(v not in (None, "") for v in row):
            continue
        out.append((n, {k: (row[i] if i < len(row) else None) for k, i in idx.items()}))
    return idx, out


def plan_import(conn, data):
    """取り込んだ場合の変更内容を返す（DB は変えない）。"""
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        raise ValueError("Excel ファイル（.xlsx）として読めません。アプリから出力したファイルを使ってください。")
    resolve = EquipmentResolver(conn)
    plan = {"catalog": {"add": [], "update": [], "delete": []}, "serial": {"add": [], "update": []},
            "errors": [], "sheets": []}
    errors = plan["errors"]

    if "部品一覧" in wb.sheetnames:
        plan["sheets"].append("部品一覧")
        idx, rows_ = _read(wb["部品一覧"], CATALOG_COLS)
        if "name" not in idx:
            errors.append("部品一覧: 「部品名」の列が見つかりません")
            rows_ = []
        current = {r["id"]: dict(r) for r in conn.execute("SELECT * FROM part_catalog")}
        same_key = lambda x: (norm(x["group_name"]), norm(x["machine"]), norm(x["name"]))
        existing = {same_key(r) for r in current.values()}
        for n, r in rows_:
            where = f"部品一覧 {n}行目"
            rid = _num(r.get("id"), "ID", errors, where, integer=True)
            if _str(r.get("delete")):
                if rid is None or rid not in current:  # 新しい行に○、または削除済み
                    continue
                plan["catalog"]["delete"].append({"id": rid, "name": current[rid]["name"], "row": n})
                continue
            if rid is not None and rid not in current:
                errors.append(f"{where}: ID {rid} の部品はありません（新しく足すときは ID を空欄に）")
                continue
            v = {
                "group_name": _str(r.get("group_name")), "machine": _str(r.get("machine")), "name": _str(r.get("name")),
                "unit_price": _num(r.get("unit_price"), "単価", errors, where, integer=True),
                "qty": _num(r.get("qty"), "数量", errors, where),
                "note": _str(r.get("note")), "memo": _str(r.get("memo")), "location": _str(r.get("location")),
                "area": _str(r.get("area")),
            }
            if not v["name"]:
                errors.append(f"{where}: 部品名が空欄です")
                continue
            eq, err = resolve(r.get("equipment"))
            if err:
                errors.append(f"{where}: {err}")
                continue
            v["equipment_id"] = eq["id"] if eq else None
            if not v["area"] and eq:
                v["area"] = eq["area"]
            if rid is None:
                if same_key(v) in existing:  # 同じファイルを2回取り込んだときなど
                    errors.append(f"{where}: 「{v['name']}」は同じ区分・機械にすでにあります（ID のある行を直すか、出力し直したファイルを使ってください）")
                    continue
                existing.add(same_key(v))
                plan["catalog"]["add"].append({"row": n, **v})
            else:
                old = current[rid]
                if "equipment" not in idx:
                    v["equipment_id"] = old["equipment_id"]
                changes = {k: [old[k], val] for k, val in v.items()
                           if (k in idx or k == "equipment_id") and (None if old[k] == "" else old[k]) != val}
                if changes:
                    plan["catalog"]["update"].append({"id": rid, "name": old["name"], "row": n, "changes": changes})

    if "研磨管理部品" in wb.sheetnames:
        plan["sheets"].append("研磨管理部品")
        idx, rows_ = _read(wb["研磨管理部品"], SERIAL_COLS)
        if "serial" not in idx:
            errors.append("研磨管理部品: 「シリアル」の列が見つかりません")
            rows_ = []
        current = {r["serial"]: dict(r) for r in conn.execute("SELECT * FROM parts")}
        seen = set()
        for n, r in rows_:
            where = f"研磨管理部品 {n}行目"
            serial = _str(r.get("serial"))
            if not serial:
                errors.append(f"{where}: シリアルが空欄です")
                continue
            serial = serial.upper()
            if serial in seen:
                errors.append(f"{where}: シリアル {serial} が2回出てきます")
                continue
            seen.add(serial)
            v = {
                "part_type": _str(r.get("part_type")), "status": _str(r.get("status")),
                "grind_count": _num(r.get("grind_count"), "研磨回数", errors, where, integer=True),
                "grind_limit": _num(r.get("grind_limit"), "研磨上限", errors, where, integer=True),
                "location": _str(r.get("location")), "note": _str(r.get("note")),
            }
            if v["status"] and v["status"] not in PART_STATUSES:
                errors.append(f"{where}: 状態「{v['status']}」は使えません（{'・'.join(PART_STATUSES)}）")
                continue
            eq, err = resolve(r.get("equipment"))
            if err:
                errors.append(f"{where}: {err}")
                continue
            v["equipment_id"] = eq["id"] if eq else None
            old = current.get(serial)
            if not old:
                if not v["part_type"]:
                    errors.append(f"{where}: 新しい部品は種別が必要です")
                    continue
                v["status"] = v["status"] or "予備"
                v["grind_limit"] = v["grind_limit"] or 10
                plan["serial"]["add"].append({"row": n, "serial": serial, **v})
            else:
                keymap = {"equipment_id": "equipment"}
                changes = {k: [old[k], val] for k, val in v.items()
                           if keymap.get(k, k) in idx and (None if old[k] == "" else old[k]) != val and not (k == "part_type" and not val)}
                if changes:
                    plan["serial"]["update"].append({"id": old["id"], "serial": serial, "row": n, "changes": changes})

    if not plan["sheets"]:
        errors.append("「部品一覧」か「研磨管理部品」のシートが必要です。アプリから出力したファイルを使ってください。")
    return plan


def apply_import(conn, plan):
    if plan["errors"]:
        raise ValueError("エラーのある行を直してから取り込んでください")
    now = dt.datetime.now().isoformat(timespec="seconds")
    c = plan["catalog"]
    max_order = conn.execute("SELECT COALESCE(MAX(sort_order), 0) FROM part_catalog").fetchone()[0]
    for i, a in enumerate(c["add"], start=1):
        conn.execute(
            """INSERT INTO part_catalog(group_name, machine, name, unit_price, qty, note, memo, location, equipment_id,
                                        area, sort_order, source, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (a["group_name"], a["machine"], a["name"], a["unit_price"], a["qty"], a["note"], a["memo"], a["location"],
             a["equipment_id"], a["area"], max_order + i, "excel", now))
    for u in c["update"]:
        sets = {k: v[1] for k, v in u["changes"].items()}
        conn.execute(f"UPDATE part_catalog SET {', '.join(k + '=?' for k in sets)}, updated_at=? WHERE id=?",
                     [*sets.values(), now, u["id"]])
    for d in c["delete"]:
        conn.execute("DELETE FROM part_catalog WHERE id=?", (d["id"],))
    s = plan["serial"]
    for a in s["add"]:
        conn.execute(
            """INSERT INTO parts(part_type, serial, equipment_id, grind_count, grind_limit, status, location, note)
               VALUES (?,?,?,?,?,?,?,?)""",
            (a["part_type"], a["serial"], a["equipment_id"], a["grind_count"], a["grind_limit"], a["status"],
             a["location"], a["note"]))
    for u in s["update"]:
        sets = {k: v[1] for k, v in u["changes"].items()}
        conn.execute(f"UPDATE parts SET {', '.join(k + '=?' for k in sets)} WHERE id=?", [*sets.values(), u["id"]])
    conn.commit()
    return {k: {kk: len(vv) for kk, vv in v.items()} for k, v in plan.items() if k in ("catalog", "serial")}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("xls", help="部品の価格表（前工程ライン部品一覧.xls など）")
    ap.add_argument("--area", default="前処理", help="設備に紐づかない部品のエリア（既定: 前処理）")
    ap.add_argument("--replace", action="store_true", help="同じファイルから取り込んだ分を消して入れ直す")
    args = ap.parse_args()
    conn = connect()
    init_db(conn)
    items, warnings = import_price_list(conn, args.xls, args.area, args.replace)
    linked = sum(1 for i in items if i.get("equipment_id"))
    print(f"{os.path.basename(args.xls)}: 部品 {len(items)} 件を取り込みました（設備に紐づけ {linked} 件）")
    for w in warnings:
        print("  要確認:", w)
