"""測定機器の Excel 出力・入力（.xlsx）。

    「測定機器」シート … ID 空欄の行は新しい機器、ID のある行はその機器を書き換える。
    「履歴」シート     … 見るためだけのもの（取り込まない）。履歴はアプリの詳細画面で記録する。
    機器は履歴とつながっているので Excel からは削除しない（使わなくなったら状態を「廃棄」にする）。
"""
import datetime as dt
import io
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation

from db import norm
from instruments import SHEETS, STATUSES, first_date
from parts_catalog import _num, _read, _sheet, col_letter

INST_COLS = [  # (キー, Excel の見出し, 列幅)
    ("id", "ID", 7), ("kind", "種類", 10), ("no", "№", 8), ("area", "エリア", 9), ("status", "状態", 8),
    # 見出しは秤量器・ノギスどちらでも通じる名前にする（最大秤量 ＝ ノギスの使用範囲 など）
    ("rank", "ランク", 7), ("ledger_no", "台帳№", 9), ("plant", "プラント・使用エリア", 22), ("process", "工程", 9), ("purpose", "用途", 14),
    ("maker", "メーカー", 12), ("type", "方式・種別", 14), ("model", "型式", 16), ("serial_no", "機器番号・製造番号", 18),
    ("capacity", "最大秤量・使用範囲", 12), ("division", "表示目量・最小表示", 22), ("accuracy_class", "精度等級", 9), ("trade_cert", "取引証明", 9),
    ("purchased", "購入日", 12), ("asset_no", "資産№", 16), ("note", "メモ", 40),
]
TEXT_FIELDS = ["kind", "no", "area", "status", "rank", "ledger_no", "plant", "process", "purpose", "maker", "type", "model",
               "serial_no", "capacity", "division", "accuracy_class", "trade_cert", "asset_no", "note"]
EV_COLS = [("kind", "種類", 10), ("no", "№", 8), ("plant", "プラント", 22), ("event_date", "日付", 12), ("event", "内容", 12),
           ("note", "メモ", 60), ("author", "記入者", 12), ("source", "記録元", 10)]


def purchased_text(i):
    """購入日の表示（読めた日付は 2020/03/02、読めなかった表記はそのまま）。"""
    if i.get("purchased"):
        return i["purchased"].replace("-", "/")
    return i.get("purchased_raw")


def _text(v):
    return norm(v) or None if v is not None else None


def export_workbook(conn, areas):
    insts = [dict(i) for i in conn.execute("SELECT * FROM instruments ORDER BY kind, sort_order IS NULL, sort_order, id")]
    for i in insts:
        i["purchased"] = purchased_text(i)
    evs = [dict(e) for e in conn.execute(
        """SELECT i.kind, i.no, i.plant, e.event_date, e.event, e.note, e.author,
                  CASE e.source WHEN 'xls' THEN '台帳から' ELSE 'アプリ' END AS source
           FROM instrument_events e JOIN instruments i ON i.id = e.instrument_id
           ORDER BY i.kind, i.sort_order, e.event_date IS NULL, e.event_date, e.id""")]
    for e in evs:
        e["event_date"] = (e["event_date"] or "").replace("-", "/") or None

    wb = Workbook()
    ws = wb.active
    ws.title = "測定機器"
    _sheet(ws, INST_COLS, insts)
    for col in ("no", "ledger_no", "serial_no", "asset_no", "purchased", "division"):  # 先頭の0や日付・分数が化けないよう文字列に
        c = col_letter(INST_COLS, col)
        for (cell,) in ws[f"{c}2:{c}{len(insts) + 1}"]:
            cell.number_format = "@"

    kinds = list(dict.fromkeys([*SHEETS, *(i["kind"] for i in insts)]))
    ws2 = wb.create_sheet("選択肢")
    ws2.append(["種類", "状態", "エリア"])
    for n in range(max(len(kinds), len(STATUSES), len(areas))):
        ws2.append([kinds[n] if n < len(kinds) else None, STATUSES[n] if n < len(STATUSES) else None,
                    areas[n] if n < len(areas) else None])
    for c, w in (("A", 12), ("B", 10), ("C", 14)):
        ws2.column_dimensions[c].width = w
    for key, src, strict in (("kind", f"選択肢!$A$2:$A${len(kinds) + 1}", False),  # 種類は新しく増やせる
                             ("status", f"選択肢!$B$2:$B${len(STATUSES) + 1}", True),
                             ("area", f"選択肢!$C$2:$C${len(areas) + 1}", True)):
        dv = DataValidation(type="list", formula1=src, allow_blank=True, showErrorMessage=strict)
        ws.add_data_validation(dv)
        c = col_letter(INST_COLS, key)
        dv.add(f"{c}2:{c}2000")

    ws3 = wb.create_sheet("履歴")
    _sheet(ws3, EV_COLS, evs)

    ws4 = wb.create_sheet("使い方")
    for line in [
        "測定機器の Excel 入力のしかた",
        "",
        "・行を足すと新しい機器として登録されます。ID は空欄のままにしてください。",
        "・ID がある行は、その機器の内容を書き換えます。ID は変えないでください。",
        "・種類と № は必須です。同じ種類の中で № が重ならないようにしてください（例：秤量器 の 前27）。",
        "・状態・エリアはプルダウン（選択肢シート）から選びます。エリアが空欄の機器（予備機など）はどのエリアでも表示されます。",
        "・購入日は 2020/3/2 のように入れます。",
        "・機器は履歴とつながっているので、Excel からは削除できません。使わなくなった機器は状態を「廃棄」にしてください。",
        "・「履歴」シートは見るためだけのものです。書き換えても取り込まれません。",
        "  故障・修理などの記録は、アプリの測定機器の詳細画面から追加してください。",
        "",
        "書き終わったら、アプリの「測定機器」→「Excel取込」でこのファイルを選びます。",
        "取り込む前に、追加・変更される内容が表示されます。",
    ]:
        ws4.append([line])
    ws4["A1"].font = Font(bold=True, size=14)
    ws4.column_dimensions["A"].width = 100
    wb.move_sheet("使い方", offset=-3)
    wb.active = 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _purchased(v, old, errors, where):
    """購入日のセル → (ISO 日付, 元の表記)。触っていなければ今の値のまま。"""
    if isinstance(v, dt.datetime):
        v = v.date()
    if isinstance(v, dt.date):
        return v.isoformat(), v.isoformat()
    s = _text(v)
    if old is not None and s == purchased_text(old):
        return old["purchased"], old["purchased_raw"]
    if not s:
        return None, None
    m = re.fullmatch(r"(\d{4})[/\-.年](\d{1,2})[/\-.月](\d{1,2})日?", s)
    if m:
        try:
            return dt.date(*map(int, m.groups())).isoformat(), s
        except ValueError:
            pass
    iso = first_date(s)
    if not iso:
        errors.append(f"{where}: 購入日「{s}」が日付として読めません（2020/3/2 のように入れてください）")
    return iso, s


def plan_import(conn, data, areas):
    """取り込んだ場合の変更内容を返す（DB は変えない）。"""
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        raise ValueError("Excel ファイル（.xlsx）として読めません。アプリから出力したファイルを使ってください。")
    plan = {"add": [], "update": [], "errors": []}
    errors = plan["errors"]
    if "測定機器" not in wb.sheetnames:
        errors.append("「測定機器」のシートが必要です。測定機器の「Excel出力」で出したファイルを使ってください。")
        return plan
    idx, rows_ = _read(wb["測定機器"], INST_COLS)
    for need in ("kind", "no"):
        if need not in idx:
            errors.append(f"測定機器: 「{dict((k, l) for k, l, _ in INST_COLS)[need]}」の列が見つかりません")
    if errors:
        return plan

    current = {i["id"]: dict(i) for i in conn.execute("SELECT * FROM instruments")}
    owner = {(i["kind"], i["no"]): i["id"] for i in current.values()}  # 種類 + № の重複チェック用
    seen = {}
    for n, r in rows_:
        where = f"{n}行目"
        n_err = len(errors)
        rid = _num(r.get("id"), "ID", errors, where, integer=True)
        if rid is not None and rid not in current:
            errors.append(f"{where}: ID {rid} の測定機器はありません（新しく足すときは ID を空欄に）")
            continue
        old = current.get(rid)
        v = {k: _text(r.get(k)) for k in TEXT_FIELDS if k in idx}
        if not v.get("kind") or not v.get("no"):
            errors.append(f"{where}: 種類と№は必須です")
            continue
        if v.get("status") and v["status"] not in STATUSES:
            errors.append(f"{where}: 状態「{v['status']}」は使えません（{'・'.join(STATUSES)} から選んでください）")
        if v.get("area") and v["area"] not in areas:
            errors.append(f"{where}: エリア「{v['area']}」はありません（選択肢シートから選ぶか、空欄に）")
        key = (v["kind"], v["no"])
        if key in seen:
            errors.append(f"{where}: {v['kind']} の № {v['no']} が {seen[key]}行目と重なっています")
        elif owner.get(key) not in (None, rid):
            errors.append(f"{where}: {v['kind']} の № {v['no']} はほかの行（ID {owner[key]}）で使われています")
        seen[key] = n
        if "purchased" in idx:
            v["purchased"], v["purchased_raw"] = _purchased(r.get("purchased"), old, errors, where)
        if len(errors) > n_err:
            continue
        label = f"{v['kind']} {v['no']}" + (f" {v['plant']}" if v.get("plant") else "")
        if not old:
            v["status"] = v.get("status") or "使用中"
            plan["add"].append({"row": n, "label": label, **v})
        else:
            changes = {k: [old[k], val] for k, val in v.items() if k != "purchased_raw" and old[k] != val}
            if "purchased" in changes or v.get("purchased_raw") != old["purchased_raw"]:
                changes.pop("purchased", None)
                changes["purchased"] = [purchased_text(old), purchased_text(v)]
            if changes:
                plan["update"].append({"id": rid, "row": n, "label": label, "changes": changes,
                                       "values": {k: v[k] for k in v}})
    return plan


def apply_import(conn, plan):
    if plan["errors"]:
        raise ValueError("エラーのある行を直してから取り込んでください")
    now = dt.datetime.now().isoformat(timespec="seconds")
    for a in plan["add"]:
        v = {k: a[k] for k in (*TEXT_FIELDS, "purchased", "purchased_raw") if k in a}
        v["sort_order"] = conn.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM instruments WHERE kind=?", (v["kind"],)).fetchone()[0]
        v |= {"source": "excel", "updated_at": now}
        conn.execute(f"INSERT INTO instruments({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
    for u in plan["update"]:
        sets = {k: u["values"][k] for k in u["changes"]}
        if "purchased" in sets:
            sets["purchased_raw"] = u["values"].get("purchased_raw")
        conn.execute(f"UPDATE instruments SET {', '.join(k + '=?' for k in sets)}, updated_at=? WHERE id=?",
                     [*sets.values(), now, u["id"]])
    conn.commit()
    return {"add": len(plan["add"]), "update": len(plan["update"])}
