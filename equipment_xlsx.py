"""設備台帳の Excel 出力・入力（.xlsx）。

    「設備台帳」シート … ID 空欄の行は新しい設備、ID のある行はその設備を書き換える。
    設備は作業履歴とつながっているので Excel からは削除しない（使わなくなったら状態を「廃止」にする）。
"""
import io

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation

from db import CATEGORIES, guess_category, log_equipment, norm
from parts_catalog import _num, _read, _sheet, _str, col_letter

STATUSES = ["稼働", "廃止"]
EQ_COLS = [  # (キー, Excel の見出し, 列幅)。vendor は取引業者の会社名
    ("id", "ID", 7), ("area", "エリア", 11), ("category", "種類", 16), ("subgroup", "小分類", 16), ("code", "設備No", 9),
    ("nickname", "通称", 22), ("name", "正式名", 34), ("model", "型式", 18), ("maker", "メーカー", 20),
    ("supplier", "納入業者", 18), ("vendor", "取引業者", 24), ("asset_no", "資産No", 15), ("serial_no", "製造番号", 14),
    ("manufactured", "製造年", 9), ("delivered", "納入年月日", 12), ("status", "状態", 7), ("note", "メモ", 40),
]
TEXT_FIELDS = ["area", "category", "subgroup", "code", "nickname", "name", "model", "maker", "supplier", "asset_no", "serial_no",
               "manufactured", "delivered", "status", "note"]
LABELS = {k: label for k, label, _ in EQ_COLS} | {"vendor_id": "取引業者"}


def export_workbook(conn):
    vendors = [dict(v) for v in conn.execute("SELECT id, name FROM vendors ORDER BY name")]
    vname = {v["id"]: v["name"] for v in vendors}
    eqs = [dict(e) for e in conn.execute(
        "SELECT * FROM equipment ORDER BY status = '廃止', area, code IS NULL, code, COALESCE(nickname, name)")]
    for e in eqs:
        e["vendor"] = vname.get(e["vendor_id"])

    wb = Workbook()
    ws = wb.active
    ws.title = "設備台帳"
    _sheet(ws, EQ_COLS, eqs)
    for col in ("asset_no", "serial_no", "code", "manufactured", "delivered"):  # 先頭の0や日付が化けないよう文字列に
        c = col_letter(EQ_COLS, col)
        for (cell,) in ws[f"{c}2:{c}{len(eqs) + 1}"]:
            cell.number_format = "@"

    ws2 = wb.create_sheet("選択肢")
    ws2.append(["種類", "状態", "取引業者"])
    for i in range(max(len(CATEGORIES), len(vendors))):
        ws2.append([CATEGORIES[i] if i < len(CATEGORIES) else None,
                    STATUSES[i] if i < len(STATUSES) else None,
                    vendors[i]["name"] if i < len(vendors) else None])
    for c, w in (("A", 22), ("B", 8), ("C", 34)):
        ws2.column_dimensions[c].width = w
    for key, src in (("category", f"選択肢!$A$2:$A${len(CATEGORIES) + 1}"), ("status", "選択肢!$B$2:$B$3"),
                     ("vendor", f"選択肢!$C$2:$C${len(vendors) + 1}")):
        dv = DataValidation(type="list", formula1=src, allow_blank=True, showErrorMessage=key != "vendor")
        ws.add_data_validation(dv)
        c = col_letter(EQ_COLS, key)
        dv.add(f"{c}2:{c}2000")

    ws3 = wb.create_sheet("使い方")
    for line in [
        "設備台帳の Excel 入力のしかた",
        "",
        "・行を足すと新しい設備として登録されます。ID は空欄のままにしてください。",
        "・ID がある行は、その設備の内容を書き換えます。ID は変えないでください。",
        "・正式名は必須です。通称は現場での呼び名（台帳で大きく表示される名前）です。",
        "・種類・状態・取引業者はプルダウン（選択肢シート）から選びます。",
        "  種類を空欄にすると、名前から自動で決めます。",
        "・小分類は種類の中の枠です（肉処理の「サイレントカッター」など）。同じ言葉を入れた設備が 1 つの枠にまとまります。空欄なら枠なし。",
        "  「X線／Aライン」のように「／」で区切ると、枠の中にさらに枠を作ります。",
        "・取引業者は、アプリの「取引業者」に登録済みの会社名だけ使えます。",
        "・設備は作業履歴とつながっているので、Excel からは削除できません。",
        "  使わなくなった設備は、状態を「廃止」にしてください。",
        "",
        "書き終わったら、アプリの「設備台帳」→「Excel取込」でこのファイルを選びます。",
        "取り込む前に、追加・変更される内容が表示されます。",
    ]:
        ws3.append([line])
    ws3["A1"].font = Font(bold=True, size=14)
    ws3.column_dimensions["A"].width = 80
    wb.move_sheet("使い方", offset=-2)
    wb.active = 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def plan_import(conn, data):
    """取り込んだ場合の変更内容を返す（DB は変えない）。"""
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        raise ValueError("Excel ファイル（.xlsx）として読めません。アプリから出力したファイルを使ってください。")
    plan = {"add": [], "update": [], "errors": [],
            "vendors": {v["id"]: v["name"] for v in conn.execute("SELECT id, name FROM vendors")}}  # 確認画面の表示用
    errors = plan["errors"]
    if "設備台帳" not in wb.sheetnames:
        errors.append("「設備台帳」のシートが必要です。設備台帳の「Excel出力」で出したファイルを使ってください。")
        return plan
    idx, rows_ = _read(wb["設備台帳"], EQ_COLS)
    if "name" not in idx:
        errors.append("設備台帳: 「正式名」の列が見つかりません")
        return plan

    current = {e["id"]: dict(e) for e in conn.execute("SELECT * FROM equipment")}
    vendors = {norm(v["name"]): v["id"] for v in conn.execute("SELECT id, name FROM vendors")}
    # 設備No の重複チェック用（取り込み前から重複しているもの（G003 など）はそのまま認める）
    code_owner = {}
    for e in current.values():
        if e["code"]:
            code_owner.setdefault(norm(e["code"]).upper(), set()).add(e["id"])

    for n, r in rows_:
        where = f"{n}行目"
        rid = _num(r.get("id"), "ID", errors, where, integer=True)
        if rid is not None and rid not in current:
            errors.append(f"{where}: ID {rid} の設備はありません（新しく足すときは ID を空欄に）")
            continue
        v = {k: _str(r.get(k)) for k in TEXT_FIELDS if k in idx}
        if not v.get("name"):
            errors.append(f"{where}: 正式名が空欄です")
            continue
        if v.get("category") and v["category"] not in CATEGORIES:
            errors.append(f"{where}: 種類「{v['category']}」は使えません（選択肢シートから選んでください）")
            continue
        if v.get("status") and v["status"] not in STATUSES:
            errors.append(f"{where}: 状態は「稼働」か「廃止」にしてください（「{v['status']}」）")
            continue
        if "vendor" in idx:
            vn = _str(r.get("vendor"))
            if vn and norm(vn) not in vendors:
                errors.append(f"{where}: 取引業者「{vn}」は登録されていません（先にアプリの「取引業者」に登録してください）")
                continue
            v["vendor_id"] = vendors[norm(vn)] if vn else None
        code = norm(v.get("code") or "").upper()
        old = current.get(rid)
        label = v.get("nickname") or v["name"]
        if not old and any(norm(e["name"]) == norm(v["name"]) and norm(e["nickname"]) == norm(v.get("nickname"))
                           and norm(e["code"] or "").upper() == code for e in current.values()):
            errors.append(f"{where}: 「{label}」はすでに登録されています。直すときは ID のある行を使ってください")
            continue
        if code and (not old or code != norm(old["code"] or "").upper()):
            others = code_owner.get(code, set()) - {rid}
            if others:
                errors.append(f"{where}: 設備No {v['code']} はほかの設備（{current[min(others)]['nickname'] or current[min(others)]['name']}）で使われています")
                continue
            code_owner.setdefault(code, set()).add(rid or -n)
        if not old:
            v.setdefault("status", None)
            v["status"] = v["status"] or "稼働"
            v["area"] = v.get("area") or "前処理"
            v["category"] = v.get("category") or guess_category(v.get("nickname"), v["name"])
            plan["add"].append({"row": n, "label": label, **v})
        else:
            blank = lambda x: None if x == "" else x  # 取り込み済みデータの空文字と空欄は同じものとして扱う
            changes = {k: [old[k], val] for k, val in v.items()
                       if blank(old[k]) != val and not (k == "category" and not val)}
            if changes:
                plan["update"].append({"id": rid, "row": n, "label": label, "changes": changes})
    return plan


def apply_import(conn, plan, by=None):
    if plan["errors"]:
        raise ValueError("エラーのある行を直してから取り込んでください")
    for a in plan["add"]:
        v = {k: a[k] for k in (*TEXT_FIELDS, "vendor_id") if k in a}
        cur = conn.execute(f"INSERT INTO equipment({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
        log_equipment(conn, cur.lastrowid, "Excel取込", by=by, note="Excel取込で追加")
    for u in plan["update"]:
        sets = {k: val[1] for k, val in u["changes"].items()}
        conn.execute(f"UPDATE equipment SET {', '.join(k + '=?' for k in sets)} WHERE id=?", [*sets.values(), u["id"]])
        log_equipment(conn, u["id"], "Excel取込", u["changes"], by=by)
    conn.commit()
    return {"add": len(plan["add"]), "update": len(plan["update"])}
