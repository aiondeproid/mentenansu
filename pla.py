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
import os
import re

import xlrd

from db import connect, init_db, norm

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
        # チェック表と紐づいている行は消さずに「点検しない」にする（つながりを残すため）
        conn.execute("""UPDATE pla_items SET active = 0 WHERE source = ?
                        AND id IN (SELECT item_id FROM pla_point_items)""", (source,))
        conn.execute("DELETE FROM pla_items WHERE source = ? AND id NOT IN (SELECT item_id FROM pla_point_items)", (source,))
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
