"""プラ点検のチェック表（紙の「プラスチック部品チェック表」）の取り込み。

    python pla_sheets.py                 「前工程 チェック表」などのフォルダの .xls をすべて取り込む
    python pla_sheets.py --replace       入れ直す（点検記録のあるチェック表・点検箇所は残す）

.xls の写真は xlrd では読めないので、Excel（このPCにインストール済みのもの）で .xlsx に変換してから読む。
    1 シート = 1 チェック表。「始業時」の見出しの下から「確認者」の行までが点検箇所（1 行 = 1 箇所）。
    A 列 = 点検箇所（場所）、B 列 = 項目。A 列に貼られた写真はその行の写真、
    それ以外（表の右や下）に貼られた写真はチェック表全体の「場所の写真」として取り込む。
    「備考」の下の行は改定履歴としてメモに入れる。
各箇所はプラリスト（pla_items）の同じ設備の行と、項目名の文字で自動的に紐づける（画面で直せる）。
"""
import argparse
import datetime as dt
import glob
import hashlib
import io
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from db import BASE_DIR, DB_PATH, connect, init_db

PHOTO_DIR = os.path.join(os.path.dirname(DB_PATH), "pla_photos")
FOLDERS = {"前工程": "前工程 チェック表", "後工程": "後工程 チェック表", "タレ工程": "タレ工程チェック表"}
TIMINGS = ["始業時", "終業時"]


# ---------------------------------------------------------------- .xls → .xlsx（Excel で変換）
def convert_with_excel(paths, out_dir):
    """Excel を裏で動かして .xlsx に保存する。元のファイルは読み取り専用で開くので変わらない。"""
    listing = os.path.join(out_dir, "list.txt")
    with open(listing, "w", encoding="utf-8-sig") as f:
        for i, p in enumerate(paths):
            f.write(f"{os.path.abspath(p)}\t{os.path.join(out_dir, f'{i}.xlsx')}\n")
    ps = f"""
$x = New-Object -ComObject Excel.Application; $x.DisplayAlerts = $false; $x.Visible = $false; $x.AskToUpdateLinks = $false
foreach ($line in Get-Content -LiteralPath '{listing}' -Encoding UTF8) {{
  $src, $dst = $line -split "`t"
  try {{ $wb = $x.Workbooks.Open($src, 0, $true); $wb.SaveAs($dst, 51); $wb.Close($false) }} catch {{ Write-Output ("NG " + $src + " " + $_.Exception.Message) }}
}}
$x.Quit()
"""
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True, encoding="cp932", errors="replace")
    if r.stdout.strip():
        print(r.stdout.strip())
    return {p: os.path.join(out_dir, f"{i}.xlsx") for i, p in enumerate(paths)}


# ---------------------------------------------------------------- 文字の照合
def key(s):
    """照合用に、全角半角・空白・個数（×2 など）・記号をそろえる。"""
    s = unicodedata.normalize("NFKC", str(s or "")).lower()
    s = re.sub(r"[×x]\s*\d+(本|個|枚|箇所|ヶ所|セット|組)?", "", s)
    s = re.sub(r"[\s、,。・:：※()（）「」\[\]～~\-ー①-⑳]", "", s)
    return s.replace("号機", "号")


def _lcs(a, b):
    best = 0
    for i in range(len(a)):
        for j in range(i + best + 1, len(a) + 1):
            if a[i:j] in b:
                best = j - i
            else:
                break
    return best


def candidate_equipment(names, *texts):
    """チェック表の題名・ファイル名から、プラリストのどの設備の表かを推定する（号機の番号も合わせる）。"""
    k = "".join(key(t) for t in texts)
    digits = set(re.findall(r"\d+", k))
    out = []
    for e in names:
        ke = key(e)
        if _lcs(ke, k) < max(3, int(len(ke) * 0.6)):
            continue
        nums = set(re.findall(r"\d+", ke))
        if "2次" in ke:  # 「2次成型」の 2 は号機の番号ではない
            nums.discard("2")
        if nums and digits and not nums & digits:
            continue
        out.append(e)
    return out


def match_items(items, location, text):
    """箇所の文字（場所＋項目の各行）とプラリストの項目名が重なるものを返す。"""
    lines = [key(x) for x in re.split(r"[\n、/]", f"{text}") if len(key(x)) >= 2]
    lines += [no_color(l) for l in lines if len(no_color(l)) >= 2]  # 「ローター(白)」→「ローター」でも探す
    hits = []
    for it in items:
        ki = key(it["item"])
        if any(l == ki or no_color(ki) == l or (len(l) >= 3 and l in ki) or (len(ki) >= 3 and ki in l) for l in lines):
            hits.append(it["id"])
    return hits


COLORS = "白|黒|赤|青|緑|黄|茶|灰|透明|グレ|橙|オレンジ|ピンク|紫|銀|乳白色|アイボリ"


def no_color(k):
    return re.sub(rf"({COLORS})+$", "", k)


# ---------------------------------------------------------------- チェック表の読み込み
def _s(v):
    return str(v).strip() if v is not None else ""


def parse_sheet(ws):
    """1 シートを {title, points[{row, location, text}], notes, photos[(row, col, ext, bytes)]} にする。なければ None。"""
    hdr = next((r for r in range(1, 16) for c in range(1, 8) if _s(ws.cell(r, c).value) == "始業時"), None)
    if not hdr:
        return None
    title = next((_s(ws.cell(r, c).value) for r in range(2, hdr) for c in range(1, 4)
                  if _s(ws.cell(r, c).value) and "チェック表" not in _s(ws.cell(r, c).value)
                  and "点検日" not in _s(ws.cell(r, c).value)), "") or ws.title
    points, notes, loc, end = [], [], "", None
    for r in range(hdr + 1, ws.max_row + 1):
        a, b = _s(ws.cell(r, 1).value), _s(ws.cell(r, 2).value)
        if "確認者" in a or "確認者" in b or a == "備考":
            end = r
            break
        if a and b:
            loc = a
        if b:
            points.append({"row": r, "location": loc, "text": b})
        elif a:  # A 列だけの行（「アドニールスプレー 壊れ・欠け」など）はそれ自体が 1 箇所
            points.append({"row": r, "location": "", "text": a})
            loc = ""
    if end:
        for r in range(end, ws.max_row + 1):
            line = " ".join(_s(ws.cell(r, c).value) for c in range(1, ws.max_column + 1) if _s(ws.cell(r, c).value))
            line = re.sub(r"^備考\s*", "", line.replace("　", " ")).strip()
            if line and "確認者" not in line:
                notes.append(line)
    photos = []
    for im in ws._images:
        try:
            a = im.anchor._from
            photos.append((a.row + 1, a.col + 1, (im.format or "jpeg").lower(), im._data()))
        except Exception:  # 図形や EMF など読めないものは飛ばす
            continue
    return {"title": re.sub(r"\s+", " ", title), "points": points, "notes": notes, "photos": photos}


def _save_photo(data, ext):
    os.makedirs(PHOTO_DIR, exist_ok=True)
    ext = {"jpeg": "jpg", "jpg": "jpg", "png": "png", "gif": "gif"}.get(ext, ext)
    name = hashlib.sha1(data).hexdigest()[:20] + "." + ext
    path = os.path.join(PHOTO_DIR, name)
    if not os.path.exists(path):
        with open(path, "wb") as f:
            f.write(data)
    return name


def sheet_is_current(file_name, sheet_name, all_names):
    """古い版・撤去済み・予備などは「使わない」にしておく（画面で戻せる）。"""
    if re.search(r"撤去", file_name):
        return False
    if re.search(r"旧|予備|履歴|^Sheet\d", sheet_name):
        return False
    m = re.match(r"(.+?)\s*\(\d+\)$", sheet_name)
    if m and m.group(1).strip() in {n.strip() for n in all_names}:
        return False
    return True


def import_folders(conn, root=BASE_DIR, replace=False, xlsx_dir=None):
    """3 つのフォルダのチェック表を取り込む。xlsx_dir に変換済みの .xlsx があればそれを使う。"""
    paths = {proc: sorted(glob.glob(os.path.join(root, folder, "*.xls"))) for proc, folder in FOLDERS.items()}
    if conn.execute("SELECT 1 FROM pla_sheets WHERE source = 'xls'").fetchone():
        if not replace:
            raise SystemExit("チェック表は取り込み済みです。入れ直すときは --replace を付けてください。")
        keep = "SELECT p.sheet_id FROM pla_points p JOIN pla_checks c ON c.point_id = p.id"
        conn.execute(f"UPDATE pla_sheets SET active = 0 WHERE source = 'xls' AND id IN ({keep})")
        drop = f"SELECT id FROM pla_sheets WHERE source = 'xls' AND id NOT IN ({keep})"
        conn.execute(f"DELETE FROM pla_point_items WHERE point_id IN (SELECT id FROM pla_points WHERE sheet_id IN ({drop}))")
        conn.execute(f"DELETE FROM pla_photos WHERE sheet_id IN ({drop})")
        conn.execute(f"DELETE FROM pla_points WHERE sheet_id IN ({drop})")
        conn.execute(f"DELETE FROM pla_sheets WHERE id IN ({drop})")
    tmp = None
    if xlsx_dir:
        conv = {p: os.path.join(xlsx_dir, FOLDERS[proc], os.path.splitext(os.path.basename(p))[0] + ".xlsx")
                for proc, ps in paths.items() for p in ps}
    else:
        tmp = tempfile.mkdtemp(prefix="pla_xlsx_")
        conv = convert_with_excel([p for ps in paths.values() for p in ps], tmp)
    now = dt.datetime.now().isoformat(timespec="seconds")
    order = conn.execute("SELECT COALESCE(MAX(sort_order), 0) FROM pla_sheets").fetchone()[0]
    stats = {"sheets": 0, "points": 0, "linked": 0, "photos": 0, "inactive": 0}
    for proc, ps in paths.items():
        items = [dict(r) for r in conn.execute("SELECT id, equipment, item FROM pla_items WHERE process = ? AND active = 1", (proc,))]
        eq_names = sorted({i["equipment"] for i in items})
        for p in ps:
            x = conv.get(p)
            if not x or not os.path.exists(x):
                print("  読めませんでした:", os.path.basename(p))
                continue
            wb = load_workbook(x)
            fname = os.path.basename(p)
            for ws in wb.worksheets:
                sh = parse_sheet(ws)
                if not sh or not sh["points"]:
                    continue
                active = 1 if sheet_is_current(fname, ws.title, wb.sheetnames) else 0
                order += 1
                cur = conn.execute(
                    """INSERT INTO pla_sheets(process, file, sheet_name, title, active, sort_order, note, source, updated_at)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (proc, fname, ws.title, sh["title"], active, order, "\n".join(sh["notes"]) or None, "xls", now))
                sid = cur.lastrowid
                stats["sheets"] += 1
                stats["inactive"] += 1 - active
                cand = candidate_equipment(eq_names, sh["title"], os.path.splitext(fname)[0], ws.title)
                pool = [i for i in items if i["equipment"] in cand]
                rows = [pt["row"] for pt in sh["points"]]
                pids = []
                for n, pt in enumerate(sh["points"], start=1):
                    pid = conn.execute("INSERT INTO pla_points(sheet_id, location, item_text, sort_order) VALUES (?,?,?,?)",
                                       (sid, pt["location"] or None, pt["text"], n)).lastrowid
                    pids.append(pid)
                    stats["points"] += 1
                    linked = match_items(pool, pt["location"], pt["text"]) if pool else []
                    for iid in linked:
                        conn.execute("INSERT OR IGNORE INTO pla_point_items(point_id, item_id) VALUES (?,?)", (pid, iid))
                    stats["linked"] += bool(linked)
                # 写真: A〜B 列で表の行にあるものはその箇所、ほかはチェック表の「場所の写真」
                for i, (row, col, ext, data) in enumerate(sorted(sh["photos"]), start=1):
                    name = _save_photo(data, ext)
                    pid = None
                    if col <= 2 and rows[0] - 1 <= row <= rows[-1] + 1:
                        idx = max([k for k, rr in enumerate(rows) if rr <= max(row, rows[0])] or [0])
                        pid = pids[idx]
                    conn.execute("""INSERT INTO pla_photos(sheet_id, point_id, file, sort_order, source, created_at)
                                    VALUES (?,?,?,?,?,?)""", (sid, pid, name, i, "xls", now))
                    stats["photos"] += 1
    conn.commit()
    if tmp:
        shutil.rmtree(tmp, ignore_errors=True)
    return stats


# ---------------------------------------------------------------- 点検記録の Excel（チェック表ごとに 1 シート。紙と同じ形）
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def export_month(conn, process, ym):
    first = dt.date.fromisoformat(ym + "-01")
    last = (first.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
    days = [first + dt.timedelta(days=i) for i in range((last - first).days + 1)]
    sheets = [dict(r) for r in conn.execute(
        "SELECT * FROM pla_sheets WHERE process = ? AND active = 1 ORDER BY sort_order", (process,))]
    res = {(r["point_id"], r["check_date"], r["timing"]): dict(r) for r in conn.execute(
        """SELECT c.* FROM pla_checks c JOIN pla_points p ON p.id = c.point_id JOIN pla_sheets s ON s.id = p.sheet_id
           WHERE s.process = ? AND c.check_date BETWEEN ? AND ?""", (process, first.isoformat(), last.isoformat()))}
    reviews = {(r["check_date"], r["timing"]): r["reviewer"] for r in conn.execute(
        "SELECT * FROM pla_reviews WHERE process = ? AND check_date BETWEEN ? AND ?", (process, first.isoformat(), last.isoformat()))}
    fill_head = PatternFill("solid", fgColor="B5001F")
    fill_ng = PatternFill("solid", fgColor="F8C9C9")
    fill_sun = PatternFill("solid", fgColor="FDECEA")
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    ng_rows = []
    for sh in sheets:
        name = re.sub(r"[\\/*?:\[\]]", "", sh["title"])[:28] or "チェック表"
        while name in used:
            name = name[:26] + "_" + str(len(used))
        used.add(name)
        ws = wb.create_sheet(name)
        points = [dict(r) for r in conn.execute(
            "SELECT * FROM pla_points WHERE sheet_id = ? AND active = 1 ORDER BY sort_order", (sh["id"],))]
        ws.append([f"プラスチック部品チェック表　{sh['title']}　{first.year}年{first.month}月　（{process}）"])
        ws["A1"].font = Font(bold=True, size=13)
        ws.append(["点検箇所", "項目"] + sum(([f"{d.day}", ""] for d in days), []))
        ws.append(["", ""] + ["始", "終"] * len(days))
        for i in range(len(days)):
            ws.merge_cells(start_row=2, start_column=3 + i * 2, end_row=2, end_column=4 + i * 2)
        for c in (1, 2):
            ws.merge_cells(start_row=2, start_column=c, end_row=3, end_column=c)
        for row in ws.iter_rows(min_row=2, max_row=3):
            for cell in row:
                cell.fill, cell.font, cell.border = fill_head, Font(bold=True, color="FFFFFF"), BOX
                cell.alignment = Alignment(horizontal="center", vertical="center")
        for pt in points:
            row = [pt["location"] or "", pt["item_text"]]
            for d in days:
                for t in TIMINGS:
                    r = res.get((pt["id"], d.isoformat(), t))
                    row.append({"ok": "○", "ng": "×", "na": "－"}.get(r["result"]) if r else "")
                    if r and r["result"] == "ng":
                        ng_rows.append([d.isoformat(), t, sh["title"], pt["location"] or "", pt["item_text"], r["note"], r["checked_by"]])
            ws.append(row)
            for i, cell in enumerate(ws[ws.max_row], start=1):
                cell.border = BOX
                cell.alignment = Alignment(horizontal="center", vertical="center") if i > 2 else Alignment(wrap_text=True, vertical="center")
                if cell.value == "×":
                    cell.fill, cell.font = fill_ng, Font(bold=True, color="C62828")
                elif i > 2 and days[(i - 3) // 2].weekday() == 6:
                    cell.fill = fill_sun
        for label in ("点検者", "確認者"):
            row = ["", label]
            for d in days:
                for t in TIMINGS:
                    if label == "確認者":
                        row.append(reviews.get((d.isoformat(), t), ""))
                    else:
                        row.append("、".join(sorted({r["checked_by"] for (pid, cd, tt), r in res.items()
                                                    if cd == d.isoformat() and tt == t and r["checked_by"]
                                                    and any(p["id"] == pid for p in points)})))
            ws.append(row)
            for i, cell in enumerate(ws[ws.max_row], start=1):
                cell.border = BOX
                cell.font = Font(bold=True) if i == 2 else Font(size=8)
                if i > 2:
                    cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True, text_rotation=255)
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 30
        for i in range(len(days) * 2):
            ws.column_dimensions[get_column_letter(3 + i)].width = 3.2
        ws.freeze_panes = "C4"
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToHeight = 0
    ws2 = wb.create_sheet("異常の内容", 0)
    ws2.append(["日付", "時", "チェック表", "点検箇所", "項目", "内容", "点検者"])
    for r in sorted(ng_rows):
        ws2.append(r)
    for cell in ws2[1]:
        cell.fill, cell.font = fill_head, Font(bold=True, color="FFFFFF")
    for i, w in enumerate((11, 7, 26, 16, 30, 40, 10), start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--replace", action="store_true", help="入れ直す（点検記録のあるチェック表は残して「使わない」にする）")
    ap.add_argument("--xlsx", help="変換済みの .xlsx があるフォルダ（省略時は Excel で変換する）")
    args = ap.parse_args()
    conn = connect()
    init_db(conn)
    s = import_folders(conn, replace=args.replace, xlsx_dir=args.xlsx)
    print(f"チェック表 {s['sheets']} 枚（使わない {s['inactive']}）・点検箇所 {s['points']}（プラリストと紐づけ {s['linked']}）・写真 {s['photos']} 枚")
