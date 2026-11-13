"""測定機器の校正（検査記録書）。

    python calibration.py        # 記録書の xls に載っている機器を測定機器に登録する（FORMS の seed。何度流しても同じ）

    様式（FORMS）。type で入力の表・計算・印刷が変わる
      - 庫内・室温計 … temp     「保管庫の庫内温度計、作業場室温計の検査記録書」（26.…【230627】.xls、11 台）
      - 庫内温度計   … temp     「製品保管庫の庫内温度計の検査記録書」（測定機器毎の管理台帳の「26」シートの 1 台）
      - タイマー     … timer    「CCPモニタリングに用いるタイマーの校正記録」（NE-023-001）
      - 蒸煮コンベア … conveyor 「蒸煮コンベア速度校正記録」（NE-023-002）
    temp は 1 行 = 測定機器 1 台（calibration_rows の列）。timer・conveyor は機器に結びつかない行で、中身は calibration_rows.data（JSON）。
    様式を増やすときは FORMS に足す。印刷は render_print（紙の記録書と同じ並び）。

    タイマー: 差異 = 実測時間② − 設定時間①（秒）。±2.00 秒以内なら合
    蒸煮コンベア: 差異 = (実測通過時間② ÷ 基準通過時間① − 1) × 100（%、小数第 2 位で四捨五入）。
      5 回とも ±1% 以内なら全体で合、1 回でも外れたら否（ちょうど 1.00% は「±1%以内」として合）

    器差の求め方（記録書の【器差の求め方と補正の仕方】のとおり）
      - 庫内温度計の最小表示が 0.1℃ … 基準温度計の補正温度（Ａ）をそのまま使う
      - 0.5℃ … Ａを「二捨三入」（小数第 1 位が 0〜2 → .0、3〜7 → .5、8〜9 → 繰り上げ）
      - 1.0℃ … Ａを「四捨五入」して整数に
      器差 = 庫内温度計の表示温度（Ｂ） − 処理後のＡ
      判定 = 日常点検が否、または |器差| ≧ 2.0℃ なら否。それ以外は合（器差分を補正して使用可能）
      補正温度 = −器差（庫内温度計の表示にこの値を足すと基準温度になる）
    マイナスの温度（冷凍区）は絶対値で丸めてから符号を戻す（−18.5 → 四捨五入で −19）。
"""
import argparse
import datetime as dt
import html
import os
import sys
from decimal import ROUND_FLOOR, ROUND_HALF_UP, Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
DAILY_ITEMS = ["庫内温度計に性能を左右するような損傷がないこと。", "表示に欠け等がなく、庫内温度が表示されていること。"]
FORMS = {
    "庫内・室温計": {
        "title": "保管庫の庫内温度計、作業場室温計の検査記録書",
        "code": "NF-102-器-026-05",
        "kind": "庫内・室温計",        # 検査する測定機器の種類（行になる）
        "reference_kind": "温度計",    # 使用基準器（検査済みデジタル温度計）の種類
        "limit": 2.0,                  # 器差 ±2.0℃未満なら合
        "zones": ["冷蔵区", "冷凍区", "室温"],  # 紙は冷蔵区・冷凍区だけ。作業場の室温計用に「室温」を足した
        "min_displays": [0.1, 0.5, 1.0],
        "daily_items": DAILY_ITEMS,
        "name_label": "保管庫・作業場名",
        "action_label": "庫内温度計・室温計 定期検査不合格時 処置記録の№",
        "approvals": [["approved_qc", "確認者"]],  # 印刷の右上の確認欄（列名, 見出し）
        "type": "temp", "has_inspector": True,
        "reference_label": "使用基準器（検査済みデジタル温度計の台帳№）", "note_label": "メモ",
        "kind_labels": {"plant": "使用エリア", "type": "保管庫・作業場名", "capacity": "使用範囲", "division": "最小表示"},
        # 記録書に載っている機器（測定機器毎の管理台帳には無い）。python calibration.py で登録する
        "seed": {"file": "26.保管庫の庫内温度計、作業場室温計の検査記録書【230627】.xls", "sheet": "装置の温度計", "no_col": 1, "name_col": 4},
    },
    "庫内温度計": {
        "title": "製品保管庫の庫内温度計の検査記録書",
        "code": "NF-102-器-026-05",
        "kind": "庫内温度計",          # 台帳の「26製品保管庫の庫内温度計」シート（instruments.py）
        "reference_kind": "温度計",
        "limit": 2.0,
        "zones": ["冷蔵区", "冷凍区"],
        "min_displays": [0.1, 0.5, 1.0],
        "daily_items": DAILY_ITEMS,
        "name_label": "保管庫名",
        "action_label": "製品保管庫の庫内温度計 定期検査不合格時の処置記録の№",
        "approvals": [["approved_manager", "工場長"], ["approved_qc", "品管"]],
        "type": "temp", "has_inspector": True,
        "reference_label": "使用基準器（検査済みデジタル温度計の台帳№）", "note_label": "メモ",
    },
    "タイマー": {
        "type": "timer",
        "title": "CCPモニタリングに用いるタイマーの校正記録",
        "code": "NE-023-001", "created": "H23/7/25",
        "limit": 2.0,                  # ±2.00 秒以内なら合
        "rows": 5,                     # 紙の行数（足りなければ画面で増やせる）
        "measure_note": "加熱装置毎、各1回",
        "reference_label": "ストップウォッチの管理番号", "reference_kind": "ストップウォッチ",
        "extra_fields": [["line_name", "ライン名"], ["process_name", "工程名"]],
        "note_label": "改善措置・対応等",
        "approvals": [["approved_manager", "工場長"], ["approved_qc", "品質管理課長"], ["approved_staff", "担当者"]],
    },
    "蒸煮コンベア": {
        "type": "conveyor",
        "title": "蒸煮コンベア速度校正記録",
        "code": "NE-023-002", "created": "H23/1/11",
        "limit": 1.0,                  # 5 回とも ±1% 以内なら合
        "rows": 5,
        "reference_label": "使用計測器", "reference_kind": "ストップウォッチ",
        "extra_fields": [["ccp_no", "CCP№"], ["inspected_time", "検査時刻"], ["product_code", "検査時製造品目 コード№"], ["product_name", "品名"]],
        "note_label": "改善措置・対応等",
        "approvals": [["approved_manager", "工場長"], ["approved_qc", "品管責任者"], ["approved_staff", "担当者"]],
    },
}
for _f in FORMS.values():
    _f.setdefault("extra_fields", [])
NUM_FIELDS = ["ref_display", "ref_corrected", "inst_display"]


def _d(v):
    return Decimal(str(v))


def round_by_min(a, step):
    """基準温度計の補正温度（Ａ）を庫内温度計の最小表示に合わせて処理する。"""
    if a is None or step is None:
        return None
    x, sign = abs(_d(a)), (-1 if a < 0 else 1)
    step = _d(step)
    if step == Decimal("0.1"):
        r = x.quantize(Decimal("0.1"), ROUND_HALF_UP)
    elif step == Decimal("0.5"):  # 二捨三入
        t = x.quantize(Decimal("0.1"), ROUND_HALF_UP)
        n = t.to_integral_value(ROUND_FLOOR)
        tenth = int((t - n) * 10)
        r = n if tenth <= 2 else n + Decimal("0.5") if tenth <= 7 else n + 1
    elif step == Decimal("1"):  # 四捨五入
        r = x.quantize(Decimal("1"), ROUND_HALF_UP)
    else:
        return None
    return float(r * sign) or 0.0  # -0.0 を 0.0 に


def compute(row, limit=2.0):
    """1 行分の 処理後の温度・器差・判定・補正温度 を計算して row に入れる。"""
    rounded = round_by_min(row.get("ref_corrected"), row.get("min_display"))
    b = row.get("inst_display")
    error = float(_d(b) - _d(rounded)) if rounded is not None and b is not None else None
    if row.get("daily_ok") == 0:
        result = "否"
    elif error is None:
        result = None
    else:
        result = "否" if abs(_d(error)) >= _d(limit) else "合"
    row.update(rounded=rounded, error=error, result=result,
               correction=(-error or 0.0) if error is not None else None)
    return row


# ---------------------------------------------------------------- タイマー・蒸煮コンベア
ROW_FIELDS = {  # 入力する項目（数字の項目, 文字の項目）
    "timer": (["set_min", "set_sec", "act_min", "act_sec"], ["device", "remark"]),
    "conveyor": (["freq", "ref_min", "ref_sec", "act_min", "act_sec"], ["remark"]),
}
NUM_LABELS = {"set_min": "設定時間（分）", "set_sec": "設定時間（秒）", "act_min": "実測時間（分）", "act_sec": "実測時間（秒）",
              "freq": "設定周波数", "ref_min": "基準通過時間（分）", "ref_sec": "基準通過時間（秒）"}


def to_num(v, label, where):
    """入力 → 数字（空欄は None）。全角・「−」・単位も受け付ける。読めなければ ValueError。"""
    if v in ("", None):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    from db import norm
    t = norm(str(v)).replace("−", "-").replace("ー", "-")
    for unit in ("℃", "秒", "分", "Hz", "%", "″", "′"):
        t = t.replace(unit, "")
    try:
        return float(t.strip())
    except ValueError:
        raise ValueError(f"{where}: {label}「{v}」が数字ではありません") from None


def seconds(m, s):
    """分・秒 → 秒（どちらも空欄なら None）。"""
    if m is None and s is None:
        return None
    return float(_d(m or 0) * 60 + _d(s or 0))


def _round2(v):
    return float(_d(v).quantize(Decimal("0.01"), ROUND_HALF_UP))


def compute_free(form, row):
    """タイマー・蒸煮コンベアの 1 行（row["data"]）の計算。差異と判定を入れる。"""
    d = row["data"]
    limit = _d(form["limit"])
    if form["type"] == "timer":
        a, b = seconds(d.get("set_min"), d.get("set_sec")), seconds(d.get("act_min"), d.get("act_sec"))
        diff = _round2(_d(b) - _d(a)) if a is not None and b is not None else None
        d.update(set_total=a, act_total=b)
    else:
        a, b = seconds(d.get("ref_min"), d.get("ref_sec")), seconds(d.get("act_min"), d.get("act_sec"))
        diff = _round2((_d(b) / _d(a) - 1) * 100) if a and b is not None else None
        d.update(ref_total=a, act_total=b)
    d["diff"] = diff
    row["result"] = None if diff is None else "合" if abs(_d(diff)) <= limit else "否"
    return row


def build_free_rows(form, raw_rows):
    """画面から来た行 → 保存する行（空の行は捨てる）。読めない数字は ValueError。"""
    nums, texts = ROW_FIELDS[form["type"]]
    out = []
    for n, r in enumerate(raw_rows or [], 1):
        where = f"{n}行目"
        d = {k: to_num(r.get(k), NUM_LABELS[k], where) for k in nums}
        from db import norm
        d |= {k: (norm(str(r.get(k) or "")).strip() or None) for k in texts}
        if all(v is None for v in d.values()):
            continue
        if form["type"] == "conveyor" and seconds(d.get("ref_min"), d.get("ref_sec")) == 0:
            raise ValueError(f"{where}: 基準通過時間が 0 秒です")
        out.append(compute_free(form, {"sort_order": n, "data": d}))
    return out


def overall_result(form, rows):
    """記録全体の判定。蒸煮コンベアは 5 回とも合で合（1 回でも否なら否、回数が足りなければ未判定）。"""
    results = [r.get("result") for r in rows]
    if "否" in results:
        return "否"
    if form.get("type") == "conveyor":
        return "合" if results.count("合") >= form["rows"] else None
    return "合" if "合" in results else None


def fmt(v, sign=False):
    """12.0 → '12.0'、器差は '+0.5' のように符号を付ける。"""
    if v is None:
        return ""
    s = f"{v:+.1f}" if sign else f"{v:.1f}"
    return "±0.0" if sign and v == 0 else s


def event_note(row):
    """測定機器の履歴に残す 1 行（例: 冷凍区 基準 -20.0℃／表示 -19.0℃ 器差 +1.0℃ 合（補正 -1.0℃））"""
    parts = []
    if row.get("zone"):
        parts.append(row["zone"])
    if row.get("rounded") is not None and row.get("inst_display") is not None:
        parts.append(f"基準 {fmt(row['rounded'])}℃／表示 {fmt(row['inst_display'])}℃ 器差 {fmt(row['error'], True)}℃")
    if row.get("daily_ok") == 0:
        parts.append("日常点検 否")
    if row.get("result"):
        parts.append(row["result"] + (f"（補正 {fmt(row['correction'], True)}℃）" if row["result"] == "合" and row.get("correction") else ""))
    if row.get("action_no"):
        parts.append(f"処置記録№{row['action_no']}")
    return " ".join(parts) or None


# ---------------------------------------------------------------- 記録書に載っている機器の登録
def read_seed(form):
    """記録書の xls から (管理番号, 保管庫・作業場名) を読む。"""
    import xlrd
    from db import norm
    cfg = form["seed"]
    sh = xlrd.open_workbook(os.path.join(HERE, cfg["file"])).sheet_by_name(cfg["sheet"])
    out = []
    for r in range(sh.nrows):
        no, name = sh.cell_value(r, cfg["no_col"]), sh.cell_value(r, cfg["name_col"])
        if isinstance(no, float) and no.is_integer() and str(name).strip():
            out.append((str(int(no)), norm(str(name)).strip()))
    return out


def seed(conn):
    """FORMS の seed にある機器を測定機器に登録する。既にある機器（種類 + №）は空欄だけ埋める（アプリで直した値は残す）。"""
    from instruments import area_by_place
    now = dt.datetime.now().isoformat(timespec="seconds")
    result = {}
    for name, form in FORMS.items():
        if "seed" not in form:
            continue
        added = 0
        for n, (no, place) in enumerate(read_seed(form), 1):
            row = conn.execute("SELECT id FROM instruments WHERE kind=? AND no=?", (form["kind"], no)).fetchone()
            if row:
                conn.execute("UPDATE instruments SET type=COALESCE(type, ?), sort_order=COALESCE(sort_order, ?) WHERE id=?", (place, n, row[0]))
                continue
            conn.execute(
                "INSERT INTO instruments(kind, no, type, area, status, sort_order, source, updated_at) VALUES (?,?,?,?, '使用中', ?,?,?)",
                (form["kind"], no, place, area_by_place(place), n, form["seed"]["file"], now))
            added += 1
        result[name] = added
    conn.commit()
    return result


# ---------------------------------------------------------------- 印刷（紙の記録書と同じ並び、A4 横）
PRINT_CSS = """
@page { size: A4 landscape; margin: 7mm 9mm; }
* { box-sizing: border-box; }
body { margin: 0; background: #fff; color: #000; font: 9.5pt/1.3 "Yu Gothic", "Meiryo", sans-serif; }
.sheet { width: 279mm; margin: 0 auto; padding: 4mm 0; }
.bar { display: flex; gap: 8px; justify-content: center; padding: 10px; background: #eef1f4; border-bottom: 1px solid #ccd; }
.bar button { font: inherit; font-size: 15px; padding: 6px 18px; cursor: pointer; }
.top { display: flex; justify-content: space-between; align-items: flex-start; }
.code { font-size: 9pt; }
h1 { font-size: 15pt; margin: 1mm 0 0; letter-spacing: .08em; text-align: center; flex: 1; }
.stamps { display: flex; border: 1px solid #000; }
.stamps div { width: 20mm; text-align: center; border-left: 1px solid #000; }
.stamps div:first-child { border-left: 0; }
.stamps b { display: block; font-weight: normal; font-size: 8.5pt; border-bottom: 1px solid #000; }
.stamps span { display: flex; align-items: center; justify-content: center; height: 12mm; font-size: 9pt; padding: 0 2px; }
.head { display: grid; grid-template-columns: 1fr 1fr; gap: 1mm 10mm; margin: 2mm 0 2mm; font-size: 10pt; }
.head .u { border-bottom: 1px solid #000; display: inline-block; min-width: 70mm; padding: 0 2mm; }
.rules { display: grid; grid-template-columns: 1fr 1.3fr; gap: 6mm; font-size: 8pt; line-height: 1.25; margin-bottom: 2mm; }
.rules table { border-collapse: collapse; width: 100%; }
.rules td, .rules th { border: 1px solid #000; padding: .5mm 1.5mm; text-align: left; font-weight: normal; }
.rules p { margin: 0 0 .5mm; }
table.main { border-collapse: collapse; width: 100%; font-size: 9pt; }
table.main th, table.main td { border: 1px solid #000; padding: .4mm 1.2mm; text-align: center; vertical-align: middle; }
table.main th { font-weight: normal; font-size: 8.5pt; line-height: 1.25; background: #f2f2f2; }
table.main td.l { text-align: left; }
table.main td.n { text-align: right; white-space: nowrap; }
table.main tbody tr { height: 8mm; page-break-inside: avoid; }
.opt { display: inline-block; min-width: 5.5mm; padding: 0 .5mm; }
.opt.on { border: 1.4px solid #000; border-radius: 50%; font-weight: bold; }
.zone { text-align: left; white-space: nowrap; font-size: 7pt; line-height: 1.05; }
.foot { display: flex; justify-content: space-between; gap: 6mm; font-size: 8.5pt; margin-top: 1.5mm; }
.memo { font-size: 9.5pt; }
@media print { .bar { display: none; } .sheet { padding: 0; } th { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
"""


def render_print(form, cal, rows):
    """印刷用の HTML（様式の type で振り分け）。cal が空 ({}) なら白紙の記録書。"""
    if form.get("type") == "timer":
        return render_print_timer(form, cal, rows)
    if form.get("type") == "conveyor":
        return render_print_conveyor(form, cal, rows)
    return render_print_temp(form, cal, rows)


def render_print_temp(form, cal, rows):
    """庫内温度計・室温計（A4 横）。rows は機器ごと（記録がなければ空欄で出す）。"""
    e = lambda v: html.escape("" if v is None else str(v))
    blank = not cal.get("id")
    d = cal.get("inspected_on")
    date = f"{int(d[:4])}年 {int(d[5:7])}月 {int(d[8:10])}日" if d else "　　　　年　　　月　　　日"

    def choice(options, cur):
        return "・".join(f'<span class="opt{" on" if o == cur else ""}">{o}</span>' for o in options)

    def num(v, sign=False):
        return f"{e(fmt(v, sign))}℃" if v is not None else "℃"

    body = []
    for n, r in enumerate(rows, 1):
        daily = {1: "合", 0: "否"}.get(r.get("daily_ok"))
        zones = "<br>".join(f'{"■" if r.get("zone") == z else "□"}{z}' for z in form["zones"])
        md = r.get("min_display")
        body.append(f"""<tr>
          <td>{n}</td><td>{e(r.get("no"))}</td><td class="l">{e(r.get("plant"))}</td><td class="l">{e(r.get("type"))}</td>
          <td>{e(f"{md:.1f}℃") if md is not None and not blank and r.get("recorded") else ""}</td>
          <td>{choice(["合", "否"], daily)}</td><td class="zone">{zones}</td>
          <td class="n">{num(r.get("ref_display"))}</td><td class="n">{num(r.get("ref_corrected"))}</td><td class="n">{num(r.get("rounded"))}</td>
          <td class="n">{num(r.get("inst_display"))}</td><td class="n">{num(r.get("error"), True)}</td>
          <td>{choice(["合", "否"], r.get("result"))}</td>
          <td class="n">{num(r.get("correction"), True)}</td><td>{e(r.get("action_no"))}</td></tr>""")
    stamps = "".join(f"<div><b>{e(label)}</b><span>{e(cal.get(key))}</span></div>" for key, label in form["approvals"])
    daily_items = "".join(f"<p>({i}) {e(t)}</p>" for i, t in enumerate(form["daily_items"], 1))
    lim = f"{form['limit']:.1f}"
    printed = dt.datetime.now().strftime("%Y/%m/%d %H:%M")
    return f"""<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(form["title"])}{" " + e(d) if d else ""}</title><style>{PRINT_CSS}</style></head><body>
<div class="bar"><button onclick="window.print()">印刷する</button><button onclick="window.close()">閉じる</button></div>
<div class="sheet">
  <div class="top"><div class="code">様式：{e(form["code"])}</div><h1>{e(form["title"])}</h1><div class="stamps">{stamps}</div></div>
  <div class="head">
    <div>検査年月日：<span class="u">{date}</span></div>
    <div>使用基準器：<span class="u">検査済みデジタル温度計（台帳№ {e(cal.get("reference")) or "　　　"}）</span></div>
    <div>検査者：<span class="u">{e(cal.get("inspector"))}</span></div><div></div>
  </div>
  <div class="rules">
    <div><p><b>日常点検項目</b></p>{daily_items}
      <p style="margin-top:1.5mm"><b>合否判定基準</b></p>
      <table><tr><td>器差：±{lim}℃未満</td><td>器差分を補正して使用可能</td></tr><tr><td>器差：±{lim}℃以上</td><td>不合格とし、処置する</td></tr></table></div>
    <div><p><b>【器差の求め方と補正の仕方】</b></p>
      <table><tr><th style="width:30mm">庫内温度計の最小表示</th><th>方法</th></tr>
        <tr><td>0.1℃の場合</td><td>基準温度計の補正温度（Ａ）と庫内温度計の表示温度（Ｂ）とそのまま比較する。器差分を補正する。</td></tr>
        <tr><td>0.5℃の場合</td><td>基準温度計の補正温度（Ａ）を「二捨三入」し、庫内温度計の表示温度（Ｂ）と比較する。器差分を補正する。</td></tr>
        <tr><td>1.0℃の場合</td><td>基準温度計の補正温度（Ａ）を「四捨五入」し、庫内温度計の表示温度（Ｂ）と比較する。器差分を補正する。</td></tr></table>
      <p style="margin-top:.8mm">器差 ＝ 表示温度（Ｂ）－ 処理後の温度　／　補正温度 ＝ －器差（庫内温度計の表示に足すと基準温度）</p></div>
  </div>
  <table class="main"><thead>
    <tr><th rowspan="2">№</th><th colspan="4">検査温度計</th><th rowspan="2">日常点検<br>項目</th><th rowspan="2">検査温度</th>
      <th rowspan="2">基準温度計<br>表示(℃)</th><th rowspan="2">基準温度計<br>補正温度(℃)<br>（Ａ）</th><th rowspan="2">四捨五入等の<br>処理後の<br>温度(℃)</th>
      <th rowspan="2">庫内温度計<br>表示温度(℃)<br>（Ｂ）</th><th rowspan="2">器差(℃)</th><th rowspan="2">判定</th><th rowspan="2">補正温度</th>
      <th rowspan="2" style="width:30mm">{e(form["action_label"])}</th></tr>
    <tr><th>管理番号</th><th>使用ｴﾘｱ</th><th>{e(form["name_label"])}</th><th>最小表示</th></tr>
  </thead><tbody>{"".join(body)}</tbody></table>
  <div class="foot"><span class="memo">{f"メモ：{e(cal.get('note'))}" if cal.get("note") else ""}</span><span>{"" if blank else "設備保全アプリで記録　"}印刷 {printed}</span></div>
</div></body></html>"""


# ---- タイマー・蒸煮コンベア（A4 縦）
PORTRAIT_CSS = PRINT_CSS.replace("size: A4 landscape; margin: 7mm 9mm;", "size: A4 portrait; margin: 14mm 16mm;") + """
.sheet { width: 178mm; }
.ptop { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4mm; }
.pt h1 { text-align: center; font-size: 15pt; margin: 6mm 0 8mm; font-weight: normal; letter-spacing: .05em; }
.stamps div { width: 22mm; }
.stamps span { height: 14mm; }
.lines { display: grid; gap: 4mm; font-size: 11pt; }
.lines .u { border-bottom: 1px solid #000; display: inline-block; min-width: 72mm; padding: 0 2mm; }
.crit { margin: 7mm 0 4mm; font-size: 11pt; line-height: 1.6; }
table.pm { border-collapse: collapse; width: 100%; font-size: 10.5pt; }
table.pm th, table.pm td { border: 1px solid #000; padding: 1mm 1.5mm; text-align: center; vertical-align: middle; }
table.pm th { font-weight: normal; background: #f2f2f2; font-size: 9.5pt; line-height: 1.3; }
table.pm tbody tr { height: 13mm; }
table.pm td.l { text-align: left; }
table.pm td.n { text-align: right; white-space: nowrap; }
table.pm .thick { border: 2.2px solid #000; }
.verdict { text-align: center; font-size: 14pt; margin: 7mm 0; }
.action { border: 1px solid #000; min-height: 70mm; margin-top: 8mm; padding: 2mm 3mm; font-size: 10.5pt; white-space: pre-wrap; }
.action b { display: block; font-weight: normal; margin-bottom: 2mm; }
.made { text-align: right; font-size: 9pt; margin-top: 4mm; }
"""


def _num_text(v, dec=None):
    """12.0 → '12'、12.5 → '12.5'（dec を指定すればその桁）。"""
    if v is None:
        return ""
    if dec is not None:
        return f"{v:.{dec}f}"
    return f"{v:f}".rstrip("0").rstrip(".")


def _print_common(form, cal):
    e = lambda v: html.escape("" if v is None else str(v))
    d = cal.get("inspected_on")
    date = f"{int(d[:4])}年 {int(d[5:7])}月 {int(d[8:10])}日" if d else "　　　　年　　　月　　　日"
    stamps = "".join(f"<div><b>{e(label)}</b><span>{e(cal.get(key))}</span></div>" for key, label in form["approvals"])
    extra = cal.get("extra") or {}
    return e, date, stamps, extra


def _page(form, cal, inner):
    e = lambda v: html.escape("" if v is None else str(v))
    printed = dt.datetime.now().strftime("%Y/%m/%d %H:%M")
    d = cal.get("inspected_on")
    return f"""<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(form["title"])}{" " + e(d) if d else ""}</title><style>{PORTRAIT_CSS}</style></head><body>
<div class="bar"><button onclick="window.print()">印刷する</button><button onclick="window.close()">閉じる</button></div>
<div class="sheet pt">{inner}
  <div class="foot"><span>{"" if not cal.get("id") else "設備保全アプリで記録"}</span><span>様式作成日：{e(form.get("created"))}　印刷 {printed}</span></div>
</div></body></html>"""


def _signed(v):
    """差異の表示（+1.35 / -2.10 / 0.00）。"""
    return "" if v is None else "0.00" if v == 0 else f"{v:+.2f}"


def _min_sec(d, m, s):
    """分・秒（片方だけ入っていればもう片方は 0）。"""
    if d.get(m) is None and d.get(s) is None:
        return "", ""
    return _num_text(d.get(m) or 0.0), _num_text(d.get(s) or 0.0)


def _choice(options, cur):
    return "・".join(f'<span class="opt{" on" if o == cur else ""}">{o}</span>' for o in options)


def _pad(rows, n):
    rows = [r for r in rows]
    return rows + [{"data": {}} for _ in range(max(0, n - len(rows)))]


def render_print_timer(form, cal, rows):
    e, date, stamps, extra = _print_common(form, cal)
    body = []
    for r in _pad(rows, form["rows"]):
        d = r.get("data") or {}
        t = lambda m, s: "{} 分　{} 秒".format(*map(e, _min_sec(d, m, s)))
        diff = d.get("diff")
        body.append(f"""<tr><td class="l">{e(d.get("device"))}</td><td>{t("set_min", "set_sec")}</td><td>{t("act_min", "act_sec")}</td>
          <td class="n">{e(_signed(diff))} 秒</td><td>{_choice(["合", "否"], r.get("result"))}</td><td class="l">{e(d.get("remark"))}</td></tr>""")
    inner = f"""
  <div class="code">様式番号：{e(form["code"])}</div>
  <h1>{e(form["title"])}</h1>
  <div class="ptop">
    <div class="lines">
      <div>検査年月日：<span class="u">{date}</span></div>
      <div>ライン名：<span class="u">{e(extra.get("line_name"))}</span></div>
      <div>工程名：<span class="u">{e(extra.get("process_name"))}</span></div>
      <div>ストップウォッチの管理番号：<span class="u" style="min-width:50mm">{e(cal.get("reference"))}</span></div>
    </div>
    <div class="stamps">{stamps}</div>
  </div>
  <div class="crit">【測定回数】 {e(form.get("measure_note"))}<br>【合否判定基準】 ±{form["limit"]:.2f}秒以内</div>
  <table class="pm"><thead><tr><th style="width:24mm">加熱装置の<br>名前又は番号</th><th>タイマー設定時間<br>（①）</th><th>実測時間<br>（②）</th>
    <th style="width:24mm">差異<br>（②－①）</th><th style="width:20mm">判定</th><th style="width:34mm">備考</th></tr></thead>
    <tbody>{"".join(body)}</tbody></table>
  <div class="action"><b>{e(form["note_label"])}</b>{e(cal.get("note"))}</div>"""
    return _page(form, cal, inner)


def render_print_conveyor(form, cal, rows):
    e, date, stamps, extra = _print_common(form, cal)
    body = []
    for n, r in enumerate(_pad(rows, form["rows"]), 1):
        d = r.get("data") or {}
        ms = lambda m, s: "{}′ {}″".format(*map(e, _min_sec(d, m, s)))
        diff = d.get("diff")
        body.append(f"""<tr><td>{n}</td><td>{e(_num_text(d.get("freq")))}</td><td>{ms("ref_min", "ref_sec")}</td><td class="n">{e(_num_text(d.get("ref_total")))}</td>
          <td>{ms("act_min", "act_sec")}</td><td class="n">{e(_num_text(d.get("act_total")))}</td>
          <td class="n thick">{e(_signed(diff))}</td><td class="l">{e(d.get("remark"))}</td></tr>""")
    ccp = e(extra.get("ccp_no")) or "＿"
    time = e(extra.get("inspected_time"))
    inner = f"""
  <div class="code">様式番号：{e(form["code"])}</div>
  <h1>CCP№{ccp} {e(form["title"])}</h1>
  <div class="ptop"><div></div><div class="stamps">{stamps}</div></div>
  <div class="lines">
    <div>検査年月日：<span class="u">{date}{"　" + time if time else ""}</span></div>
    <div>検査時製造品目：<span class="u">コード№ {e(extra.get("product_code"))}　　品名 {e(extra.get("product_name"))}</span></div>
    <div>使用計測器：<span class="u">{e(cal.get("reference"))}</span></div>
  </div>
  <div class="crit" style="font-size:10pt">判定については、{form["rows"]}回の測定を実施し、基準通過時間との差異（太線囲み）が{form["rows"]}回とも±{_num_text(form["limit"])}％以内であれば判定合格、1回でも{_num_text(form["limit"])}％以上の差異が発生した場合は判定否とし、改善措置をとることとする。</div>
  <table class="pm"><thead>
    <tr><th rowspan="2" style="width:9mm">№</th><th>設定周波数</th><th colspan="2">基準通過時間</th><th colspan="2">実測通過時間</th><th class="thick">差異<br>(②/①-1)×100</th><th rowspan="2" style="width:26mm">備考</th></tr>
    <tr><th>(Hz)</th><th>(分′秒″)</th><th>(秒) ①</th><th>(分′秒″)</th><th>(秒) ②</th><th class="thick">(％)</th></tr>
  </thead><tbody>{"".join(body)}</tbody></table>
  <div class="verdict">判定：　{_choice(["合", "否"], cal.get("result"))}</div>
  <div class="action" style="min-height:50mm"><b>{e(form["note_label"])}</b>{e(cal.get("note"))}</div>"""
    return _page(form, cal, inner)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    from db import connect, init_db
    conn = connect()
    init_db(conn)
    for name, added in seed(conn).items():
        n = conn.execute("SELECT COUNT(*) FROM instruments WHERE kind=?", (FORMS[name]["kind"],)).fetchone()[0]
        print(f"{name}: 追加 {added} 台（登録済み {n} 台）")


if __name__ == "__main__":
    main()
