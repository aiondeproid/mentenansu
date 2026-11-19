"""交換アラートをメールで通知する。タスクスケジューラから平日の朝に実行する想定。

    python notify.py              新しく期限に入った項目を送る（月曜は期限超過のままの項目もまとめて送る）
    python notify.py --dry-run    送らずに、送る内容を表示するだけ
    python notify.py --init       今あるアラートを「通知済み」にする（初回の一斉送信を防ぐ）
    python notify.py --test 宛先  テストメールを 1 通送る

    python notify.py --check      Notes で送れるか確認する（送信はしない。method = notes のとき）

送信設定は notify.ini（notify.ini.example をコピーして作る）。送り方は 2 通り:
  method = smtp   SMTP サーバーへ直接送る（Domino の SMTP なら通常ポート 25）
  method = notes  このパソコンの Notes クライアントから送る（SMTP の設定は不要）。
                  Notes は 32 ビット版のことが多いので、送信は Notes と同じビット数の
                  PowerShell に任せる（notes_send.ps1）。Python のビット数は問わない。
"""
import argparse
import configparser
import datetime as dt
import json
import os
import smtplib
import struct
import subprocess
import sys
from email.message import EmailMessage
from email.utils import formatdate, make_msgid

from db import connect, init_db

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.environ.get("MAINT_NOTIFY_INI", os.path.join(BASE_DIR, "notify.ini"))
NOTES_SCRIPT = os.path.join(BASE_DIR, "notes_send.ps1")
NOTES_TIMEOUT = 120  # 秒。Notes がパスワード入力待ちで止まったときに諦めるまで
WEEKDAYS = "月火水木金土日"
LEVEL_LABEL = {"over": "期限超過", "soon": "まもなく"}


def load_config():
    cp = configparser.ConfigParser()
    cp.read(CONFIG_PATH, encoding="utf-8")
    mail = cp["mail"] if cp.has_section("mail") else {}
    smtp = cp["smtp"] if cp.has_section("smtp") else {}
    notes = cp["notes"] if cp.has_section("notes") else {}
    app = cp["app"] if cp.has_section("app") else {}
    method = (mail.get("method", "smtp") or "smtp").strip().lower()
    return {
        "method": method if method in ("smtp", "notes") else "smtp",
        "notes_password": notes.get("password", ""),
        "host": smtp.get("host", "").strip(),
        "port": int(smtp.get("port", "25") or 25),
        "starttls": str(smtp.get("starttls", "no")).lower() in ("1", "yes", "true", "on"),
        "user": smtp.get("user", "").strip(),
        "password": smtp.get("password", ""),
        "from_addr": smtp.get("from", "").strip(),
        "app_url": (app.get("url", "") or "http://localhost:5000/").rstrip("/") + "/",
        "reminder_weekday": int(app.get("reminder_weekday", "0") or 0),  # 0=月曜
    }


def config_status(cfg):
    """画面表示用（パスワードは返さない）。"""
    if cfg["method"] == "notes":
        notes_ok, notes_msg = notes_powershell()
        configured = notes_ok is not None
    else:
        notes_msg = ""
        configured = bool(cfg["host"] and cfg["from_addr"])
    return {
        "method": cfg["method"], "notes_info": notes_msg,
        "configured": configured,
        "host": cfg["host"], "port": cfg["port"], "starttls": cfg["starttls"],
        "auth": bool(cfg["user"]), "from_addr": cfg["from_addr"], "app_url": cfg["app_url"],
        "reminder_weekday": WEEKDAYS[cfg["reminder_weekday"] % 7],
        "config_path": CONFIG_PATH, "config_exists": os.path.exists(CONFIG_PATH),
    }


# ---------------------------------------------------------------- 通知する項目
def collect_items(conn):
    from app import load_rules, part_state  # app 側の判定をそのまま使う

    items = []
    for r in load_rules(conn):
        if r["level"] not in LEVEL_LABEL:
            continue
        name = r["equipment_nickname"] or r["equipment_name"]
        parts = []
        if r["due_date"]:
            when = f"{-r['days_left']}日超過" if r["days_left"] < 0 else f"あと{r['days_left']}日"
            parts.append(f"期限 {r['due_date'].replace('-', '/')}（{when}）")
        if r["hours_left"] is not None:
            parts.append(f"稼働 {r['hours_used']}/{r['interval_hours']}h")
        items.append({
            "key": f"rule:{r['id']}:{r['last_done_date'] or '-'}:{r['level']}",
            "area": r["area"] or "未分類",
            "level": r["level"],
            "label": LEVEL_LABEL[r["level"]],
            "title": f"{name}（{r['equipment_code']}）" if r["equipment_code"] else name,
            "detail": " / ".join([r["title"]] + parts),
            "sort": r["days_left"] if r["days_left"] is not None else 0,
        })
    for p in conn.execute(
        """SELECT p.*, e.area, COALESCE(e.nickname, e.name) AS equipment_name FROM parts p
           LEFT JOIN equipment e ON e.id = p.equipment_id"""):
        lv = part_state(p)
        if lv not in LEVEL_LABEL:
            continue
        items.append({
            "key": f"part:{p['id']}:{p['grind_count']}:{lv}",
            "area": p["area"] or "未分類",
            "level": lv,
            "label": "研磨上限" if lv == "over" else "研磨上限が近い",
            "title": f"{p['part_type']} {p['serial']}" + (f"（{p['equipment_name']}）" if p["equipment_name"] else ""),
            "detail": f"研磨 {p['grind_count']}/{p['grind_limit'] or 10}回",
            "sort": 0,
        })
    return items


def recipients_by_area(conn):
    result = {}
    for r in conn.execute("SELECT area, email FROM notify_recipients ORDER BY id"):
        result.setdefault(r["area"], []).append(r["email"])
    return result


def build_messages(conn, cfg, today=None, include_all=False):
    """送るメールを組み立てる。戻り値: (メールのリスト, 警告のリスト)"""
    today = today or dt.date.today()
    items = collect_items(conn)
    sent = {r["key"] for r in conn.execute("SELECT key FROM notify_log")}
    reminder_day = today.weekday() == cfg["reminder_weekday"]
    recips = recipients_by_area(conn)
    common = recips.get("*", [])

    by_area = {}
    for it in items:
        bucket = by_area.setdefault(it["area"], {"new": [], "stale": []})
        if include_all or it["key"] not in sent:
            bucket["new"].append(it)
        elif it["level"] == "over" and reminder_day:
            bucket["stale"].append(it)

    messages, warnings = [], []
    for area, b in by_area.items():
        if not b["new"] and not b["stale"]:
            continue
        to = sorted(set(recips.get(area, []) + common))
        if not to:
            warnings.append(f"{area}: 宛先が設定されていないため送りません（{len(b['new']) + len(b['stale'])}件）")
            continue
        for k in ("new", "stale"):
            b[k].sort(key=lambda i: (i["level"] != "over", i["sort"]))
        subject = f"【設備保全】{area} 交換アラート"
        counts = []
        if b["new"]:
            counts.append(f"新規{len(b['new'])}件")
        if b["stale"]:
            counts.append(f"期限超過のまま{len(b['stale'])}件")
        subject += " " + " / ".join(counts)
        messages.append({
            "area": area, "to": to, "subject": subject,
            "body": compose_body(area, b["new"], b["stale"], today, cfg),
            "keys": [i["key"] for i in b["new"]],
            "count": len(b["new"]) + len(b["stale"]),
        })
    return messages, warnings


def compose_body(area, new, stale, today, cfg):
    line = lambda i: f"［{i['label']}］{i['title']}\n    {i['detail']}"
    out = [f"設備保全アプリからのお知らせ（{today:%Y/%m/%d}）", f"エリア: {area}", ""]
    if new:
        out += [f"■ 新しく期限が近づいた・過ぎた項目（{len(new)}件）"] + [line(i) for i in new] + [""]
    if stale:
        out += [f"■ 期限を過ぎたままの項目（毎週{WEEKDAYS[cfg['reminder_weekday'] % 7]}曜にお知らせ）（{len(stale)}件）"]
        out += [line(i) for i in stale] + [""]
    out += [
        "対応したら、アプリの「交換アラート」から「実施を記録」してください。",
        f"{cfg['app_url']}#/plans",
        "",
        "※ このメールは設備保全アプリから自動送信しています。",
    ]
    return "\n".join(out)


def test_message(cfg, today=None):
    """テストメールの件名と本文。本番と同じ組み立てで、見本の項目を載せる。"""
    today = today or dt.date.today()
    d = lambda n: (today + dt.timedelta(days=n)).strftime("%Y/%m/%d")
    new = [
        {"level": "over", "label": LEVEL_LABEL["over"], "title": "真空包装機 2号（P-012）",
         "detail": f"シール用ヒーター交換 / 期限 {d(-8)}（8日超過）"},
        {"level": "soon", "label": LEVEL_LABEL["soon"], "title": "コンベア A（C-003）",
         "detail": f"ベルト交換 / 期限 {d(7)}（あと7日） / 稼働 1650/2000h"},
        {"level": "soon", "label": "研磨上限が近い", "title": "スライサー刃 SN-0421（スライサー 1号）",
         "detail": "研磨 9/10回"},
    ]
    stale = [{"level": "over", "label": LEVEL_LABEL["over"], "title": "排水ポンプ（W-001）",
              "detail": f"メカニカルシール交換 / 期限 {d(-18)}（18日超過）"}]
    head = [
        "これは設備保全アプリからのテストメールです。このメールが届いていれば送信設定は正しく動いています。",
        "以下は実際の通知メールの見本です（設備名・日付は架空のものです）。",
        "-" * 40,
        "",
    ]
    subject = "【設備保全】テストメール（見本: 前処理 交換アラート 新規3件 / 期限超過のまま1件）"
    return subject, "\n".join(head) + "\n" + compose_body("前処理（見本）", new, stale, today, cfg)


# ---------------------------------------------------------------- 送信
def send_mail(cfg, to, subject, body):
    if cfg["method"] == "notes":
        _run_notes({"mode": "send", "to": list(to), "subject": subject, "body": body}, cfg)
    else:
        _send_smtp(cfg, to, subject, body)


def _send_smtp(cfg, to, subject, body):
    if not cfg["host"] or not cfg["from_addr"]:
        raise RuntimeError(f"送信設定がありません。{CONFIG_PATH} の [smtp] host と from を設定してください")
    msg = EmailMessage()
    msg["From"] = cfg["from_addr"]
    msg["To"] = ", ".join(to)
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()
    msg.set_content(body, charset="utf-8")
    with smtplib.SMTP(cfg["host"], cfg["port"], timeout=30) as s:
        if cfg["starttls"]:
            s.starttls()
        if cfg["user"]:
            s.login(cfg["user"], cfg["password"])
        s.send_message(msg)


# ---------------------------------------------------------------- Notes から送る
def _notes_clsid_bits():
    """Lotus.NotesSession の COM が何ビットで登録されているか（32 / 64）。未登録なら None。"""
    try:
        import winreg
    except ImportError:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Lotus.NotesSession\CLSID") as k:
            clsid = winreg.QueryValue(k, None)
    except OSError:
        return None
    for bits, view in ((64, winreg.KEY_WOW64_64KEY), (32, winreg.KEY_WOW64_32KEY)):
        try:
            winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\CLSID\{clsid}\InprocServer32",
                           0, winreg.KEY_READ | view).Close()
            return bits
        except OSError:
            continue
    return None


def notes_powershell():
    """Notes と同じビット数の powershell.exe。戻り値: (パス または None, 画面に出す説明)"""
    bits = _notes_clsid_bits()
    if bits is None:
        return None, "このパソコンに Notes（Lotus.NotesSession）が見つかりません"
    windir = os.environ.get("SystemRoot", r"C:\Windows")
    py_bits = struct.calcsize("P") * 8
    if bits == 32:
        sub = "SysWOW64" if py_bits == 64 else "System32"
    else:
        sub = "Sysnative" if py_bits == 32 else "System32"  # 32 ビットの Python から 64 ビット側を呼ぶとき
    exe = os.path.join(windir, sub, "WindowsPowerShell", "v1.0", "powershell.exe")
    if not os.path.exists(exe):
        return None, f"{bits} ビット版の PowerShell が見つかりません: {exe}"
    return exe, f"このパソコンの Notes（{bits} ビット版）から送信"


def _run_notes(payload, cfg):
    exe, msg = notes_powershell()
    if not exe:
        raise RuntimeError(msg)
    payload = dict(payload, password=cfg["notes_password"])
    try:
        p = subprocess.run(
            [exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", NOTES_SCRIPT],
            input=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            capture_output=True, timeout=NOTES_TIMEOUT,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"Notes が {NOTES_TIMEOUT} 秒たっても応答しません。パスワードの入力待ちで止まっている可能性があります"
                           "（notify.ini の [notes] password を設定するか、Notes の「他の Notes 対応プログラムとパスワードを共有する」を有効にする）")
    out = p.stdout.decode("utf-8", "replace").strip()
    err = p.stderr.decode("utf-8", "replace").strip()
    if p.returncode != 0:
        raise RuntimeError(f"Notes でエラーになりました: {err or out or f'終了コード {p.returncode}'}")
    return out


def notes_check():
    """Notes にログインしてメール DB を開けるか確かめる（送信はしない）。戻り値: Notes のユーザー名"""
    return _run_notes({"mode": "check"}, load_config())


def mark_sent(conn, keys, area=None):
    conn.executemany("INSERT OR IGNORE INTO notify_log(key, area) VALUES (?, ?)", [(k, area) for k in keys])
    conn.commit()


def run(dry_run=False, force=False):
    cfg = load_config()
    conn = connect()
    init_db(conn)
    today = dt.date.today()
    if today.weekday() >= 5 and not force:
        print("土日は送りません（--force で送信）")
        return 0
    messages, warnings = build_messages(conn, cfg, today)
    for w in warnings:
        print("! " + w)
    if not messages:
        print("送る項目はありません")
        return 0
    failed = 0
    for m in messages:
        if dry_run:
            print(f"--- 宛先: {', '.join(m['to'])}\n件名: {m['subject']}\n\n{m['body']}\n")
            continue
        try:
            send_mail(cfg, m["to"], m["subject"], m["body"])
            mark_sent(conn, m["keys"], m["area"])
            print(f"送信しました: {m['area']} → {', '.join(m['to'])}（{m['count']}件）")
        except Exception as e:  # 1 エリアの失敗で他エリアを止めない
            failed += 1
            print(f"! 送信失敗: {m['area']}: {e}", file=sys.stderr)
    return 1 if failed else 0


def init_log():
    conn = connect()
    init_db(conn)
    items = collect_items(conn)
    for it in items:
        mark_sent(conn, [it["key"]], it["area"])
    print(f"今あるアラート {len(items)} 件を通知済みにしました。以降は新しく期限に入ったものだけ送ります。")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="交換アラートのメール通知")
    ap.add_argument("--dry-run", action="store_true", help="送らずに内容を表示する")
    ap.add_argument("--init", action="store_true", help="今あるアラートを通知済みにする")
    ap.add_argument("--force", action="store_true", help="土日でも送る")
    ap.add_argument("--test", metavar="宛先", help="テストメールを送る")
    ap.add_argument("--check", action="store_true", help="Notes で送れるか確認する（送信しない）")
    args = ap.parse_args()
    try:
        if args.check:
            print(notes_powershell()[1])
            print(f"Notes にログインできました: {notes_check()}")
        elif args.init:
            init_log()
        elif args.test:
            cfg = load_config()
            send_mail(cfg, [args.test], *test_message(cfg))
            print(f"テストメールを送信しました: {args.test}")
        else:
            sys.exit(run(dry_run=args.dry_run, force=args.force))
    except RuntimeError as e:
        print(f"! {e}", file=sys.stderr)
        sys.exit(1)
