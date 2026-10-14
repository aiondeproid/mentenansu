"""SQLite スキーマと接続ヘルパー。"""
import os
import re
import sqlite3
import unicodedata

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("MAINT_DB", os.path.join(BASE_DIR, "data", "maintenance.db"))

SCHEMA = """
CREATE TABLE IF NOT EXISTS vendors (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    tel TEXT, fax TEXT, address TEXT, contact TEXT, mobile TEXT
);

CREATE TABLE IF NOT EXISTS equipment (
    id INTEGER PRIMARY KEY,
    code TEXT,                 -- 設備No (K001 など)
    name TEXT NOT NULL,
    model TEXT,
    maker TEXT,
    vendor_id INTEGER REFERENCES vendors(id),
    supplier TEXT,             -- 納入業者名
    asset_no TEXT,
    serial_no TEXT,
    manufactured TEXT,
    delivered TEXT,
    area TEXT DEFAULT '前処理',
    status TEXT DEFAULT '稼働', -- 稼働 / 廃止
    note TEXT,
    current_hours REAL,        -- 最新アワーメーター値
    hours_updated TEXT,
    sheet_name TEXT,
    source_file TEXT,
    nickname TEXT,             -- 通称（Excel のシート名。現場での呼び名）
    category TEXT              -- 種類（成型・包装など。台帳のグループ分けに使う）
);

-- 作業履歴（Excel の 1 行 = 1 レコード）
CREATE TABLE IF NOT EXISTS records (
    id INTEGER PRIMARY KEY,
    equipment_id INTEGER NOT NULL REFERENCES equipment(id),
    kind TEXT,                 -- 定期 / 不定期
    status TEXT DEFAULT '完了',
    done_date TEXT,            -- YYYY-MM-DD（不明なら NULL）
    date_raw TEXT,             -- 取り込み元の表記
    date_uncertain INTEGER DEFAULT 0,
    downtime_min INTEGER,      -- ライン停止時間（分）
    downtime_raw TEXT,
    hour_meter REAL,           -- 部品・設備の稼働時間（h）
    content TEXT, reason TEXT, result TEXT, countermeasure TEXT,
    author TEXT,
    cost INTEGER,              -- 費用（円）
    part_serial TEXT,
    rule_id INTEGER REFERENCES rules(id),
    source TEXT,               -- 'xls' / 'app'
    created_at TEXT DEFAULT (datetime('now','localtime')),
    updated_at TEXT
);

-- 定期交換・点検ルール（作業計画）
CREATE TABLE IF NOT EXISTS rules (
    id INTEGER PRIMARY KEY,
    equipment_id INTEGER NOT NULL REFERENCES equipment(id),
    title TEXT NOT NULL,
    interval_days INTEGER,
    interval_hours REAL,
    last_done_date TEXT,
    last_hours REAL,
    keyword TEXT,              -- 履歴から前回実施日を拾うキーワード
    note TEXT,
    active INTEGER DEFAULT 1
);

-- シリアル管理部品（フィルタースクリュー等）
CREATE TABLE IF NOT EXISTS parts (
    id INTEGER PRIMARY KEY,
    part_type TEXT NOT NULL,
    serial TEXT NOT NULL UNIQUE,
    equipment_id INTEGER REFERENCES equipment(id),
    grind_count INTEGER DEFAULT 0,
    grind_limit INTEGER DEFAULT 10,
    status TEXT DEFAULT '予備', -- 使用中 / 予備 / 研磨中 / 使用禁止 / 廃棄
    note TEXT
);

CREATE TABLE IF NOT EXISTS part_events (
    id INTEGER PRIMARY KEY,
    part_id INTEGER NOT NULL REFERENCES parts(id),
    event_date TEXT,
    event TEXT,                -- 装着 / 取外し / 研磨 / 返品 / 廃棄 / 交換記録 / メモ
    hour_meter REAL,
    record_id INTEGER REFERENCES records(id),
    note TEXT
);

-- 交換アラートの通知先（area='*' は全エリア共通）
CREATE TABLE IF NOT EXISTS notify_recipients (
    id INTEGER PRIMARY KEY,
    area TEXT NOT NULL,
    email TEXT NOT NULL,
    UNIQUE(area, email)
);

-- 通知済みの記録（同じ内容を毎朝送らないため）
CREATE TABLE IF NOT EXISTS notify_log (
    key TEXT PRIMARY KEY,      -- rule:<id>:<前回実施日>:<状態> / part:<id>:<研磨回数>:<状態>
    area TEXT,
    sent_at TEXT DEFAULT (datetime('now','localtime'))
);

CREATE INDEX IF NOT EXISTS idx_records_eq ON records(equipment_id);
CREATE INDEX IF NOT EXISTS idx_records_date ON records(done_date);
CREATE INDEX IF NOT EXISTS idx_part_events_part ON part_events(part_id);
"""


def connect(path=None):
    path = path or DB_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# 設備の種類（上から順に判定。台帳でもこの順に並べる）
CATEGORY_RULES = [
    ("検査機（X線・金検）", r"X線|金検|ウェイト"),
    ("冷凍・冷蔵・解凍", r"フリーザー|予冷|冷蔵|冷凍|解凍|製氷"),
    ("成型", r"成型|包餡|包あん|丸目|ハントマン|メガフォーマ|ロトフォーム"),
    ("包装・充填・計数", r"包装|ムルチ|ポリスター|封函|カートン|計数|インクジェット|印字|充填"),
    ("搬送", r"コンベア|バーチカル|シャトル"),
    ("切断・粉砕", r"カッター|ダイサー|チョッパー|グラインダー|セパレーター|フレーカー"),
    ("配合・混合", r"配合|ミキサー|ニーダー"),
    ("洗浄", r"発泡|洗浄|ジェッター"),
    ("空調・換気", r"GHP|ガスヒート|換気|ファン|排気"),
    ("加熱・殺菌", r"焼き|蒸煮|フライヤー|炒め|殺菌"),
    ("排水処理", r"排水|汚泥|原水|調整槽|加圧浮上|脱水"),
    ("ボイラー・蒸気", r"ボイラ|蒸気"),
    ("給水・圧縮空気", r"コンプレッサー|受水|貯水|井水"),
]
# 台帳での表示順（工程の流れ → 検査・搬送 → 付帯設備）
CATEGORIES = [
    "切断・粉砕", "配合・混合", "成型", "加熱・殺菌", "冷凍・冷蔵・解凍", "包装・充填・計数",
    "検査機（X線・金検）", "搬送", "洗浄", "ボイラー・蒸気", "給水・圧縮空気", "空調・換気", "排水処理", "その他",
]
assert set(CATEGORIES) == {c for c, _ in CATEGORY_RULES} | {"その他"}


def norm(s):
    """半角カナと全角英数だけを正規化する（①や㈱はそのまま残す）。"""
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"[｡-ﾟ]+", lambda m: unicodedata.normalize("NFKC", m.group()), s)
    s = re.sub(r"[！-～]+", lambda m: unicodedata.normalize("NFKC", m.group()), s)
    return s.replace("　", " ").strip()


def guess_category(*texts):
    joined = " ".join(norm(t) for t in texts if t)
    for category, pattern in CATEGORY_RULES:
        if re.search(pattern, joined):
            return category
    return "その他"


def make_nickname(sheet_name, name):
    """シート名を通称にする（「Sheet1」や複製の「 (2)」は除く。正式名と同じなら不要）。"""
    nick = re.sub(r"\s*\(\d+\)$", "", norm(sheet_name or ""))
    if not nick or re.fullmatch(r"Sheet\d+", nick) or nick == norm(name):
        return None
    return nick


def init_db(conn):
    conn.executescript(SCHEMA)
    # 既存 DB への列追加（取り込み済みデータは残したまま）
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(equipment)")}
    for col in ("source_file", "nickname", "category"):
        if col not in cols:
            conn.execute(f"ALTER TABLE equipment ADD COLUMN {col} TEXT")
    for e in conn.execute("SELECT id, name, sheet_name FROM equipment WHERE category IS NULL").fetchall():
        conn.execute("UPDATE equipment SET category=?, nickname=COALESCE(nickname, ?) WHERE id=?",
                     (guess_category(e["sheet_name"], e["name"]), make_nickname(e["sheet_name"], e["name"]), e["id"]))
    conn.commit()
