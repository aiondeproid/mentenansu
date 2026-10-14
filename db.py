"""SQLite スキーマと接続ヘルパー。"""
import os
import sqlite3

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
    sheet_name TEXT
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


def init_db(conn):
    conn.executescript(SCHEMA)
    conn.commit()
