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
    location TEXT,             -- 保管場所
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

-- 部品一覧（価格表）。シリアルを持たない交換部品・消耗品
CREATE TABLE IF NOT EXISTS part_catalog (
    id INTEGER PRIMARY KEY,
    group_name TEXT,           -- 区分（Aライン・コンベアネット など。価格表のシート名）
    machine TEXT,              -- 機械・場所（DeJong成型機 など。価格表の見出し）
    name TEXT NOT NULL,
    unit_price INTEGER,        -- 単価（円）
    qty REAL,                  -- 数量（一度に使う数）
    note TEXT,                 -- 備考
    memo TEXT,                 -- 補足（価格表の E 列）
    location TEXT,             -- 保管場所
    equipment_id INTEGER REFERENCES equipment(id),
    area TEXT,
    sort_order INTEGER,
    source TEXT,               -- 取り込み元ファイル名 / 'app' / 'excel'
    updated_at TEXT
);

-- プラ・樹脂箇所の点検（プラ点検）
CREATE TABLE IF NOT EXISTS pla_items (        -- 点検箇所（プラ樹脂点検リストの 1 行）
    id INTEGER PRIMARY KEY,
    process TEXT NOT NULL,     -- 前工程 / 後工程 / タレ工程
    equipment TEXT NOT NULL,   -- 設備名（点検表の表記のまま）
    item TEXT NOT NULL,        -- 機械・器具
    color TEXT,
    material TEXT,             -- 材質大括り（ぷらしる）
    material_pct TEXT,         -- ％（ぷらしる）
    material_maker TEXT,       -- 材質（メーカー提示）
    hazard TEXT,               -- 危害度 A〜D
    methods TEXT,              -- 点検方法（始業前点検表,プラリスク点検表 などをカンマ区切り）
    is_new INTEGER DEFAULT 0,
    frequency TEXT,            -- ２回／日 など
    severity TEXT, possibility TEXT, evaluation TEXT,
    note TEXT,
    active INTEGER DEFAULT 1,  -- 0 = 撤去・項目から外した（記録のため残す）
    sort_order INTEGER,
    source TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS pla_sheets (       -- チェック表（紙の「プラスチック部品チェック表」1 枚）
    id INTEGER PRIMARY KEY,
    process TEXT NOT NULL,
    file TEXT, sheet_name TEXT,
    title TEXT NOT NULL,
    active INTEGER DEFAULT 1,  -- 0 = 古い版・撤去済み（記録のため残す）
    sort_order INTEGER,
    note TEXT,                 -- 改定履歴など
    source TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS pla_points (       -- チェック表の 1 行（点検する箇所）
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES pla_sheets(id),
    location TEXT,             -- 点検箇所（場所）
    item_text TEXT NOT NULL,   -- 項目
    sort_order INTEGER,
    active INTEGER DEFAULT 1
);
CREATE TABLE IF NOT EXISTS pla_point_items (  -- チェック表の行 ⇔ プラリストの行
    point_id INTEGER NOT NULL REFERENCES pla_points(id),
    item_id INTEGER NOT NULL REFERENCES pla_items(id),
    PRIMARY KEY (point_id, item_id)
);
CREATE TABLE IF NOT EXISTS pla_photos (       -- 写真（point_id が空ならチェック表全体の「場所の写真」）
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES pla_sheets(id),
    point_id INTEGER REFERENCES pla_points(id),
    file TEXT NOT NULL,        -- data/pla_photos/ の中のファイル名
    sort_order INTEGER,
    caption TEXT,
    source TEXT,               -- xls / app
    created_at TEXT
);
CREATE TABLE IF NOT EXISTS pla_checks (       -- 点検結果（箇所 × 日 × 始業時/終業時）
    id INTEGER PRIMARY KEY,
    point_id INTEGER NOT NULL REFERENCES pla_points(id),
    check_date TEXT NOT NULL,
    timing TEXT NOT NULL,      -- 始業時 / 終業時
    result TEXT NOT NULL,      -- ok / ng / na（稼働なし）
    note TEXT,                 -- 異常の内容
    checked_by TEXT,
    checked_at TEXT,
    UNIQUE(point_id, check_date, timing)
);
CREATE TABLE IF NOT EXISTS pla_reviews (      -- 確認（工程 × 日 × 始業時/終業時 ごとに確認者が見た記録）
    id INTEGER PRIMARY KEY,
    process TEXT NOT NULL,
    check_date TEXT NOT NULL,
    timing TEXT NOT NULL,
    reviewer TEXT,
    reviewed_at TEXT,
    comment TEXT,
    UNIQUE(process, check_date, timing)
);
CREATE INDEX IF NOT EXISTS idx_pla_checks_date ON pla_checks(check_date);
CREATE INDEX IF NOT EXISTS idx_pla_points_sheet ON pla_points(sheet_id);

-- 作業記録の添付ファイル（PDF・写真）
CREATE TABLE IF NOT EXISTS record_files (
    id INTEGER PRIMARY KEY,
    record_id INTEGER NOT NULL REFERENCES records(id),
    file TEXT NOT NULL,        -- data/record_files/ の中のファイル名
    name TEXT,                 -- 元のファイル名（表示・ダウンロード用）
    size INTEGER,
    created_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_record_files_rec ON record_files(record_id);

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
    # 箇所ごとの点検結果（pla_results）はチェック表ごと（pla_checks）に置き換えた。記録がなければ消す
    if conn.execute("SELECT 1 FROM sqlite_master WHERE name = 'pla_results'").fetchone():
        if not conn.execute("SELECT 1 FROM pla_results LIMIT 1").fetchone():
            conn.execute("DROP TABLE pla_results")
    for table in ("parts", "part_catalog"):
        if "location" not in {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN location TEXT")  # 保管場所
    for e in conn.execute("SELECT id, name, sheet_name FROM equipment WHERE category IS NULL").fetchall():
        conn.execute("UPDATE equipment SET category=?, nickname=COALESCE(nickname, ?) WHERE id=?",
                     (guess_category(e["sheet_name"], e["name"]), make_nickname(e["sheet_name"], e["name"]), e["id"]))
    conn.commit()
