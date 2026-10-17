"""data フォルダ（DB・点検写真・作業記録の添付ファイル）をサーバーの共有フォルダへバックアップする。
タスクスケジューラから毎日実行する想定。アプリを動かしたままでも実行できる。

    python backup.py \\\\fileserver\\share\\setsubi_backup    指定した場所へバックアップ
    python backup.py                                       環境変数 MAINT_BACKUP_DIR の場所へ
    python backup.py 保存先 --keep 60                       DB を 60 世代残す（既定 30）

保存先の中身:
    db/maintenance_YYYYMMDD_HHMMSS.db   DB の世代バックアップ（古いものから消す）
    pla_photos/                         点検写真（追加・変更されたものだけコピー。アプリで消した写真も残す）
    record_files/                       作業記録の添付ファイル（PDF・写真。コピーの仕方は pla_photos と同じ）

戻すとき: アプリを止めて、残したい世代の .db を data/maintenance.db に、
pla_photos/・record_files/ を data/ の同じ名前のフォルダにコピーしてから起動する。
"""
import argparse
import datetime as dt
import os
import shutil
import sqlite3
import sys

from db import DB_PATH

DATA_DIR = os.path.dirname(DB_PATH)
FILE_DIRS = [("pla_photos", "写真"), ("record_files", "添付")]  # 差分だけコピーするフォルダ


def backup_db(dest_dir, keep):
    db_dir = os.path.join(dest_dir, "db")
    os.makedirs(db_dir, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(db_dir, f"maintenance_{stamp}.db")
    tmp = out + ".tmp"

    # 書き込み中でも整合性のとれた写しを作る（ファイルをそのままコピーすると壊れることがある）
    src = sqlite3.connect(DB_PATH)
    dst = sqlite3.connect(tmp)
    try:
        src.backup(dst)
        ok = dst.execute("PRAGMA integrity_check").fetchone()[0]
    finally:
        dst.close()
        src.close()
    if ok != "ok":
        os.remove(tmp)
        raise RuntimeError(f"バックアップした DB の検査で異常: {ok}")
    os.replace(tmp, out)

    olds = sorted(f for f in os.listdir(db_dir) if f.startswith("maintenance_") and f.endswith(".db"))
    removed = olds[:-keep] if keep > 0 else []
    for f in removed:
        os.remove(os.path.join(db_dir, f))
    return out, os.path.getsize(out), len(removed)


def backup_files(dest_dir, folder):
    src_dir = os.path.join(DATA_DIR, folder)
    if not os.path.isdir(src_dir):
        return 0, 0
    out_dir = os.path.join(dest_dir, folder)
    os.makedirs(out_dir, exist_ok=True)
    # 保存先は 1 回の一覧取得でまとめて読む（ネットワーク越しに 1 件ずつ調べると遅い）
    have = {}
    with os.scandir(out_dir) as it:
        for e in it:
            if e.is_file():
                st = e.stat()
                have[e.name] = (st.st_size, int(st.st_mtime))
    copied = total = 0
    with os.scandir(src_dir) as it:
        for e in it:
            if not e.is_file():
                continue
            total += 1
            st = e.stat()
            old = have.get(e.name)
            if old and old[0] == st.st_size and old[1] >= int(st.st_mtime):
                continue
            shutil.copy2(e.path, os.path.join(out_dir, e.name))
            copied += 1
    return copied, total


def main():
    ap = argparse.ArgumentParser(description="data フォルダをサーバーへバックアップする")
    ap.add_argument("dest", nargs="?", default=os.environ.get("MAINT_BACKUP_DIR", ""),
                    help=r"保存先（例: \\fileserver\share\setsubi_backup）")
    ap.add_argument("--keep", type=int, default=30, help="DB を何世代残すか（0 で全部残す。既定 30）")
    args = ap.parse_args()

    if not args.dest:
        ap.error("保存先を指定するか、環境変数 MAINT_BACKUP_DIR を設定してください")
    if not os.path.exists(DB_PATH):
        sys.exit(f"[エラー] DB がありません: {DB_PATH}")
    try:
        os.makedirs(args.dest, exist_ok=True)
    except OSError as e:
        sys.exit(f"[エラー] 保存先に書き込めません: {args.dest}（{e}）")

    out, size, removed = backup_db(args.dest, args.keep)
    print(f"DB  : {out}（{size / 1024 / 1024:.1f} MB）" + (f" 古い世代を {removed} 件削除" if removed else ""))
    for folder, label in FILE_DIRS:
        copied, total = backup_files(args.dest, folder)
        print(f"{label}: {copied} 件コピー（全 {total} 件）")


if __name__ == "__main__":
    main()
