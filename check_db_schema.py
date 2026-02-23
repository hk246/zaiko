"""
データベースのLotテーブルスキーマを確認
"""
import sqlite3
import os
import json

CONFIG_FILE = 'config.json'

def get_database_path():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                db_folder = config.get('database_folder', 'instance')
        except:
            db_folder = 'instance'
    else:
        db_folder = 'instance'
    
    return os.path.join(db_folder, 'inventory.db')

db_path = get_database_path()
print(f"データベースパス: {db_path}\n")

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Lotテーブルの情報を取得
cursor.execute("PRAGMA table_info(lot)")
columns = cursor.fetchall()

print("=== Lotテーブルのカラム一覧 ===")
for col in columns:
    print(f"  {col[1]:20s} {col[2]:15s} NOT NULL={col[3]} DEFAULT={col[4]}")

# is_fractionカラムが存在するか
has_is_fraction = any(col[1] == 'is_fraction' for col in columns)
print(f"\nis_fractionカラム: {'存在します ✓' if has_is_fraction else '存在しません ✗'}")

conn.close()
