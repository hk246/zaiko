"""
データベースマイグレーション: Lotテーブルにis_fractionカラムを追加
"""
import sqlite3
import os
from pathlib import Path

# データベースパスを取得（config.jsonから）
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

if not os.path.exists(db_path):
    print(f"エラー: データベースが見つかりません: {db_path}")
    exit(1)

print(f"データベースパス: {db_path}")

# データベースに接続
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

try:
    # is_fractionカラムが存在するか確認
    cursor.execute("PRAGMA table_info(lot)")
    columns = [column[1] for column in cursor.fetchall()]
    
    if 'is_fraction' in columns:
        print("✓ is_fractionカラムは既に存在します")
    else:
        print("is_fractionカラムを追加中...")
        # カラムを追加（デフォルト値: False = 0）
        cursor.execute("ALTER TABLE lot ADD COLUMN is_fraction BOOLEAN DEFAULT 0")
        conn.commit()
        print("✓ is_fractionカラムを追加しました")
    
    # Contactテーブルが存在するか確認して、なければ作成
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='contact'")
    if not cursor.fetchone():
        print("Contactテーブルを作成中...")
        cursor.execute("""
            CREATE TABLE contact (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name VARCHAR(100) NOT NULL,
                email VARCHAR(120) NOT NULL,
                description VARCHAR(200),
                date_created DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        print("✓ Contactテーブルを作成しました")
    else:
        print("✓ Contactテーブルは既に存在します")
    
    # work_orderテーブルが存在するか確認して、なければ作成（worker_nameカラム含む）
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='work_order'")
    if not cursor.fetchone():
        print("work_orderテーブルを作成中...")
        cursor.execute("""
            CREATE TABLE work_order (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                process_name VARCHAR(200) NOT NULL,
                lot_name VARCHAR(100),
                material_id INTEGER REFERENCES raw_material(id) ON DELETE SET NULL,
                status VARCHAR(20) DEFAULT 'pending',
                priority VARCHAR(20) DEFAULT 'none',
                invoice_issued BOOLEAN DEFAULT 0,
                worker_name VARCHAR(100),
                experiment_type VARCHAR(20) DEFAULT 'standard',
                planned_start DATE,
                planned_end DATE,
                actual_start DATE,
                actual_end DATE,
                notes TEXT,
                archived BOOLEAN DEFAULT 0,
                date_created DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        print("✓ work_orderテーブルを作成しました")
    else:
        print("✓ work_orderテーブルは既に存在します")
        # worker_nameカラムが存在するか確認
        cursor.execute("PRAGMA table_info(work_order)")
        wo_columns = [col[1] for col in cursor.fetchall()]
        if 'worker_name' not in wo_columns:
            print("work_orderテーブルにworker_nameカラムを追加中...")
            cursor.execute("ALTER TABLE work_order ADD COLUMN worker_name VARCHAR(100)")
            conn.commit()
            print("✓ worker_nameカラムを追加しました")
        else:
            print("✓ worker_nameカラムは既に存在します")

    # material_labelテーブルが存在するか確認して、なければ作成
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='material_label'")
    if not cursor.fetchone():
        print("material_labelテーブルを作成中...")
        cursor.execute("""
            CREATE TABLE material_label (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name VARCHAR(50) NOT NULL UNIQUE,
                color VARCHAR(7) NOT NULL DEFAULT '#6c757d',
                date_created DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        print("material_labelテーブルを作成しました")
    else:
        print("material_labelテーブルは既に存在します")

    # raw_materialテーブルにlabel_idカラムを追加
    cursor.execute("PRAGMA table_info(raw_material)")
    rm_cols = [r[1] for r in cursor.fetchall()]
    if 'label_id' not in rm_cols:
        print("raw_materialテーブルにlabel_idカラムを追加中...")
        cursor.execute("ALTER TABLE raw_material ADD COLUMN label_id INTEGER REFERENCES material_label(id)")
        conn.commit()
        print("label_idカラムを追加しました")
    else:
        print("label_idカラムは既に存在します")

    # property_fieldテーブルが存在するか確認して、なければ作成
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='property_field'")
    if not cursor.fetchone():
        print("property_fieldテーブルを作成中...")
        cursor.execute("""
            CREATE TABLE property_field (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                label_id INTEGER REFERENCES material_label(id) ON DELETE CASCADE,
                name VARCHAR(100) NOT NULL,
                field_type VARCHAR(20) NOT NULL DEFAULT 'string',
                unit VARCHAR(30),
                order_index INTEGER DEFAULT 0,
                date_created DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        print("✓ property_fieldテーブルを作成しました")
    else:
        print("✓ property_fieldテーブルは既に存在します")

    # lot_propertyテーブルが存在するか確認して、なければ作成
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='lot_property'")
    if not cursor.fetchone():
        print("lot_propertyテーブルを作成中...")
        cursor.execute("""
            CREATE TABLE lot_property (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                lot_id INTEGER NOT NULL REFERENCES lot(id) ON DELETE CASCADE,
                field_id INTEGER NOT NULL REFERENCES property_field(id) ON DELETE CASCADE,
                value_string VARCHAR(500),
                value_number REAL,
                date_updated DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(lot_id, field_id)
            )
        """)
        conn.commit()
        print("✓ lot_propertyテーブルを作成しました")
    else:
        print("✓ lot_propertyテーブルは既に存在します")

    print("\nマイグレーション完了！")

except Exception as e:
    print(f"\n❌ エラーが発生しました: {e}")
    conn.rollback()
finally:
    conn.close()
