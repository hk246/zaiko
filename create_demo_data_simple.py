"""
シンプルなデモデータベース作成スクリプト
SQLiteを直接使用してデモデータを作成します
"""
import sqlite3
import os
from datetime import datetime

def create_demo_database():
    """デモ用のデータベースを作成"""
    
    # バックアップフォルダを確保
    backup_folder = 'backups'
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
    
    demo_db_path = os.path.join(backup_folder, 'inventory_DEMO_データ.db')
    
    # 既存のデモDBがあれば削除
    if os.path.exists(demo_db_path):
        os.remove(demo_db_path)
    
    # データベース接続
    conn = sqlite3.connect(demo_db_path)
    cursor = conn.cursor()
    
    # テーブル作成
    cursor.execute('''
        CREATE TABLE raw_material (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name VARCHAR(100) NOT NULL,
            weight FLOAT NOT NULL,
            unit VARCHAR(20) DEFAULT 'kg',
            min_weight FLOAT DEFAULT 0.0
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE lot (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            material_id INTEGER NOT NULL,
            lot_name VARCHAR(100) NOT NULL,
            weight FLOAT NOT NULL,
            date_created DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (material_id) REFERENCES raw_material (id)
        )
    ''')
    
    # デモデータ
    materials = [
        ('小麦粉（強力粉）', 'kg', 50.0, [
            ('A-2024-001', 125.5),
            ('A-2024-002', 89.2),
            ('A-2024-003', 45.8),
        ]),
        ('グラニュー糖', 'kg', 30.0, [
            ('S-2024-015', 78.3),
            ('S-2024-016', 42.1),
        ]),
        ('無塩バター', 'kg', 20.0, [
            ('B-240201', 35.6),
            ('B-240208', 28.9),
            ('B-240215', 15.2),
        ]),
        ('卵（全卵液）', 'kg', 15.0, [
            ('EGG-240205', 22.5),
            ('EGG-240207', 18.3),
        ]),
        ('牛乳', 'L', 25.0, [
            ('MILK-0205', 45.0),
            ('MILK-0207', 32.5),
        ]),
        ('チョコレート（カカオ70%）', 'kg', 10.0, [
            ('CHO-2024-A', 18.5),
            ('CHO-2024-B', 12.8),
        ]),
        ('ベーキングパウダー', 'kg', 3.0, [
            ('BP-2024-001', 5.2),
        ]),
        ('バニラエッセンス', 'mL', 500.0, [
            ('VAN-500ML-001', 850.0),
            ('VAN-500ML-002', 420.0),
        ]),
        ('アーモンドプードル', 'kg', 5.0, [
            ('ALM-2024-01', 12.3),
        ]),
        ('生クリーム（乳脂肪45%）', 'L', 10.0, [
            ('CRM-240206', 15.5),
            ('CRM-240208', 8.2),
        ]),
    ]
    
    # データを挿入
    for material_name, unit, min_weight, lots in materials:
        # 合計重量を計算
        total_weight = sum(lot_weight for _, lot_weight in lots)
        
        # 原料を挿入
        cursor.execute('''
            INSERT INTO raw_material (name, weight, unit, min_weight)
            VALUES (?, ?, ?, ?)
        ''', (material_name, total_weight, unit, min_weight))
        
        material_id = cursor.lastrowid
        
        # ロットを挿入
        for lot_name, lot_weight in lots:
            cursor.execute('''
                INSERT INTO lot (material_id, lot_name, weight)
                VALUES (?, ?, ?)
            ''', (material_id, lot_name, lot_weight))
    
    conn.commit()
    conn.close()
    
    print("="  * 60)
    print("✅ デモデータベースの作成が完了しました！")
    print("=" * 60)
    print(f"\n📊 作成されたデータ:")
    print(f"  - 原料数: {len(materials)} 種類")
    total_lots = sum(len(lots) for _, _, _, lots in materials)
    print(f"  - ロット数: {total_lots} 個")
    print(f"\n💾 保存場所: {demo_db_path}")
    print("\n📍 使い方:")
    print("  1. アプリを起動してバックアップ管理画面（/backup）に移動")
    print("  2. 'inventory_DEMO_データ.db' を見つける")
    print("  3. 「復元」ボタンをクリックしてデモデータを読み込む")
    print("\n✨ これでアプリの全機能をデモデータで試すことができます！")
    print("=" * 60)

if __name__ == '__main__':
    create_demo_database()
