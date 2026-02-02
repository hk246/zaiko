"""データベースマイグレーションスクリプト - emailフィールドとLotテーブル、予約機能拡張、レシピ機能、エクセルパス・アクションタイプを追加"""
import sqlite3

conn = sqlite3.connect('instance/inventory.db')
cursor = conn.cursor()

try:
    # emailカラムを追加
    cursor.execute('ALTER TABLE raw_material ADD COLUMN email VARCHAR(120)')
    conn.commit()
    print("✓ emailカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ emailカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # excel_pathカラムを追加
    cursor.execute('ALTER TABLE raw_material ADD COLUMN excel_path VARCHAR(500)')
    conn.commit()
    print("✓ excel_pathカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ excel_pathカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # action_typeカラムを追加（デフォルト値: 'none'）
    cursor.execute("ALTER TABLE raw_material ADD COLUMN action_type VARCHAR(20) DEFAULT 'none'")
    conn.commit()
    print("✓ action_typeカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ action_typeカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Lotテーブルを作成
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS lot (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            material_id INTEGER NOT NULL,
            lot_name VARCHAR(100) NOT NULL,
            weight FLOAT NOT NULL,
            date_created DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (material_id) REFERENCES raw_material (id)
        )
    ''')
    conn.commit()
    print("✓ Lotテーブルを作成しました")
except sqlite3.Error as e:
    print(f"Lotテーブル作成エラー: {e}")

try:
    # Reservationテーブルにlot_idカラムを追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN lot_id INTEGER')
    conn.commit()
    print("✓ Reservationテーブルにlot_idカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ lot_idカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Reservationテーブルにlot_nameカラムを追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN lot_name VARCHAR(100)')
    conn.commit()
    print("✓ Reservationテーブルにlot_nameカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ lot_nameカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Reservationテーブルにscheduled_dateカラムを追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN scheduled_date DATE')
    conn.commit()
    print("✓ Reservationテーブルにscheduled_dateカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ scheduled_dateカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Reservationテーブルにexecutedカラムを追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN executed BOOLEAN DEFAULT 0')
    conn.commit()
    print("✓ Reservationテーブルにexecutedカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ executedカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Reservationテーブルにrecipe_idカラムを追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN recipe_id INTEGER')
    conn.commit()
    print("✓ Reservationテーブルにrecipe_idカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ recipe_idカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Reservationテーブルにactual_quantityカラムを追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN actual_quantity FLOAT')
    conn.commit()
    print("✓ Reservationテーブルにactual_quantityカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ actual_quantityカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Reservationテーブルにuser_nameカラムを追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN user_name VARCHAR(100)')
    conn.commit()
    print("✓ Reservationテーブルにuser_nameカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ user_nameカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Reservationテーブルにpurposeカラムを追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN purpose VARCHAR(200)')
    conn.commit()
    print("✓ Reservationテーブルにpurposeカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ purposeカラムは既に存在します")
    else:
        print(f"エラー: {e}")

try:
    # Recipeテーブルを作成
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS recipe (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name VARCHAR(100) NOT NULL,
            description VARCHAR(200),
            type VARCHAR(20) NOT NULL,
            date_created DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    print("✓ Recipeテーブルを作成しました")
except sqlite3.Error as e:
    print(f"Recipeテーブル作成エラー: {e}")

try:
    # RecipeItemテーブルを作成
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS recipe_item (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            recipe_id INTEGER NOT NULL,
            material_id INTEGER NOT NULL,
            quantity FLOAT NOT NULL,
            lot_name VARCHAR(100),
            FOREIGN KEY (recipe_id) REFERENCES recipe (id),
            FOREIGN KEY (material_id) REFERENCES raw_material (id)
        )
    ''')
    conn.commit()
    print("✓ RecipeItemテーブルを作成しました")
except sqlite3.Error as e:
    print(f"RecipeItemテーブル作成エラー: {e}")

try:
    # executed_dateカラムをReservationテーブルに追加
    cursor.execute('ALTER TABLE reservation ADD COLUMN executed_date DATETIME')
    conn.commit()
    print("✓ executed_dateカラムを追加しました")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✓ executed_dateカラムは既に存在します")
    else:
        print(f"エラー: {e}")

conn.close()
print("\n✅ マイグレーション完了！")



