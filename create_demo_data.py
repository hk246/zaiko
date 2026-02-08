"""
デモ用データベース作成スクリプト
このアプリの機能を説明するためのサンプルデータを作成します
"""
import os
import shutil
from datetime import datetime
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

# デモ用の一時アプリケーションを作成
demo_app = Flask(__name__)
demo_app.config['SECRET_KEY'] = 'demo-secret-key'

# バックアップフォルダに直接作成
backup_folder = 'backups'
if not os.path.exists(backup_folder):
    os.makedirs(backup_folder)

# 絶対パスを使用
demo_db_filename = 'inventory_DEMO_データ.db'
demo_db_path = os.path.abspath(os.path.join(backup_folder, demo_db_filename))
demo_app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{demo_db_path}'
demo_db = SQLAlchemy(demo_app)

# モデル定義（app.pyから複製）
class RawMaterial(demo_db.Model):
    id = demo_db.Column(demo_db.Integer, primary_key=True)
    name = demo_db.Column(demo_db.String(100), nullable=False)
    weight = demo_db.Column(demo_db.Float, nullable=False)
    unit = demo_db.Column(demo_db.String(20), default='kg')
    min_weight = demo_db.Column(demo_db.Float, default=0.0)

class Lot(demo_db.Model):
    id = demo_db.Column(demo_db.Integer, primary_key=True)
    material_id = demo_db.Column(demo_db.Integer, demo_db.ForeignKey('raw_material.id'), nullable=False)
    lot_name = demo_db.Column(demo_db.String(100), nullable=False)
    weight = demo_db.Column(demo_db.Float, nullable=False)
    date_created = demo_db.Column(demo_db.DateTime, default=demo_db.func.current_timestamp())
    
    material = demo_db.relationship('RawMaterial', backref=demo_db.backref('lots', lazy=True))

def create_demo_database():
    """デモ用のデータベースを作成"""
    
    # 既存のデータベースがあればバックアップ
    main_db_path = 'instance/inventory.db'
    if os.path.exists(main_db_path):
        backup_path = f'instance/inventory_backup_before_demo_{datetime.now().strftime("%Y%m%d_%H%M%S")}.db'
        shutil.copy2(main_db_path, backup_path)
        print(f"既存の本番データベースをバックアップしました: {backup_path}\n")
    
    with demo_app.app_context():
        # 既存のデモデータベースがあれば削除
        if os.path.exists(demo_db_path):
            os.remove(demo_db_path)
        
        # テーブルを作成
        demo_db.create_all()
        
        # デモデータ: 原料とロット
        demo_data = [
            {
                'material': {'name': '小麦粉（強力粉）', 'unit': 'kg', 'min_weight': 50.0},
                'lots': [
                    {'lot_name': 'A-2024-001', 'weight': 125.5},
                    {'lot_name': 'A-2024-002', 'weight': 89.2},
                    {'lot_name': 'A-2024-003', 'weight': 45.8},
                ]
            },
            {
                'material': {'name': 'グラニュー糖', 'unit': 'kg', 'min_weight': 30.0},
                'lots': [
                    {'lot_name': 'S-2024-015', 'weight': 78.3},
                    {'lot_name': 'S-2024-016', 'weight': 42.1},
                ]
            },
            {
                'material': {'name': '無塩バター', 'unit': 'kg', 'min_weight': 20.0},
                'lots': [
                    {'lot_name': 'B-240201', 'weight': 35.6},
                    {'lot_name': 'B-240208', 'weight': 28.9},
                    {'lot_name': 'B-240215', 'weight': 15.2},
                ]
            },
            {
                'material': {'name': '卵（全卵液）', 'unit': 'kg', 'min_weight': 15.0},
                'lots': [
                    {'lot_name': 'EGG-240205', 'weight': 22.5},
                    {'lot_name': 'EGG-240207', 'weight': 18.3},
                ]
            },
            {
                'material': {'name': '牛乳', 'unit': 'L', 'min_weight': 25.0},
                'lots': [
                    {'lot_name': 'MILK-0205', 'weight': 45.0},
                    {'lot_name': 'MILK-0207', 'weight': 32.5},
                ]
            },
            {
                'material': {'name': 'チョコレート（カカオ70%）', 'unit': 'kg', 'min_weight': 10.0},
                'lots': [
                    {'lot_name': 'CHO-2024-A', 'weight': 18.5},
                    {'lot_name': 'CHO-2024-B', 'weight': 12.8},
                ]
            },
            {
                'material': {'name': 'ベーキングパウダー', 'unit': 'kg', 'min_weight': 3.0},
                'lots': [
                    {'lot_name': 'BP-2024-001', 'weight': 5.2},
                ]
            },
            {
                'material': {'name': 'バニラエッセンス', 'unit': 'mL', 'min_weight': 500.0},
                'lots': [
                    {'lot_name': 'VAN-500ML-001', 'weight': 850.0},
                    {'lot_name': 'VAN-500ML-002', 'weight': 420.0},
                ]
            },
            {
                'material': {'name': 'アーモンドプードル', 'unit': 'kg', 'min_weight': 5.0},
                'lots': [
                    {'lot_name': 'ALM-2024-01', 'weight': 12.3},
                ]
            },
            {
                'material': {'name': '生クリーム（乳脂肪45%）', 'unit': 'L', 'min_weight': 10.0},
                'lots': [
                    {'lot_name': 'CRM-240206', 'weight': 15.5},
                    {'lot_name': 'CRM-240208', 'weight': 8.2},
                ]
            },
        ]
        
        # データを追加
        for item in demo_data:
            # 原料を作成
            material = RawMaterial(
                name=item['material']['name'],
                weight=0,  # 初期値（ロットで計算される）
                unit=item['material']['unit'],
                min_weight=item['material']['min_weight']
            )
            demo_db.session.add(material)
            demo_db.session.flush()  # IDを取得するため
            
            # ロットを作成
            total_weight = 0
            for lot_data in item['lots']:
                lot = Lot(
                    material_id=material.id,
                    lot_name=lot_data['lot_name'],
                    weight=lot_data['weight']
                )
                demo_db.session.add(lot)
                total_weight += lot_data['weight']
            
            # 原料の合計重量を更新
            material.weight = total_weight
        
        # コミット
        demo_db.session.commit()
        
        print("✅ デモデータベースの作成が完了しました！")
        print(f"\n📊 作成されたデータ:")
        print(f"  - 原料数: {len(demo_data)} 種類")
        total_lots = sum(len(item['lots']) for item in demo_data)
        print(f"  - ロット数: {total_lots} 個")
        
        # データベース接続を閉じる（app_context内で）
        demo_db.session.close()
    
    print(f"\n💾 デモデータベースを保存しました: {demo_db_path}")
    print("\n📍 使い方:")
    print("  1. アプリを起動してバックアップ管理画面に移動")
    print("  2. 'inventory_DEMO_データ.db' を見つける")
    print("  3. 「復元」ボタンをクリックしてデモデータを読み込む")
    print("\n✨ これでアプリの全機能をデモデータで試すことができます！")

if __name__ == '__main__':
    print("=" * 60)
    print("デモデータベース作成スクリプト")
    print("=" * 60)
    print("\nこのスクリプトは以下を実行します：")
    print("1. 現在のデータベースをバックアップ（既存の場合）")
    print("2. 新しいデモデータベースを作成")
    print("3. 10種類の原料と複数のロットを登録")
    print("4. デモデータベースをbackupsフォルダに保存")
    print()
    
    # 自動実行（デモ用）
    create_demo_database()
