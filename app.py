from flask import Flask, render_template, request, redirect, url_for, make_response, jsonify, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_wtf import FlaskForm
from flask_wtf.csrf import CSRFProtect
from wtforms import StringField, FloatField, SubmitField, SelectField, DateField
from wtforms.validators import DataRequired, Email, Optional
from datetime import datetime, date
import csv
import io
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import shutil
import os
import json
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

# 設定ファイルのパス
CONFIG_FILE = 'config.json'

def load_config():
    """設定ファイルを読み込む"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_config(config):
    """設定ファイルに保存"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

def select_database_folder():
    """データベースフォルダを選択"""
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    folder = filedialog.askdirectory(
        title='データベースフォルダを選択してください',
        initialdir=os.getcwd()
    )
    
    root.destroy()
    return folder

def get_database_path():
    """データベースパスを取得または設定"""
    config = load_config()
    
    # 設定にデータベースフォルダがあるか確認
    if 'database_folder' in config and os.path.exists(config['database_folder']):
        db_folder = config['database_folder']
    else:
        # フォルダ選択ダイアログを表示
        root = Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        messagebox.showinfo(
            'データベースフォルダの選択',
            'データベースを保存するフォルダを選択してください。\n'
            '共有フォルダを指定すると、複数人で同じデータベースを使用できます。'
        )
        
        db_folder = select_database_folder()
        root.destroy()
        
        if not db_folder:
            messagebox.showerror('エラー', 'フォルダが選択されませんでした。\nデフォルトのinstanceフォルダを使用します。')
            db_folder = os.path.join(os.getcwd(), 'instance')
            os.makedirs(db_folder, exist_ok=True)
        
        # 設定を保存
        config['database_folder'] = db_folder
        save_config(config)
    
    # データベースファイルのパスを返す
    db_path = os.path.join(db_folder, 'inventory.db')
    return db_path

# Flaskアプリケーションの初期化
app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'

# データベースパスを取得
db_path = get_database_path()
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{db_path}'

db = SQLAlchemy(app)
csrf = CSRFProtect(app)

class RawMaterial(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    weight = db.Column(db.Float, nullable=False)  # 原料全体の重量（表示用）
    unit = db.Column(db.String(20), default='g')  # 単位はg固定
    min_weight = db.Column(db.Float, default=0.0)
    email = db.Column(db.String(120), nullable=True)  # 購入担当者メール
    excel_path = db.Column(db.String(500), nullable=True)  # エクセルファイルパス
    action_type = db.Column(db.String(20), default='none')  # 'email', 'excel', 'none'

    def get_total_lot_weight(self):
        """全ロットの現在重量の合計"""
        return sum(lot.weight for lot in self.lots)
    
    def get_predicted_stock(self):
        """現在量 + 未実行の補充予約 - 未実行の使用予約 = 予測在庫量（ロットの合計）"""
        total_current = self.get_total_lot_weight()
        # 実行済みの予約はカウントしない
        replenish = sum(r.quantity for r in self.reservations if r.type == 'replenish' and not r.executed)
        use = sum(r.quantity for r in self.reservations if r.type == 'use' and not r.executed)
        return total_current + replenish - use
    
    def is_low_stock_alert(self):
        """予測在庫が最低量を下回るかチェック（途中の期間も含む）"""
        # 途中で最低量を下回る期間がある場合もアラートとする
        critical_periods = self.get_critical_periods()
        return len(critical_periods) > 0

    def get_critical_periods(self):
        """最低重量を下回る期間を計算"""
        from datetime import datetime, timedelta
        
        # 現在の在庫量
        current_stock = self.get_total_lot_weight()
        
        # 未実行の予約を日付順に取得
        reservations = sorted(
            [r for r in self.reservations if not r.executed and r.scheduled_date],
            key=lambda x: x.scheduled_date
        )
        
        if not reservations:
            # 予約がない場合、現在の在庫が最低重量を下回っているかチェック
            if current_stock < self.min_weight:
                return [{
                    'start_date': datetime.now().date(),
                    'end_date': None,
                    'min_stock': current_stock,
                    'shortage': self.min_weight - current_stock
                }]
            return []
        
        critical_periods = []
        running_stock = current_stock
        period_start = None
        period_start_date = None
        min_stock_in_period = running_stock
        
        # 現在の在庫が既に不足している場合
        if running_stock < self.min_weight:
            period_start = True
            period_start_date = datetime.now().date()
            min_stock_in_period = running_stock
        
        # 各予約を時系列で処理
        for reservation in reservations:
            # 予約実行前の在庫状態をチェック
            prev_stock = running_stock
            
            # 予約を実行
            if reservation.type == 'use':
                running_stock -= reservation.quantity
            else:  # replenish
                running_stock += reservation.quantity
            
            # 使用予約で最低重量を下回った場合、期間開始
            if reservation.type == 'use' and prev_stock >= self.min_weight and running_stock < self.min_weight:
                period_start = True
                period_start_date = reservation.scheduled_date
                min_stock_in_period = running_stock
            
            # 既に期間中で、さらに在庫が減少
            elif period_start and running_stock < self.min_weight:
                min_stock_in_period = min(min_stock_in_period, running_stock)
            
            # 補充予約で最低重量を上回った場合、期間終了
            if reservation.type == 'replenish' and period_start and running_stock >= self.min_weight:
                critical_periods.append({
                    'start_date': period_start_date,
                    'end_date': reservation.scheduled_date,
                    'min_stock': min_stock_in_period,
                    'shortage': self.min_weight - min_stock_in_period
                })
                period_start = False
                period_start_date = None
                min_stock_in_period = running_stock
        
        # 最後の期間が終了していない場合
        if period_start:
            critical_periods.append({
                'start_date': period_start_date,
                'end_date': None,  # 終了日未定（補充予約が必要）
                'min_stock': min_stock_in_period,
                'shortage': self.min_weight - min_stock_in_period
            })
        
        return critical_periods

    def get_usage_stats(self, period_days):
        """指定期間の使用量・補充量を集計"""
        from datetime import datetime, timedelta
        
        end_date = datetime.now()
        start_date = end_date - timedelta(days=period_days)
        
        # 実行済みの予約のみを対象（実行日時が記録されているもの）
        executed_reservations = [r for r in self.reservations if r.executed and r.executed_date]
        
        # 期間内の予約をフィルタ
        period_reservations = [
            r for r in executed_reservations 
            if start_date <= r.executed_date <= end_date
        ]
        
        # 使用量と補充量を集計
        total_used = sum(r.actual_quantity or r.quantity for r in period_reservations if r.type == 'use')
        total_replenished = sum(r.actual_quantity or r.quantity for r in period_reservations if r.type == 'replenish')
        
        # 日別データ
        daily_data = {}
        for r in period_reservations:
            date_key = r.executed_date.strftime('%Y-%m-%d')
            if date_key not in daily_data:
                daily_data[date_key] = {'used': 0, 'replenished': 0}
            
            if r.type == 'use':
                daily_data[date_key]['used'] += r.actual_quantity or r.quantity
            else:
                daily_data[date_key]['replenished'] += r.actual_quantity or r.quantity
        
        return {
            'period_days': period_days,
            'start_date': start_date.strftime('%Y-%m-%d'),
            'end_date': end_date.strftime('%Y-%m-%d'),
            'total_used': round(total_used, 3),
            'total_replenished': round(total_replenished, 3),
            'net_change': round(total_replenished - total_used, 3),
            'daily_data': daily_data,
            'transaction_count': len(period_reservations)
        }

    def __repr__(self):
        return f'<RawMaterial {self.name}>'

class Lot(db.Model):
    """ロット（原料の下位管理単位）"""
    id = db.Column(db.Integer, primary_key=True)
    material_id = db.Column(db.Integer, db.ForeignKey('raw_material.id'), nullable=False)
    lot_name = db.Column(db.String(100), nullable=False)  # ロット名
    weight = db.Column(db.Float, nullable=False)  # ロットの重量
    date_created = db.Column(db.DateTime, default=db.func.current_timestamp())

    material = db.relationship('RawMaterial', backref=db.backref('lots', lazy=True, cascade='all, delete-orphan'))

    def __repr__(self):
        return f'<Lot {self.lot_name}>'

class Reservation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    material_id = db.Column(db.Integer, db.ForeignKey('raw_material.id'), nullable=False)
    lot_id = db.Column(db.Integer, db.ForeignKey('lot.id'), nullable=True)  # 既存ロット指定（オプショナル）
    lot_name = db.Column(db.String(100), nullable=True)  # 新規ロット名（オプショナル）
    recipe_id = db.Column(db.Integer, db.ForeignKey('recipe.id'), nullable=True)  # レシピからの予約
    type = db.Column(db.String(20), nullable=False)  # 'use' or 'replenish'
    quantity = db.Column(db.Float, nullable=False)  # 予約量
    actual_quantity = db.Column(db.Float, nullable=True)  # 実際の量
    user_name = db.Column(db.String(100), nullable=True)  # 使用者名
    purpose = db.Column(db.String(200), nullable=True)  # 目的
    scheduled_date = db.Column(db.Date, nullable=True)  # 予定日
    date = db.Column(db.DateTime, default=db.func.current_timestamp())  # 登録日
    executed = db.Column(db.Boolean, default=False)  # 実行済みかどうか
    executed_date = db.Column(db.DateTime, nullable=True)  # 実行日時

    material = db.relationship('RawMaterial', backref=db.backref('reservations', lazy=True, cascade='all, delete-orphan'))
    lot = db.relationship('Lot', backref=db.backref('reservations', lazy=True, cascade='all, delete-orphan'))
    recipe = db.relationship('Recipe', backref=db.backref('reservations', lazy=True))

    def is_overdue(self):
        """期限切れかどうかをチェック"""
        if not self.scheduled_date or self.executed:
            return False
        return self.scheduled_date < date.today()

    def __repr__(self):
        return f'<Reservation {self.type} {self.quantity}>'

class Recipe(db.Model):
    """複数原料の組み合わせ（レシピ）"""
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)  # レシピ名
    description = db.Column(db.String(200), nullable=True)  # 説明
    type = db.Column(db.String(20), nullable=False)  # 'use' or 'replenish'
    date_created = db.Column(db.DateTime, default=db.func.current_timestamp())

    def __repr__(self):
        return f'<Recipe {self.name}>'

class RecipeItem(db.Model):
    """レシピの各原料と量"""
    id = db.Column(db.Integer, primary_key=True)
    recipe_id = db.Column(db.Integer, db.ForeignKey('recipe.id'), nullable=False)
    material_id = db.Column(db.Integer, db.ForeignKey('raw_material.id'), nullable=False)
    quantity = db.Column(db.Float, nullable=False)
    lot_name = db.Column(db.String(100), nullable=True)  # ロット名（オプショナル）

    recipe = db.relationship('Recipe', backref=db.backref('items', lazy=True, cascade='all, delete-orphan'))
    material = db.relationship('RawMaterial')

    def __repr__(self):
        return f'<RecipeItem {self.material.name} {self.quantity}>'

class MaterialForm(FlaskForm):
    name = StringField('Name', validators=[DataRequired()])
    weight = FloatField('Weight (g)', validators=[Optional()], default=0.0)
    min_weight = FloatField('Min Weight (g)', default=0.0)
    action_type = SelectField('Action Type', choices=[('none', 'なにもしない'), ('email', 'メール連絡'), ('excel', 'エクセルを開く')], default='none')
    email = StringField('Purchase Email', validators=[Optional(), Email()])
    excel_path = StringField('Excel File Path', validators=[Optional()])
    submit = SubmitField('Submit')

class LotForm(FlaskForm):
    lot_name = StringField('Lot Name', validators=[DataRequired()])
    weight = FloatField('Weight', validators=[DataRequired()])
    submit = SubmitField('Submit')

class ReservationForm(FlaskForm):
    lot_id = SelectField('Existing Lot (Optional)', coerce=int, validators=[Optional()])
    lot_name = StringField('New Lot Name (Optional)', validators=[Optional()])
    quantity = FloatField('Quantity', validators=[DataRequired()])

    user_name = StringField('User Name (Optional)', validators=[Optional()])
    purpose = StringField('Purpose (Optional)', validators=[Optional()])
    scheduled_date = DateField('Scheduled Date (Optional)', format='%Y-%m-%d', validators=[Optional()])
    submit = SubmitField('Reserve')

class RecipeForm(FlaskForm):
    name = StringField('Recipe Name', validators=[DataRequired()])
    description = StringField('Description (Optional)', validators=[Optional()])
    submit = SubmitField('Save Recipe')

@app.route('/')
def index():
    search = request.args.get('search', '')
    sort_by = request.args.get('sort_by', 'name')
    materials = RawMaterial.query
    if search:
        materials = materials.filter(RawMaterial.name.contains(search))
    if sort_by == 'name':
        materials = materials.order_by(RawMaterial.name)
    elif sort_by == 'weight':
        # ソート用: SQLでは直接計算できないため、Pythonでソート
        materials = materials.all()
        materials = sorted(materials, key=lambda m: m.get_total_lot_weight())
        return render_template('index.html', materials=materials, search=search, sort_by=sort_by)
    materials = materials.all()
    return render_template('index.html', materials=materials, search=search, sort_by=sort_by)

@app.route('/add', methods=['GET', 'POST'])
def add():
    form = MaterialForm()
    if form.validate_on_submit():
        material = RawMaterial(
            name=form.name.data, 
            weight=form.weight.data if form.weight.data is not None else 0.0, 
            unit='g',  # g固定
            min_weight=form.min_weight.data,
            email=form.email.data,
            excel_path=form.excel_path.data,
            action_type=form.action_type.data
        )
        db.session.add(material)
        db.session.commit()
        return redirect(url_for('index'))
    return render_template('add.html', form=form)

@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit(id):
    material = RawMaterial.query.get_or_404(id)
    form = MaterialForm()
    if form.validate_on_submit():
        material.name = form.name.data
        material.weight = form.weight.data
        material.unit = 'g'  # g固定
        material.min_weight = form.min_weight.data
        material.email = form.email.data
        material.excel_path = form.excel_path.data
        material.action_type = form.action_type.data
        db.session.commit()
        return redirect(url_for('index'))
    elif request.method == 'GET':
        form.name.data = material.name
        form.weight.data = material.weight
        form.min_weight.data = material.min_weight
        form.email.data = material.email
        form.excel_path.data = material.excel_path
        form.action_type.data = material.action_type
    return render_template('edit.html', form=form)

@app.route('/material_stats/<int:id>')
def material_stats(id):
    """原料の統計ページ"""
    material = RawMaterial.query.get_or_404(id)
    return render_template('material_stats.html', material=material)

@app.route('/delete/<int:id>')
def delete(id):
    material = RawMaterial.query.get_or_404(id)
    
    try:
        # 関連する予約を先に削除
        Reservation.query.filter_by(material_id=id).delete()
        
        # 関連するロットを削除（ロットに紐づく予約もカスケード削除される）
        Lot.query.filter_by(material_id=id).delete()
        
        # 原料を削除
        db.session.delete(material)
        db.session.commit()
        flash(f'原料「{material.name}」と関連データを削除しました', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'削除に失敗しました: {str(e)}', 'danger')
    
    return redirect(url_for('index'))

@app.route('/reserve_use/<int:id>', methods=['GET', 'POST'])
def reserve_use(id):
    material = RawMaterial.query.get_or_404(id)
    form = ReservationForm()
    # ロット選択肢を追加（空欄も含む）
    form.lot_id.choices = [(0, '既存ロットから選択しない')] + [(lot.id, lot.lot_name) for lot in material.lots]
    if form.validate_on_submit():
        lot_id = form.lot_id.data if form.lot_id.data != 0 else None
        
        # 予約後の予測重量を計算
        predicted_after_reserve = material.get_predicted_stock() - form.quantity.data
        
        # 最低重量を下回る場合は警告
        if predicted_after_reserve < material.min_weight:
            shortage = material.min_weight - predicted_after_reserve
            warning_msg = f'⚠️ 警告: この予約により予測在庫が最低量を{shortage:.1f}g下回ります（予測: {predicted_after_reserve:.1f}g / 最低: {material.min_weight:.1f}g）'
            
            # アクションタイプに応じた対処を促す
            if material.action_type == 'email' and material.email:
                warning_msg += f' → 購入担当者（{material.email}）にメール連絡してください'
            elif material.action_type == 'excel' and material.excel_path:
                warning_msg += f' → <a href="/open_excel/{material.id}" class="alert-link">発注用エクセルを開く</a>'
            
            flash(warning_msg, 'warning')
        
        reservation = Reservation(
            material_id=id, 
            lot_id=lot_id,
            lot_name=form.lot_name.data if form.lot_name.data else None,
            type='use', 
            quantity=form.quantity.data,
            user_name=form.user_name.data,
            purpose=form.purpose.data,
            scheduled_date=form.scheduled_date.data
        )
        db.session.add(reservation)
        db.session.commit()
        flash('使用予約を登録しました', 'success')
        return redirect(url_for('index'))
    return render_template('reserve.html', form=form, material=material, action='use')

@app.route('/reserve_replenish/<int:id>', methods=['GET', 'POST'])
def reserve_replenish(id):
    material = RawMaterial.query.get_or_404(id)
    form = ReservationForm()
    # ロット選択肢を追加（空欄も含む）
    form.lot_id.choices = [(0, '既存ロットから選択しない')] + [(lot.id, lot.lot_name) for lot in material.lots]
    if form.validate_on_submit():
        lot_id = form.lot_id.data if form.lot_id.data != 0 else None
        reservation = Reservation(
            material_id=id, 
            lot_id=lot_id,
            lot_name=form.lot_name.data if form.lot_name.data else None,
            type='replenish', 
            quantity=form.quantity.data,
            user_name=form.user_name.data,
            purpose=form.purpose.data,
            scheduled_date=form.scheduled_date.data
        )
        db.session.add(reservation)
        db.session.commit()
        flash('補充予約を登録しました', 'success')
        return redirect(url_for('index'))
    return render_template('reserve.html', form=form, material=material, action='replenish')

@app.route('/lots/<int:material_id>')
def lots(material_id):
    """原料のロット一覧"""
    material = RawMaterial.query.get_or_404(material_id)
    return render_template('lots.html', material=material)

@app.route('/add_lot/<int:material_id>', methods=['GET', 'POST'])
def add_lot(material_id):
    """ロット追加"""
    material = RawMaterial.query.get_or_404(material_id)
    form = LotForm()
    if form.validate_on_submit():
        lot = Lot(material_id=material_id, lot_name=form.lot_name.data, weight=form.weight.data)
        db.session.add(lot)
        db.session.flush()  # lotのIDを取得するため
        
        # 統計用に実行済み予約を作成（ロット追加は補充扱い）
        if form.weight.data > 0:
            auto_reservation = Reservation(
                material_id=material_id,
                lot_id=lot.id,
                lot_name=form.lot_name.data,
                type='replenish',
                quantity=form.weight.data,
                actual_quantity=form.weight.data,
                user_name='システム',
                purpose=f'ロット追加（{form.lot_name.data}）',
                scheduled_date=datetime.now(),
                executed=True,
                executed_date=datetime.now()
            )
            db.session.add(auto_reservation)
        
        db.session.commit()
        flash(f'ロット「{form.lot_name.data}」を追加しました', 'success')
        return redirect(url_for('lots', material_id=material_id))
    return render_template('add_lot.html', form=form, material=material)

@app.route('/edit_lot/<int:id>', methods=['GET', 'POST'])
def edit_lot(id):
    """ロット編集"""
    lot = Lot.query.get_or_404(id)
    form = LotForm()
    if form.validate_on_submit():
        old_weight = lot.weight
        new_weight = form.weight.data
        lot.lot_name = form.lot_name.data
        lot.weight = new_weight
        
        # 重量が変化した場合、統計用に実行済み予約を作成
        weight_diff = new_weight - old_weight
        if weight_diff != 0:
            transaction_type = 'replenish' if weight_diff > 0 else 'use'
            quantity = abs(weight_diff)
            
            auto_reservation = Reservation(
                material_id=lot.material_id,
                lot_id=lot.id,
                lot_name=lot.lot_name,
                type=transaction_type,
                quantity=quantity,
                actual_quantity=quantity,
                user_name='システム',
                purpose=f'ロット直接編集（{lot.lot_name}）',
                scheduled_date=datetime.now(),
                executed=True,
                executed_date=datetime.now()
            )
            db.session.add(auto_reservation)
        
        db.session.commit()
        flash(f'ロット「{lot.lot_name}」を更新しました', 'success')
        return redirect(url_for('lots', material_id=lot.material_id))
    elif request.method == 'GET':
        form.lot_name.data = lot.lot_name
        form.weight.data = lot.weight
    return render_template('edit_lot.html', form=form, lot=lot)

@app.route('/delete_lot/<int:id>')
def delete_lot(id):
    """ロット削除"""
    lot = Lot.query.get_or_404(id)
    material_id = lot.material_id
    lot_name = lot.lot_name
    lot_weight = lot.weight
    
    # 統計用に実行済み予約を作成（ロット削除は使用扱い）
    if lot_weight > 0:
        auto_reservation = Reservation(
            material_id=material_id,
            lot_id=None,  # ロット削除後なのでNone
            lot_name=lot_name,
            type='use',
            quantity=lot_weight,
            actual_quantity=lot_weight,
            user_name='システム',
            purpose=f'ロット削除（{lot_name}）',
            scheduled_date=datetime.now(),
            executed=True,
            executed_date=datetime.now()
        )
        db.session.add(auto_reservation)
    
    db.session.delete(lot)
    db.session.commit()
    flash(f'ロット「{lot_name}」を削除しました', 'success')
    return redirect(url_for('lots', material_id=material_id))

@app.route('/reservations')
def reservations():
    """予約管理ページ"""
    use_reservations = Reservation.query.filter_by(type='use', executed=False).order_by(Reservation.scheduled_date.asc(), Reservation.date.desc()).all()
    replenish_reservations = Reservation.query.filter_by(type='replenish', executed=False).order_by(Reservation.scheduled_date.asc(), Reservation.date.desc()).all()
    
    # 期限切れ予約を抽出
    overdue_reservations = [r for r in use_reservations + replenish_reservations if r.is_overdue()]
    
    # レシピ予約をグループ化
    recipe_groups = {}
    for r in use_reservations:
        if r.recipe_id:
            if r.recipe_id not in recipe_groups:
                recipe_groups[r.recipe_id] = {
                    'recipe': r.recipe,
                    'scheduled_date': r.scheduled_date,
                    'user_name': r.user_name,
                    'purpose': r.purpose,
                    'reservations': []
                }
            recipe_groups[r.recipe_id]['reservations'].append(r)
    
    return render_template('reservations.html', 
                         use_reservations=use_reservations,
                         replenish_reservations=replenish_reservations,
                         overdue_count=len(overdue_reservations),
                         recipe_groups=recipe_groups)

@app.route('/execute_reservation/<int:id>', methods=['GET', 'POST'])
def execute_reservation(id):
    """予約を実行して在庫に反映"""
    reservation = Reservation.query.get_or_404(id)
    material = reservation.material
    
    # POSTリクエストの場合、実際の量とロット情報を取得
    if request.method == 'POST':
        actual_quantity = float(request.form.get('actual_quantity', reservation.quantity))
        reservation.actual_quantity = actual_quantity
        quantity_to_use = actual_quantity
        
        # 使用予約の場合はロット選択を取得
        if reservation.type == 'use':
            lot_id = request.form.get('lot_id')
            if lot_id:
                reservation.lot_id = int(lot_id)
        # 補充予約の場合はロット名を取得
        elif reservation.type == 'replenish':
            lot_name = request.form.get('lot_name', '').strip()
            if lot_name:
                reservation.lot_name = lot_name
    else:
        # GETリクエストの場合は予約量を使用（後方互換性のため）
        quantity_to_use = reservation.quantity
        if reservation.actual_quantity:
            quantity_to_use = reservation.actual_quantity
    
    try:
        if reservation.type == 'use':
            # 使用予約の実行
            if not reservation.lot_id:
                flash('エラー: ロットを選択してください', 'danger')
                return redirect(url_for('reservations'))
            
            lot = reservation.lot
            if lot.weight >= quantity_to_use:
                lot.weight -= quantity_to_use
            else:
                flash(f'エラー: ロット「{lot.lot_name}」の在庫が不足しています', 'danger')
                return redirect(url_for('reservations'))
        
        elif reservation.type == 'replenish':
            # 補充予約の実行
            if not reservation.lot_name:
                flash('エラー: ロット名を入力してください', 'danger')
                return redirect(url_for('reservations'))
            
            # 新規ロット名が指定されている場合
            existing_lot = Lot.query.filter_by(material_id=material.id, lot_name=reservation.lot_name).first()
            if existing_lot:
                # 既存ロットに追加
                existing_lot.weight += quantity_to_use
            else:
                # 新規ロット作成
                new_lot = Lot(material_id=material.id, lot_name=reservation.lot_name, weight=quantity_to_use)
                db.session.add(new_lot)
        
        # 予約を実行済みにマーク
        reservation.executed = True
        reservation.executed_date = datetime.now()
        db.session.commit()
        flash(f'予約を実行しました: {material.name} ({quantity_to_use} {material.unit})', 'success')
    
    except Exception as e:
        db.session.rollback()
        flash(f'エラーが発生しました: {str(e)}', 'danger')
    
    return redirect(url_for('reservations'))

@app.route('/execute_recipe/<int:recipe_id>', methods=['POST'])
def execute_recipe(recipe_id):
    """レシピ予約を一括実行"""
    recipe = Recipe.query.get_or_404(recipe_id)
    
    # このレシピに紐づく未実行の使用予約を取得
    reservations = Reservation.query.filter_by(
        recipe_id=recipe_id,
        type='use',
        executed=False
    ).all()
    
    if not reservations:
        flash('実行する予約が見つかりません', 'warning')
        return redirect(url_for('reservations'))
    
    try:
        # 各原料の実績値を取得して実行
        for reservation in reservations:
            material = reservation.material
            actual_quantity_key = f'actual_quantity_{reservation.id}'
            actual_quantity = float(request.form.get(actual_quantity_key, reservation.quantity))
            
            # ロット選択を取得
            lot_id_key = f'lot_id_{reservation.id}'
            lot_id = request.form.get(lot_id_key)
            
            if not lot_id:
                flash(f'エラー: {material.name}のロットを選択してください', 'danger')
                db.session.rollback()
                return redirect(url_for('reservations'))
            
            reservation.actual_quantity = actual_quantity
            reservation.lot_id = int(lot_id)
            
            # 在庫を減少
            lot = Lot.query.get(reservation.lot_id)
            if not lot:
                flash(f'エラー: ロットが見つかりません', 'danger')
                db.session.rollback()
                return redirect(url_for('reservations'))
            
            if lot.weight >= actual_quantity:
                lot.weight -= actual_quantity
            else:
                flash(f'エラー: ロット「{lot.lot_name}」の在庫が不足しています', 'danger')
                db.session.rollback()
                return redirect(url_for('reservations'))
            
            # 予約を実行済みにマーク
            reservation.executed = True
            reservation.executed_date = datetime.now()
        
        db.session.commit()
        flash(f'レシピ「{recipe.name}」の予約を一括実行しました', 'success')
    
    except Exception as e:
        db.session.rollback()
        flash(f'エラーが発生しました: {str(e)}', 'danger')
    
    return redirect(url_for('reservations'))

@app.route('/edit_reservation/<int:id>', methods=['GET', 'POST'])
def edit_reservation(id):
    """予約編集"""
    reservation = Reservation.query.get_or_404(id)
    
    if request.method == 'POST':
        reservation.user_name = request.form.get('user_name', '')
        reservation.purpose = request.form.get('purpose', '')
        reservation.quantity = float(request.form.get('quantity', reservation.quantity))
        
        # 予定日の処理
        scheduled_date_str = request.form.get('scheduled_date')
        if scheduled_date_str:
            reservation.scheduled_date = datetime.strptime(scheduled_date_str, '%Y-%m-%d').date()
        
        # ロット名/ロット選択の処理
        if reservation.type == 'use':
            lot_id = request.form.get('lot_id')
            if lot_id:
                reservation.lot_id = int(lot_id)
        else:  # replenish
            lot_name = request.form.get('lot_name', '').strip()
            if lot_name:
                reservation.lot_name = lot_name
        
        db.session.commit()
        flash('予約を更新しました', 'success')
        return redirect(url_for('reservations'))
    
    return render_template('edit_reservation.html', reservation=reservation)

@app.route('/delete_reservation/<int:id>')
def delete_reservation(id):
    """予約削除"""
    reservation = Reservation.query.get_or_404(id)
    db.session.delete(reservation)
    db.session.commit()
    flash('予約を削除しました', 'success')
    return redirect(url_for('reservations'))

@app.route('/export')
def export():
    materials = RawMaterial.query.all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['ID', '名前', '現在重量', '単位', '最低量', '予測在庫'])
    for material in materials:
        current_weight = material.get_total_lot_weight()
        predicted = material.get_predicted_stock()
        writer.writerow([material.id, material.name, round(current_weight, 2), material.unit, material.min_weight, round(predicted, 2)])
    output.seek(0)
    # BOM付きUTF-8でエンコードして日本語文字化けを防止
    csv_data = '\ufeff' + output.getvalue()
    response = make_response(csv_data)
    response.headers['Content-Disposition'] = 'attachment; filename=inventory.csv'
    response.headers['Content-type'] = 'text/csv; charset=utf-8-sig'
    return response

@app.route('/send_alert_email/<int:id>', methods=['POST'])
def send_alert_email(id):
    """アラートメールを送信"""
    material = RawMaterial.query.get_or_404(id)
    
    if not material.email:
        flash('購入担当者のメールアドレスが登録されていません。', 'warning')
        return redirect(url_for('index'))
    
    try:
        # メール内容
        current_weight = material.get_total_lot_weight()
        predicted_stock = material.get_predicted_stock()
        subject = f"【在庫アラート】{material.name}の補充が必要です"
        body = f"""
在庫管理システムからの自動通知

原料名: {material.name}
現在量: {current_weight:.2f} {material.unit}
最低量: {material.min_weight} {material.unit}
予測在庫量: {predicted_stock:.2f} {material.unit}

予測在庫量が最低量を下回る見込みです。
至急、補充の手配をお願いします。

※このメールは在庫管理システムから自動送信されています。
        """
        
        # 実際のメール送信（Gmail使用例）
        # 注意: 本番環境では環境変数やconfigファイルで設定してください
        sender_email = "your-email@gmail.com"  # 送信元メール
        sender_password = "your-app-password"  # アプリパスワード
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = material.email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Gmail SMTPサーバー経由で送信（実際の送信を有効にする場合はコメント解除）
        # server = smtplib.SMTP('smtp.gmail.com', 587)
        # server.starttls()
        # server.login(sender_email, sender_password)
        # server.send_message(msg)
        # server.quit()
        
        # デモ用: 実際には送信せずにメッセージのみ表示
        flash(f'アラートメールを {material.email} に送信しました（デモモード）', 'success')
        
    except Exception as e:
        flash(f'メール送信に失敗しました: {str(e)}', 'danger')
    
    return redirect(url_for('index'))

@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')

@app.route('/api/stats')
def api_stats():
    materials = RawMaterial.query.all()
    
    # 総在庫数
    total_materials = len(materials)
    
    # 低在庫アラート数（予測在庫で判定）
    low_stock_count = sum(1 for m in materials if m.is_low_stock_alert())
    
    # アラート一覧
    alert_materials = []
    for material in materials:
        if material.is_low_stock_alert():
            total_weight = material.get_total_lot_weight()
            predicted = material.get_predicted_stock()
            critical_periods = material.get_critical_periods()
            
            # 日付をJSON互換形式に変換
            serialized_periods = []
            for period in critical_periods:
                serialized_periods.append({
                    'start_date': period['start_date'].isoformat() if period['start_date'] else None,
                    'end_date': period['end_date'].isoformat() if period['end_date'] else None,
                    'min_stock': round(period['min_stock'], 2),
                    'shortage': round(period['shortage'], 2)
                })
            
            alert_materials.append({
                'id': material.id,
                'name': material.name,
                'current': round(total_weight, 2),
                'predicted': round(predicted, 2),
                'min_weight': material.min_weight,
                'unit': material.unit,
                'email': material.email,
                'excel_path': material.excel_path,
                'action_type': material.action_type,
                'critical_periods': serialized_periods
            })
    
    # 在庫状況データ
    materials_data = []
    for material in materials:
        total_weight = material.get_total_lot_weight()
        predicted_stock = material.get_predicted_stock()
        materials_data.append({
            'name': material.name,
            'current': round(total_weight, 2),
            'predicted': round(predicted_stock, 2),
            'min_weight': material.min_weight,
            'unit': material.unit
        })
    
    # 予約情報の集計
    use_reservations = Reservation.query.filter_by(type='use').order_by(Reservation.date.desc()).limit(5).all()
    replenish_reservations = Reservation.query.filter_by(type='replenish').order_by(Reservation.date.desc()).limit(5).all()
    
    use_list = [{
        'material': r.material.name,
        'lot': r.lot.lot_name if r.lot else '原料全体',
        'quantity': r.quantity,
        'date': r.date.strftime('%Y/%m/%d %H:%M') if r.date else 'N/A'
    } for r in use_reservations]
    
    replenish_list = [{
        'material': r.material.name,
        'lot': r.lot.lot_name if r.lot else '原料全体',
        'quantity': r.quantity,
        'date': r.date.strftime('%Y/%m/%d %H:%M') if r.date else 'N/A'
    } for r in replenish_reservations]
    
    # 期限切れ予約数を計算
    from datetime import date, timedelta
    today = date.today()
    all_reservations = Reservation.query.filter_by(executed=False).all()
    overdue_count = sum(1 for r in all_reservations if r.scheduled_date and r.scheduled_date < today)
    
    # 今週（7日以内）の予約数を計算
    week_later = today + timedelta(days=7)
    week_reservations = sum(1 for r in all_reservations if r.scheduled_date and today <= r.scheduled_date <= week_later)
    
    return jsonify({
        'total_materials': total_materials,
        'low_stock_count': low_stock_count,
        'alert_materials': alert_materials,
        'materials': materials_data,
        'use_reservations': use_list,
        'replenish_reservations': replenish_list,
        'overdue_count': overdue_count,
        'week_reservations': week_reservations
    })

@app.route('/api/material_stats/<int:id>')
def api_material_stats(id):
    """原料の期間別統計データを取得"""
    material = RawMaterial.query.get_or_404(id)
    
    # 各期間の統計を取得
    stats = {
        '1d': material.get_usage_stats(1),
        '7d': material.get_usage_stats(7),
        '1m': material.get_usage_stats(30),
        '3m': material.get_usage_stats(90),
        '6m': material.get_usage_stats(180),
        '1y': material.get_usage_stats(365)
    }
    
    return jsonify({
        'material_id': material.id,
        'material_name': material.name,
        'current_stock': material.get_total_lot_weight(),
        'unit': material.unit,
        'stats': stats
    })

# Recipe Management Routes
@app.route('/recipes')
def recipes():
    recipes = Recipe.query.order_by(Recipe.date_created.desc()).all()
    return render_template('recipes.html', recipes=recipes)

@app.route('/add_recipe', methods=['GET', 'POST'])
def add_recipe():
    form = RecipeForm()
    materials = RawMaterial.query.all()
    if form.validate_on_submit():
        recipe = Recipe(
            name=form.name.data,
            description=form.description.data,
            type='use'
        )
        db.session.add(recipe)
        db.session.flush()  # Get recipe.id before adding items
        
        # Add recipe items from form data
        for material in materials:
            quantity_key = f'material_{material.id}_quantity'
            lot_name_key = f'material_{material.id}_lot_name'
            quantity = request.form.get(quantity_key, type=float)
            lot_name = request.form.get(lot_name_key, '')
            
            if quantity and quantity > 0:
                recipe_item = RecipeItem(
                    recipe_id=recipe.id,
                    material_id=material.id,
                    quantity=quantity,
                    lot_name=lot_name if lot_name else None
                )
                db.session.add(recipe_item)
        
        db.session.commit()
        flash('レシピを登録しました', 'success')
        return redirect(url_for('recipes'))
    
    return render_template('add_recipe.html', form=form, materials=materials)

@app.route('/edit_recipe/<int:id>', methods=['GET', 'POST'])
def edit_recipe(id):
    recipe = Recipe.query.get_or_404(id)
    form = RecipeForm()
    materials = RawMaterial.query.all()
    
    if form.validate_on_submit():
        recipe.name = form.name.data
        recipe.description = form.description.data
        
        # Delete existing recipe items
        RecipeItem.query.filter_by(recipe_id=recipe.id).delete()
        
        # Add new recipe items
        for material in materials:
            quantity_key = f'material_{material.id}_quantity'
            lot_name_key = f'material_{material.id}_lot_name'
            quantity = request.form.get(quantity_key, type=float)
            lot_name = request.form.get(lot_name_key, '')
            
            if quantity and quantity > 0:
                recipe_item = RecipeItem(
                    recipe_id=recipe.id,
                    material_id=material.id,
                    quantity=quantity,
                    lot_name=lot_name if lot_name else None
                )
                db.session.add(recipe_item)
        
        db.session.commit()
        flash('レシピを更新しました', 'success')
        return redirect(url_for('recipes'))
    
    if request.method == 'GET':
        form.name.data = recipe.name
        form.description.data = recipe.description
    
    return render_template('edit_recipe.html', form=form, recipe=recipe, materials=materials)

@app.route('/delete_recipe/<int:id>')
def delete_recipe(id):
    recipe = Recipe.query.get_or_404(id)
    RecipeItem.query.filter_by(recipe_id=recipe.id).delete()
    db.session.delete(recipe)
    db.session.commit()
    flash('レシピを削除しました', 'success')
    return redirect(url_for('recipes'))

@app.route('/use_recipe/<int:recipe_id>', methods=['POST'])
@csrf.exempt
def use_recipe(recipe_id):
    recipe = Recipe.query.get_or_404(recipe_id)
    user_name = request.form.get('user_name', '')
    purpose = request.form.get('purpose', '')
    scheduled_date_str = request.form.get('scheduled_date', '')
    scheduled_date = datetime.strptime(scheduled_date_str, '%Y-%m-%d').date() if scheduled_date_str else None
    
    # Create reservations for each item in the recipe
    for item in recipe.items:
        reservation = Reservation(
            material_id=item.material_id,
            recipe_id=recipe.id,
            lot_name=item.lot_name,
            type='use',
            quantity=item.quantity,
            user_name=user_name,
            purpose=purpose,
            scheduled_date=scheduled_date
        )
        db.session.add(reservation)
    
    db.session.commit()
    flash(f'レシピ「{recipe.name}」から使用予約を作成しました', 'success')
    return redirect(url_for('reservations'))

# Backup Management Routes
def get_backup_folder():
    """バックアップフォルダのパスを取得"""
    config = load_config()
    if 'database_folder' in config:
        return os.path.join(config['database_folder'], 'backups')
    return 'backups'

def get_db_path():
    """データベースファイルのパスを取得"""
    return app.config['SQLALCHEMY_DATABASE_URI'].replace('sqlite:///', '')

def ensure_backup_folder():
    """バックアップフォルダの存在を確認し、なければ作成"""
    backup_folder = get_backup_folder()
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)

@app.route('/backup')
def backup_management():
    """バックアップ管理ページ"""
    ensure_backup_folder()
    backups = []
    backup_folder = get_backup_folder()
    
    if os.path.exists(backup_folder):
        for filename in os.listdir(backup_folder):
            if filename.endswith('.db'):
                filepath = os.path.join(backup_folder, filename)
                stat = os.stat(filepath)
                backups.append({
                    'filename': filename,
                    'size': stat.st_size / 1024,  # KB
                    'created': datetime.fromtimestamp(stat.st_mtime).strftime('%Y/%m/%d %H:%M:%S')
                })
    
    backups.sort(key=lambda x: x['created'], reverse=True)
    return render_template('backup.html', backups=backups)

@app.route('/backup/create', methods=['POST'])
def create_backup():
    """新規バックアップを作成"""
    try:
        ensure_backup_folder()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_filename = f'inventory_backup_{timestamp}.db'
        backup_folder = get_backup_folder()
        backup_path = os.path.join(backup_folder, backup_filename)
        db_path = get_db_path()
        
        if os.path.exists(db_path):
            shutil.copy2(db_path, backup_path)
            flash(f'バックアップを作成しました: {backup_filename}', 'success')
        else:
            flash('データベースファイルが見つかりません', 'danger')
    except Exception as e:
        flash(f'バックアップの作成に失敗しました: {str(e)}', 'danger')
    
    return redirect(url_for('backup_management'))

@app.route('/backup/restore/<filename>', methods=['POST'])
def restore_backup(filename):
    """バックアップから復元"""
    try:
        backup_folder = get_backup_folder()
        backup_path = os.path.join(backup_folder, filename)
        db_path = get_db_path()
        
        if not os.path.exists(backup_path):
            flash('指定されたバックアップファイルが見つかりません', 'danger')
            return redirect(url_for('backup_management'))
        
        # 現在のDBをバックアップ（復元前の安全策）
        if os.path.exists(db_path):
            db_folder = os.path.dirname(db_path)
            safety_backup = os.path.join(db_folder, f'inventory_before_restore_{datetime.now().strftime("%Y%m%d_%H%M%S")}.db')
            shutil.copy2(db_path, safety_backup)
        
        # バックアップから復元
        shutil.copy2(backup_path, db_path)
        flash(f'バックアップから復元しました: {filename}', 'success')
    except Exception as e:
        flash(f'復元に失敗しました: {str(e)}', 'danger')
    
    return redirect(url_for('backup_management'))

@app.route('/backup/download/<filename>')
def download_backup(filename):
    """バックアップファイルをダウンロード"""
    try:
        backup_folder = get_backup_folder()
        backup_path = os.path.join(backup_folder, filename)
        if os.path.exists(backup_path):
            return send_file(backup_path, as_attachment=True, download_name=filename)
        else:
            flash('指定されたバックアップファイルが見つかりません', 'danger')
            return redirect(url_for('backup_management'))
    except Exception as e:
        flash(f'ダウンロードに失敗しました: {str(e)}', 'danger')
        return redirect(url_for('backup_management'))

@app.route('/backup/delete/<filename>', methods=['POST'])
def delete_backup(filename):
    """バックアップファイルを削除"""
    try:
        backup_folder = get_backup_folder()
        backup_path = os.path.join(backup_folder, filename)
        if os.path.exists(backup_path):
            os.remove(backup_path)
            flash(f'バックアップを削除しました: {filename}', 'success')
        else:
            flash('指定されたバックアップファイルが見つかりません', 'danger')
    except Exception as e:
        flash(f'削除に失敗しました: {str(e)}', 'danger')
    
    return redirect(url_for('backup_management'))

@app.route('/open_excel/<int:id>')
def open_excel(id):
    """指定された原料のエクセルファイルを開く"""
    material = RawMaterial.query.get_or_404(id)
    
    if not material.excel_path:
        flash('エクセルファイルパスが登録されていません', 'warning')
        return redirect(url_for('dashboard'))
    
    try:
        import subprocess
        import platform
        
        # ファイルの存在確認
        if not os.path.exists(material.excel_path):
            flash(f'ファイルが見つかりません: {material.excel_path}', 'danger')
            return redirect(url_for('dashboard'))
        
        # OSに応じてファイルを開く
        if platform.system() == 'Windows':
            os.startfile(material.excel_path)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(['open', material.excel_path])
        else:  # Linux
            subprocess.call(['xdg-open', material.excel_path])
        
        flash(f'{material.name}のエクセルファイルを開きました', 'success')
    except Exception as e:
        flash(f'ファイルを開けませんでした: {str(e)}', 'danger')
    
    return redirect(url_for('dashboard'))

@app.route('/settings')
def settings():
    """設定画面"""
    config = load_config()
    current_db_folder = config.get('database_folder', 'デフォルト（instance）')
    current_db_path = app.config['SQLALCHEMY_DATABASE_URI'].replace('sqlite:///', '')
    
    return render_template('settings.html', 
                         db_folder=current_db_folder,
                         db_path=current_db_path)

@app.route('/change_database_folder', methods=['POST'])
def change_database_folder():
    """データベースフォルダを変更"""
    try:
        from tkinter import Tk, filedialog, messagebox
        
        root = Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        folder = filedialog.askdirectory(
            title='新しいデータベースフォルダを選択してください',
            initialdir=os.path.dirname(app.config['SQLALCHEMY_DATABASE_URI'].replace('sqlite:///', ''))
        )
        
        root.destroy()
        
        if folder:
            config = load_config()
            config['database_folder'] = folder
            save_config(config)
            
            flash(f'データベースフォルダを変更しました: {folder}\nアプリを再起動してください。', 'success')
        else:
            flash('フォルダが選択されませんでした', 'warning')
    except Exception as e:
        flash(f'エラー: {str(e)}', 'danger')
    
    return redirect(url_for('settings'))

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
