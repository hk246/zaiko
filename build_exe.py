"""
在庫管理システムをPyInstallerでビルドするスクリプト

使用方法:
1. 必要なパッケージをインストール
   pip install pyinstaller

2. このスクリプトを実行
   python build_exe.py

3. dist/在庫管理システム フォルダが生成されます
"""

import sys
import os
from tkinter import Tk, filedialog, messagebox

# PyInstallerがインストールされているか確認
try:
    import PyInstaller.__main__
except ImportError:
    print("エラー: PyInstallerがインストールされていません")
    print("以下のコマンドを実行してください:")
    print("  pip install pyinstaller")
    sys.exit(1)

# カレントディレクトリを取得
current_dir = os.path.dirname(os.path.abspath(__file__))

print("="*60)
print("在庫管理システムのビルドを開始します...")
print("="*60)

# アイコンファイルの選択
icon_path = None
root = Tk()
root.withdraw()
root.attributes('-topmost', True)

# デフォルトのアイコンファイルをチェック
if os.path.exists('icon.ico'):
    use_default = messagebox.askyesno(
        'アイコンファイル',
        'icon.ico が見つかりました。\nこのアイコンを使用しますか？\n\n「いいえ」を選択すると別のアイコンを選択できます。'
    )
    if use_default:
        icon_path = 'icon.ico'
        print(f"✓ デフォルトアイコンを使用: {icon_path}")

# アイコンが選択されていない場合、ダイアログを表示
if icon_path is None:
    select_icon = messagebox.askyesno(
        'アイコンの選択',
        'アイコンファイル(.ico)を選択しますか？\n\n「いいえ」を選択するとデフォルトアイコンで生成されます。'
    )
    
    if select_icon:
        icon_path = filedialog.askopenfilename(
            title='アイコンファイルを選択',
            filetypes=[('アイコンファイル', '*.ico'), ('すべてのファイル', '*.*')],
            initialdir=current_dir
        )
        if icon_path:
            print(f"✓ 選択されたアイコン: {icon_path}")
        else:
            print("! アイコンが選択されませんでした。デフォルトアイコンを使用します。")

root.destroy()

# ビルドオプション
build_options = [
    'app.py',
    '--name=在庫管理システム',
    '--onedir',  # フォルダ形式で出力（起動が高速）
    '--windowed',  # コンソールウィンドウを表示しない
    '--add-data=templates;templates',
    '--add-data=static;static',
    '--hidden-import=flask',
    '--hidden-import=flask_sqlalchemy',
    '--hidden-import=flask_wtf',
    '--hidden-import=wtforms',
    '--hidden-import=wtforms.validators',
    '--hidden-import=email.mime.text',
    '--hidden-import=email.mime.multipart',
    '--hidden-import=tkinter',
    '--hidden-import=tkinter.filedialog',
    '--hidden-import=tkinter.messagebox',
    '--hidden-import=webbrowser',
    '--hidden-import=threading',
    '--hidden-import=shutil',
    '--hidden-import=sqlalchemy.sql.default_comparator',
    '--clean',
    '--noconfirm',
]

# アイコンファイルを追加
if icon_path and os.path.exists(icon_path):
    build_options.append(f'--icon={icon_path}')
    print(f"✓ アイコンを設定: {icon_path}")

print("\nビルドを開始します。数分かかる場合があります...")
print("-"*60)

try:
    PyInstaller.__main__.run(build_options)
    
    # 配布用READMEをコピー
    dist_folder = os.path.join('dist', '在庫管理システム')
    readme_source = 'DISTRIBUTION_README.txt'
    readme_dest = os.path.join(dist_folder, 'README.txt')
    
    if os.path.exists(readme_source) and os.path.exists(dist_folder):
        import shutil
        shutil.copy2(readme_source, readme_dest)
        print(f"\n✓ 配布用READMEをコピーしました")
    
    print("\n" + "="*60)
    print("✓ ビルド完了！")
    print("="*60)
    print(f"\n📁 出力フォルダ: dist\\在庫管理システム\\")
    print(f"   実行ファイル: dist\\在庫管理システム\\在庫管理システム.exe")
    print("\n📦 配布方法:")
    print("  1. dist\\在庫管理システム フォルダ全体を相手先PCにコピー")
    print("     （フォルダ内の全ファイルが必要です）")
    print("  2. 在庫管理システム.exe を実行")
    print("  3. 初回起動時にデータベース保存場所を選択してください")
    print("\n💡 注意:")
    print("  - フォルダ内の全ファイルをまとめて配布してください")
    print("  - 相手先PCにPythonのインストールは不要です")
    print("  - Windows Defenderに警告される場合があります（正常な動作です）")
    print("  - 詳細は DISTRIBUTION.md をご覧ください")
    print("="*60)
    
except Exception as e:
    print("\n" + "="*60)
    print("✗ ビルドに失敗しました")
    print("="*60)
    print(f"エラー: {e}")
    print("\n対処方法:")
    print("  1. requirements.txt の全パッケージがインストールされているか確認")
    print("  2. templates/ と static/ フォルダが存在するか確認")
    print("  3. 管理者権限で実行してみる")
    sys.exit(1)

