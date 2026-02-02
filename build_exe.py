"""
在庫管理システムを実行ファイル(.exe)にビルドするスクリプト

使い方:
1. PyInstallerをインストール: pip install pyinstaller
2. このスクリプトを実行: python build_exe.py
3. distフォルダ内に実行ファイルが作成されます
"""

import os
import subprocess
import sys
import shutil

def install_pyinstaller():
    """PyInstallerをインストール"""
    print("PyInstallerをインストールしています...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstallerのインストール完了")
        return True
    except subprocess.CalledProcessError:
        print("✗ PyInstallerのインストールに失敗しました")
        return False

def check_pyinstaller():
    """PyInstallerがインストールされているか確認"""
    try:
        import PyInstaller
        return True
    except ImportError:
        return False

def clean_build_folders():
    """ビルド前にクリーンアップ"""
    print("\nビルドフォルダをクリーンアップしています...")
    folders_to_clean = ['build', 'dist', '__pycache__']
    for folder in folders_to_clean:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"✓ {folder} を削除しました")
    
    # .specファイルも削除
    if os.path.exists('app.spec'):
        os.remove('app.spec')
        print("✓ app.spec を削除しました")

def create_spec_file():
    """PyInstaller用の.specファイルを作成"""
    print("\n.specファイルを作成しています...")
    
    # アイコンファイルの検索
    icon_path = None
    for icon_name in ['icon.ico', 'app.ico', '在庫管理.ico']:
        if os.path.exists(icon_name):
            icon_path = icon_name
            print(f"✓ アイコンファイルを検出: {icon_name}")
            break
    
    if not icon_path:
        print("  アイコンファイルが見つかりません（デフォルトアイコンを使用）")
    
    # iconパラメータの設定
    icon_param = f"'{icon_path}'" if icon_path else "None"
    
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
    ],
    hiddenimports=[
        'flask',
        'flask_sqlalchemy',
        'flask_wtf',
        'wtforms',
        'email.mime.text',
        'email.mime.multipart',
        'tkinter',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='在庫管理システム',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # コンソールを表示（エラー確認用）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON_PATH,  # アイコンファイルがあれば指定
)
"""
    
    # icon_paramを実際の値に置き換え
    spec_content = spec_content.replace('ICON_PATH', icon_param)
    
    with open('app.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✓ app.spec を作成しました")

def build_exe():
    """PyInstallerで実行ファイルをビルド"""
    print("\n実行ファイルをビルドしています...")
    print("（数分かかる場合があります）\n")
    
    try:
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller",
            "app.spec",
            "--clean"
        ])
        print("\n✓ ビルド完了！")
        return True
    except subprocess.CalledProcessError:
        print("\n✗ ビルドに失敗しました")
        return False

def create_readme():
    """配布用のREADMEを作成"""
    print("\n配布用のREADMEを作成しています...")
    
    readme_content = """# 在庫管理システム

## 使い方

### 初回起動
1. `在庫管理システム.exe` をダブルクリックして起動
2. データベースフォルダを選択（共有フォルダを選択すると複数人で使用可能）
3. ブラウザで自動的に開かない場合は、http://127.0.0.1:5000 にアクセス

### 複数人で使用する場合
1. 共有フォルダ（ネットワークドライブやOneDrive等）を用意
2. 各PCで同じ共有フォルダをデータベースフォルダに指定
3. 全員が同じデータベースにアクセスできます

### データベースフォルダの変更
- アプリ内の「設定」メニューから変更可能
- 変更後は再起動が必要です

## 注意事項
- アプリを終了するには、コンソールウィンドウで Ctrl+C を押すか、ウィンドウを閉じてください
- データは選択したフォルダ内の `inventory.db` に保存されます
- 定期的にバックアップを作成してください

## トラブルシューティング
- ブラウザが開かない場合: 手動で http://127.0.0.1:5000 にアクセス
- エラーが発生する場合: コンソールに表示されるエラーメッセージを確認
- データベースが開けない場合: 別のアプリがファイルを使用していないか確認

## システム要件
- Windows 10/11
- Webブラウザ（Chrome、Edge、Firefox等）
"""
    
    dist_folder = 'dist'
    if os.path.exists(dist_folder):
        readme_path = os.path.join(dist_folder, 'README.txt')
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write(readme_content)
        print(f"✓ {readme_path} を作成しました")

def main():
    print("=" * 60)
    print("在庫管理システム - 実行ファイルビルドツール")
    print("=" * 60)
    
    # PyInstallerの確認とインストール
    if not check_pyinstaller():
        print("\nPyInstallerがインストールされていません")
        response = input("インストールしますか？ (y/n): ").strip().lower()
        if response == 'y':
            if not install_pyinstaller():
                print("\nビルドを中止します")
                return
        else:
            print("\nビルドを中止します")
            return
    else:
        print("\n✓ PyInstallerがインストールされています")
    
    # クリーンアップ
    clean_build_folders()
    
    # .specファイル作成
    create_spec_file()
    
    # ビルド実行
    if build_exe():
        create_readme()
        
        print("\n" + "=" * 60)
        print("ビルド成功！")
        print("=" * 60)
        print(f"\n実行ファイルの場所: {os.path.abspath('dist')}")
        print("\n配布方法:")
        print("1. distフォルダ内の「在庫管理システム.exe」を配布")
        print("2. README.txt も一緒に配布すると親切です")
        print("\n注意:")
        print("- 初回起動時は少し時間がかかります")
        print("- ウイルス対策ソフトが警告を出す場合がありますが、安全です")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("ビルド失敗")
        print("=" * 60)
        print("\nエラーメッセージを確認して、以下を試してください:")
        print("1. 必要なパッケージがインストールされているか確認")
        print("2. app.py が正しいか確認")
        print("3. templates と static フォルダが存在するか確認")

if __name__ == '__main__':
    try:
        main()
        input("\nEnterキーを押して終了...")
    except KeyboardInterrupt:
        print("\n\nビルドをキャンセルしました")
    except Exception as e:
        print(f"\n予期しないエラーが発生しました: {e}")
        input("\nEnterキーを押して終了...")
