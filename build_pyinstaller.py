#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller ビルドスクリプト (onedir モード)
==========================================

使い方:
    python build_pyinstaller.py

出力:
    dist/在庫管理システム/ フォルダ（exe + 依存ライブラリ + templates）

配布方法:
    dist/在庫管理システム/ フォルダをまるごと ZIP して配布してください。
    受け取り側は ZIP を解凍し、「在庫管理システム.exe」をダブルクリックするだけです。
"""

import subprocess
import sys
import os
import shutil

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DIST_NAME = "在庫管理システム"
OUTPUT_DIR    = os.path.join(BASE_DIR, "dist")
DIST_APP_DIR  = os.path.join(OUTPUT_DIR, DIST_NAME)


def check_pyinstaller():
    """PyInstaller がインストールされているか確認"""
    try:
        result = subprocess.run(
            [sys.executable, "-m", "PyInstaller", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=30,
        )
        if result.returncode == 0:
            version = (result.stdout or result.stderr).decode("utf-8", errors="replace").strip()
            print(f"PyInstaller バージョン: {version}")
            return True
    except subprocess.TimeoutExpired:
        print("エラー: PyInstaller の応答がタイムアウトしました")
    except Exception:
        pass
    print("エラー: PyInstaller がインストールされていません")
    print("  pip install pyinstaller  を実行してください")
    return False


def build():
    """PyInstaller でビルドを実行"""

    templates_dir = os.path.join(BASE_DIR, "templates")
    static_dir    = os.path.join(BASE_DIR, "static")

    if not os.path.exists(templates_dir):
        print(f"エラー: templates/ フォルダが見つかりません: {templates_dir}")
        return False

    # Windows の --add-data 区切り文字は「;」
    sep = ";" if sys.platform == "win32" else ":"

    cmd = [
        sys.executable, "-m", "PyInstaller",

        # ========== ビルドモード ==========
        "--onedir",                  # フォルダ形式（onedir）で出力
        "--noconsole",               # コンソールウィンドウを非表示
        "--clean",                   # 前回のキャッシュを削除してクリーンビルド

        # ========== 出力設定 ==========
        f"--name={DIST_NAME}",
        f"--distpath={OUTPUT_DIR}",
        f"--workpath={os.path.join(BASE_DIR, 'build')}",
        f"--specpath={BASE_DIR}",

        # ========== データファイル ==========
        f"--add-data={templates_dir}{sep}templates",
    ]

    # static フォルダが存在すれば含める
    if os.path.exists(static_dir):
        cmd.append(f"--add-data={static_dir}{sep}static")

    cmd += [
        # ========== 隠れた依存パッケージ（Flask 系） ==========
        "--hidden-import=flask",
        "--hidden-import=flask_sqlalchemy",
        "--hidden-import=flask_wtf",
        "--hidden-import=wtforms",
        "--hidden-import=wtforms.validators",
        "--hidden-import=wtforms.fields",

        # ========== SQLAlchemy + SQLite ダイアレクト ==========
        "--hidden-import=sqlalchemy",
        "--hidden-import=sqlalchemy.dialects.sqlite",
        "--hidden-import=sqlalchemy.dialects.sqlite.pysqlite",
        "--hidden-import=sqlalchemy.orm",

        # ========== Jinja2 / Werkzeug / 周辺パッケージ ==========
        "--hidden-import=jinja2",
        "--hidden-import=werkzeug",
        "--hidden-import=werkzeug.serving",
        "--hidden-import=markupsafe",
        "--hidden-import=itsdangerous",
        "--hidden-import=click",
        "--hidden-import=blinker",

        # ========== Excel エクスポート ==========
        "--hidden-import=openpyxl",
        "--hidden-import=et_xmlfile",

        # ========== tkinter（フォルダ選択ダイアログ） ==========
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",

        # ========== ソースファイル ==========
        os.path.join(BASE_DIR, "app.py"),
    ]

    print("=" * 60)
    print("PyInstaller ビルド開始")
    print("=" * 60)
    print(f"出力先: {DIST_APP_DIR}")
    print()

    result = subprocess.run(cmd, cwd=BASE_DIR)

    if result.returncode != 0:
        print()
        print("=" * 60)
        print("ビルドに失敗しました")
        print("上記のエラーメッセージを確認してください")
        print("=" * 60)
        return False

    print()
    print("=" * 60)
    print("ビルド成功!")
    print("=" * 60)

    post_build()
    return True


def post_build():
    """ビルド後の後処理: instance フォルダ作成・README 生成・ZIP 作成"""

    if not os.path.exists(DIST_APP_DIR):
        print(f"警告: dist フォルダが見つかりません: {DIST_APP_DIR}")
        return

    # instance フォルダを作成（初回起動時の DB 保存先）
    instance_dir = os.path.join(DIST_APP_DIR, "instance")
    os.makedirs(instance_dir, exist_ok=True)
    print(f"instance/ フォルダを作成: {instance_dir}")

    # 配布先向け README.txt を作成
    readme_path = os.path.join(DIST_APP_DIR, "README.txt")
    readme_content = """在庫管理システム
================

【起動方法】
  「在庫管理システム.exe」をダブルクリックしてください。
  起動後、自動的にブラウザが開きます。

【初回起動】
  データベースの保存フォルダを選択するダイアログが表示されます。
  共有フォルダを指定すると、複数人で同じデータを使用できます。
  選択しない場合は、このフォルダ内の「instance」フォルダに保存されます。

【終了方法】
  ブラウザを閉じてから、タスクバーのアイコンを右クリックして終了してください。
  または、タスクマネージャーから「在庫管理システム.exe」を終了してください。

【データのバックアップ】
  データベースファイル（inventory.db）を定期的にバックアップしてください。
  アプリ内の「バックアップ」メニューからも操作できます。
"""
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(readme_content)
    print(f"README.txt を作成: {readme_path}")

    # ZIP ファイルを作成（プロジェクトルートに保存）
    zip_path = os.path.join(BASE_DIR, DIST_NAME)
    print(f"\nZIP ファイルを作成中: {zip_path}.zip")
    shutil.make_archive(zip_path, "zip", OUTPUT_DIR, DIST_NAME)
    print(f"ZIP 作成完了: {zip_path}.zip")

    print()
    print("=" * 60)
    print("【配布ファイル】")
    print(f"  フォルダ: {DIST_APP_DIR}")
    print(f"  ZIP:      {zip_path}.zip")
    print("=" * 60)
    print()
    print("「在庫管理システム」フォルダをまるごと渡すか、")
    print("「在庫管理システム.zip」を渡してください。")


if __name__ == "__main__":
    print("在庫管理システム PyInstaller ビルドスクリプト")
    print()

    if not check_pyinstaller():
        sys.exit(1)

    success = build()
    if not success:
        sys.exit(1)

    print()
    print("完了!")
