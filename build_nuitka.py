#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Nuitka ビルドスクリプト
=======================
使い方:
    python build_nuitka.py

出力:
    dist/在庫管理システム.dist/ フォルダ（exe + 依存ライブラリ + templates）

配布方法:
    dist/在庫管理システム.dist/ フォルダをまるごとZIPして配布してください。
    受け取り側は ZIP を解凍し、「在庫管理システム.exe」をダブルクリックするだけです。

注意:
    - ビルド前に `pip install nuitka` が必要です
    - 初回ビルドは数分かかります（C コンパイルを実行するため）
    - C コンパイラが必要です（Windows: MSVC または MinGW-w64）
      MinGW-w64 インストール: https://www.mingw-w64.org/
      または Visual Studio Build Tools をインストールしてください
"""

import subprocess
import sys
import os
import shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_NAME = "在庫管理システム"
OUTPUT_DIR = os.path.join(BASE_DIR, "dist")
DIST_APP_DIR = os.path.join(OUTPUT_DIR, f"{DIST_NAME}.dist")


def check_nuitka():
    """Nuitka がインストールされているか確認"""
    try:
        # text=True / capture_output=True は Windows で文字コードデッドロックが起きるため
        # bytes モードで受け取り、手動デコードする
        result = subprocess.run(
            [sys.executable, "-m", "nuitka", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=60,
        )
        if result.returncode == 0:
            version = (result.stdout or result.stderr).decode("utf-8", errors="replace").strip()
            print(f"Nuitka バージョン: {version}")
            return True
    except subprocess.TimeoutExpired:
        print("エラー: Nuitka の応答がタイムアウトしました（60秒）")
        print("  ターミナルで  python -m nuitka --version  を直接実行して確認してください")
        return False
    except Exception:
        pass
    print("エラー: Nuitka がインストールされていません")
    print("  pip install nuitka  を実行してインストールしてください")
    return False


def build():
    """Nuitka でビルドを実行"""

    # templates / static フォルダの存在確認
    templates_dir = os.path.join(BASE_DIR, "templates")
    static_dir    = os.path.join(BASE_DIR, "static")

    if not os.path.exists(templates_dir):
        print(f"エラー: templates/ フォルダが見つかりません: {templates_dir}")
        return False

    # ビルドコマンド組み立て
    cmd = [
        sys.executable, "-m", "nuitka",

        # ========== ビルドモード ==========
        "--standalone",          # 依存ライブラリを dist フォルダにまとめる

        # ========== Flask コア ==========
        "--include-package=flask",
        "--include-package=flask_sqlalchemy",
        "--include-package=flask_wtf",
        "--include-package=wtforms",

        # ========== SQLAlchemy + SQLite ダイアレクト ==========
        "--include-package=sqlalchemy",
        # SQLite ダイアレクトは動的ロードのため明示的に含める
        "--include-package=sqlalchemy.dialects",
        "--include-package=sqlalchemy.dialects.sqlite",

        # ========== Jinja2 / Werkzeug / 依存パッケージ ==========
        "--include-package=jinja2",
        "--include-package=werkzeug",
        "--include-package=markupsafe",
        "--include-package=itsdangerous",
        "--include-package=click",
        "--include-package=blinker",         # Flask のシグナル処理

        # ========== openpyxl（Excel エクスポート） ==========
        "--include-package=openpyxl",
        "--include-package=et_xmlfile",       # openpyxl 依存
        # openpyxl のデータファイル（Excelテンプレート XML など）
        "--include-package-data=openpyxl",

        # ========== メール送信 ==========
        # smtplib は stdlib なので自動的に含まれる

        # ========== email_validator (Flask-WTF Email バリデータ) ==========
        # インストールされていない場合は wtforms のデフォルトバリデータを使用
    ]

    # email_validator が入っていれば含める（任意）
    try:
        import email_validator
        cmd.append("--include-package=email_validator")
        import dns
        cmd.append("--include-package=dns")   # dnspython (email_validator 依存)
    except ImportError:
        pass

    cmd += [
        # ========== データファイル ==========
        # templates フォルダを dist に含める
        f"--include-data-dir={templates_dir}=templates",
    ]

    # static フォルダがあれば含める
    if os.path.exists(static_dir):
        cmd.append(f"--include-data-dir={static_dir}=static")

    cmd += [
        # ========== tkinter プラグイン ==========
        "--enable-plugin=tk-inter",

        # ========== Windows 設定 ==========
        "--windows-disable-console",    # コンソールウィンドウを非表示
        # "--windows-icon-from-ico=icon.ico",  # アイコンファイルがあれば有効化

        # ========== 出力設定 ==========
        f"--output-filename={DIST_NAME}",
        f"--output-dir={OUTPUT_DIR}",

        # ========== ソースファイル ==========
        os.path.join(BASE_DIR, "app.py"),
    ]

    print("=" * 60)
    print("Nuitka ビルド開始")
    print("=" * 60)
    print(f"出力先: {DIST_APP_DIR}")
    print()

    result = subprocess.run(cmd, cwd=BASE_DIR)

    if result.returncode != 0:
        print()
        print("=" * 60)
        print("ビルドに失敗しました")
        print("=" * 60)
        return False

    print()
    print("=" * 60)
    print("ビルド成功!")
    print("=" * 60)

    # ビルド後の後処理
    post_build()
    return True


def post_build():
    """ビルド後の後処理: 不要ファイルの削除、配布用フォルダ整理"""

    if not os.path.exists(DIST_APP_DIR):
        print(f"警告: dist フォルダが見つかりません: {DIST_APP_DIR}")
        return

    # instance フォルダを作成（初回起動時の DB 保存先）
    instance_dir = os.path.join(DIST_APP_DIR, "instance")
    os.makedirs(instance_dir, exist_ok=True)
    print(f"instance/ フォルダを作成: {instance_dir}")

    # README.txt を作成（配布先向け）
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

    # ZIP ファイルを作成
    zip_path = os.path.join(OUTPUT_DIR, DIST_NAME)
    print(f"\nZIPファイルを作成中: {zip_path}.zip")
    shutil.make_archive(zip_path, "zip", OUTPUT_DIR, f"{DIST_NAME}.dist")
    print(f"ZIP作成完了: {zip_path}.zip")

    print()
    print("【配布ファイル】")
    print(f"  フォルダ: {DIST_APP_DIR}")
    print(f"  ZIP:      {zip_path}.zip")
    print()
    print("「在庫管理システム.dist」フォルダをまるごと渡すか、")
    print("「在庫管理システム.zip」を渡してください。")


if __name__ == "__main__":
    print("在庫管理システム Nuitka ビルドスクリプト")
    print()

    if not check_nuitka():
        sys.exit(1)

    success = build()
    if not success:
        sys.exit(1)

    print()
    print("完了!")
