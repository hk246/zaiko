"""PyInstaller build script – generates a onedir distribution.

Usage:
    python build.py

ビルドPCに必要なもの:
  - Python 3.12+ (pip install pyinstaller 済み)
  - Node.js は不要 (frontend/dist/ はgitに含まれている)

出力先: dist/在庫管理システム/
"""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FRONTEND_DIST = ROOT / "frontend" / "dist"
DIST_DIR = ROOT / "dist"


def check_frontend():
    """Verify frontend/dist/ exists (pre-built and committed to git)."""
    index = FRONTEND_DIST / "index.html"
    if not index.exists():
        print("ERROR: frontend/dist/index.html が見つかりません。")
        print()
        print("  開発PCで以下を実行してからgit pushしてください:")
        print("    cd frontend && npm run build")
        print()
        sys.exit(1)
    print(f"  Frontend dist OK: {FRONTEND_DIST}")


def build_pyinstaller():
    """Run PyInstaller with the spec file."""
    print("\n=== Running PyInstaller ===")
    spec_file = ROOT / "在庫管理.spec"
    if not spec_file.exists():
        print(f"ERROR: {spec_file} が見つかりません。")
        sys.exit(1)

    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", str(spec_file), "--noconfirm"],
        cwd=str(ROOT),
    )
    if result.returncode != 0:
        print("ERROR: PyInstaller build failed!")
        sys.exit(1)

    out = DIST_DIR / "在庫管理システム"
    print(f"\n{'=' * 40}")
    print(f"  Build complete!")
    print(f"  Output: {out}")
    print(f"  実行:   {out / '在庫管理システム.exe'}")
    print(f"{'=' * 40}")


if __name__ == "__main__":
    print("=== 在庫管理システム ビルド ===")
    check_frontend()
    build_pyinstaller()
