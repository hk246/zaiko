"""PyInstaller build script – generates a onedir distribution.

Usage:
    python build.py

This will:
  1. Build the React frontend (npm run build)
  2. Run PyInstaller to create a distributable folder
  3. Output at: dist/在庫管理システム/
"""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FRONTEND_DIR = ROOT / "frontend"
DIST_DIR = ROOT / "dist"


def build_frontend():
    """Build the React frontend with Vite."""
    print("\n=== Building frontend ===")
    result = subprocess.run(
        ["npm", "run", "build"],
        cwd=str(FRONTEND_DIR),
        shell=True,
    )
    if result.returncode != 0:
        print("ERROR: Frontend build failed!")
        sys.exit(1)
    print("Frontend build OK\n")


def build_pyinstaller():
    """Run PyInstaller with the spec file."""
    print("=== Running PyInstaller ===")
    spec_file = ROOT / "在庫管理.spec"
    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", str(spec_file), "--noconfirm"],
        cwd=str(ROOT),
    )
    if result.returncode != 0:
        print("ERROR: PyInstaller build failed!")
        sys.exit(1)
    print(f"\nBuild complete! Output: {DIST_DIR / '在庫管理システム'}")


if __name__ == "__main__":
    build_frontend()
    build_pyinstaller()
