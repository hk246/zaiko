"""Application configuration using Pydantic Settings."""

import json
import os
import sys
from pathlib import Path

from pydantic_settings import BaseSettings


def _get_app_root() -> Path:
    """Return the application root directory.

    - PyInstaller (frozen): directory where the .exe lives
    - Normal Python:        project root (one level up from backend/)
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


PROJECT_ROOT = _get_app_root()
CONFIG_FILE = PROJECT_ROOT / "config.json"
DEFAULT_DB_FOLDER = PROJECT_ROOT / "instance"


def _load_config() -> dict:
    """Load config.json as dict, fallback to empty."""
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _save_config(data: dict) -> None:
    """Persist config dict to config.json."""
    existing = _load_config()
    existing.update(data)
    CONFIG_FILE.write_text(
        json.dumps(existing, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _load_db_folder() -> str:
    """Load database folder from config.json, fallback to default."""
    folder = _load_config().get("database_folder", "")
    if folder and Path(folder).is_dir():
        return folder
    DEFAULT_DB_FOLDER.mkdir(parents=True, exist_ok=True)
    return str(DEFAULT_DB_FOLDER)


def _load_backup_folder() -> str:
    """Load backup folder from config.json, fallback to default."""
    folder = _load_config().get("backup_folder", "")
    if folder:
        return folder
    return str(PROJECT_ROOT / "backups")


class Settings(BaseSettings):
    app_name: str = "在庫管理システム"
    app_version: str = "2.0.0"
    host: str = "127.0.0.1"
    port: int = 5000
    debug: bool = False
    secret_key: str = "change-me-in-production"
    database_folder: str = ""
    backup_folder: str = ""

    def model_post_init(self, __context: object) -> None:
        if not self.database_folder:
            self.database_folder = _load_db_folder()
        if not self.backup_folder:
            self.backup_folder = _load_backup_folder()
        Path(self.backup_folder).mkdir(parents=True, exist_ok=True)

    @property
    def database_url(self) -> str:
        db_path = Path(self.database_folder) / "inventory.db"
        return f"sqlite:///{db_path}"

    def save_database_folder(self, folder: str) -> None:
        """Persist database folder to config.json."""
        self.database_folder = folder
        _save_config({"database_folder": folder})

    def save_backup_folder(self, folder: str) -> None:
        """Persist backup folder to config.json."""
        self.backup_folder = folder
        Path(folder).mkdir(parents=True, exist_ok=True)
        _save_config({"backup_folder": folder})


settings = Settings()
