"""Backup & restore service for SQLite database."""

import shutil
from datetime import datetime
from pathlib import Path

from backend.config import settings


def _db_path() -> Path:
    return Path(settings.database_folder) / "inventory.db"


def _backup_dir() -> Path:
    d = Path(settings.backup_folder)
    d.mkdir(parents=True, exist_ok=True)
    return d


def list_backups() -> list[dict]:
    """List all backup files with metadata."""
    backups = []
    for f in sorted(_backup_dir().glob("*.db"), reverse=True):
        stat = f.stat()
        backups.append({
            "filename": f.name,
            "size_bytes": stat.st_size,
            "created_at": datetime.fromtimestamp(stat.st_mtime).isoformat(),
        })
    return backups


def create_backup() -> dict:
    """Create a timestamped backup of the current database."""
    src = _db_path()
    if not src.exists():
        raise FileNotFoundError("Database file not found")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = _backup_dir() / f"backup_{ts}.db"
    shutil.copy2(str(src), str(dest))

    return {
        "filename": dest.name,
        "size_bytes": dest.stat().st_size,
        "created_at": datetime.now().isoformat(),
    }


def restore_backup(filename: str) -> None:
    """Restore database from a backup file.

    Automatically creates a safety backup before restoring.
    """
    backup_file = _backup_dir() / filename
    if not backup_file.exists():
        raise FileNotFoundError(f"Backup {filename} not found")

    # Safety backup
    src = _db_path()
    if src.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safety = _backup_dir() / f"pre_restore_{ts}.db"
        shutil.copy2(str(src), str(safety))

    shutil.copy2(str(backup_file), str(src))


def delete_backup(filename: str) -> None:
    """Delete a backup file."""
    backup_file = _backup_dir() / filename
    if not backup_file.exists():
        raise FileNotFoundError(f"Backup {filename} not found")
    backup_file.unlink()


def get_backup_path(filename: str) -> Path:
    """Get full path for downloading."""
    p = _backup_dir() / filename
    if not p.exists():
        raise FileNotFoundError(f"Backup {filename} not found")
    return p
