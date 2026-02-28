"""Settings endpoint – database folder configuration."""

from pathlib import Path

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from backend.config import settings

router = APIRouter(prefix="/settings", tags=["settings"])


class SettingsRead(BaseModel):
    database_folder: str
    backup_folder: str
    app_version: str


class FolderUpdate(BaseModel):
    folder: str


@router.get("", response_model=SettingsRead)
def get_settings():
    return SettingsRead(
        database_folder=settings.database_folder,
        backup_folder=settings.backup_folder,
        app_version=settings.app_version,
    )


@router.put("/database-folder")
def update_database_folder(body: FolderUpdate):
    folder = body.folder
    if not Path(folder).is_dir():
        raise HTTPException(400, f"フォルダが存在しません: {folder}")
    settings.save_database_folder(folder)
    return {"status": "ok", "database_folder": folder, "restart_required": True}


@router.put("/backup-folder")
def update_backup_folder(body: FolderUpdate):
    folder = body.folder
    try:
        Path(folder).mkdir(parents=True, exist_ok=True)
    except OSError:
        raise HTTPException(400, f"フォルダを作成できません: {folder}")
    settings.save_backup_folder(folder)
    return {"status": "ok", "backup_folder": folder, "restart_required": True}
