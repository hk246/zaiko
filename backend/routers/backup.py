"""Backup & export endpoints."""

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse, Response
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.services import backup_service, export_service

router = APIRouter(tags=["backup"])


# ── Backup ──

@router.get("/backup", response_model=list[dict])
def list_backups():
    return backup_service.list_backups()


@router.post("/backup")
def create_backup():
    try:
        return backup_service.create_backup()
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))


@router.post("/backup/restore/{filename}")
def restore_backup(filename: str):
    try:
        backup_service.restore_backup(filename)
        return {"status": "restored", "filename": filename}
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))


@router.get("/backup/download/{filename}")
def download_backup(filename: str):
    try:
        path = backup_service.get_backup_path(filename)
        return FileResponse(path, filename=filename, media_type="application/octet-stream")
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))


@router.delete("/backup/{filename}", status_code=204)
def delete_backup(filename: str):
    try:
        backup_service.delete_backup(filename)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))


# ── Export ──

@router.get("/export/csv")
def export_csv(db: Session = Depends(get_db)):
    csv_data = export_service.export_materials_csv(db)
    return Response(
        content=csv_data,
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=inventory.csv"},
    )


@router.get("/export/properties")
def export_properties(label_id: int | None = None, db: Session = Depends(get_db)):
    data = export_service.export_properties_excel(db, label_id)
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=properties.xlsx"},
    )
