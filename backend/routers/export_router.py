"""Export endpoint – download all data as a single Excel file, and reimport."""

from datetime import datetime
from urllib.parse import quote

from fastapi import APIRouter, Depends, HTTPException, Query, UploadFile, File
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.services.export_service import generate_export_workbook
from backend.services.reimport_service import preview_reimport, execute_reimport

router = APIRouter(prefix="/export", tags=["export"])


@router.get("")
def export_all(db: Session = Depends(get_db)):
    """Generate and download a multi-sheet Excel workbook with all data."""
    try:
        buf = generate_export_workbook(db)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"エクスポート生成に失敗しました: {e}")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename_jp = f"在庫データ_{timestamp}.xlsx"
    filename_ascii = f"inventory_export_{timestamp}.xlsx"
    encoded = quote(filename_jp)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename={filename_ascii}; filename*=UTF-8''{encoded}",
        },
    )


@router.post("/import")
async def reimport_data(
    confirm: bool = Query(default=False),
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    """Reimport exported Excel to update/create records.

    - confirm=false: preview (validate only, no DB changes)
    - confirm=true: execute upsert
    """
    file_bytes = await file.read()

    if not confirm:
        try:
            return preview_reimport(file_bytes, db)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Excelファイルの解析に失敗しました: {e}")
    else:
        try:
            return execute_reimport(file_bytes, db)
        except ValueError as e:
            raise HTTPException(status_code=422, detail=str(e))
        except Exception as e:
            db.rollback()
            raise HTTPException(status_code=500, detail=f"インポート実行中にエラーが発生しました: {e}")
