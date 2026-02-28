"""Excel import endpoints – template download, preview, and confirm."""

from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.services.import_service import (
    VALID_TYPES,
    TYPE_LABELS,
    generate_template,
    parse_excel,
    validate_lots,
    validate_lot_properties,
    import_lots,
    import_lot_properties,
    VALIDATORS,
    IMPORTERS,
)

router = APIRouter(prefix="/import", tags=["import"])

# Data types that require material_id
_MATERIAL_SCOPED = {"lots", "lot_properties"}


@router.get("/template/{data_type}")
def download_template(data_type: str):
    """Download an Excel template for the given data type."""
    if data_type not in VALID_TYPES:
        raise HTTPException(400, f"Invalid data type: {data_type}")
    buf = generate_template(data_type)
    filename = f"template_{data_type}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.post("/{data_type}")
def import_data(
    data_type: str,
    confirm: bool = Query(False),
    material_id: int | None = Query(None),
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    """
    Upload an Excel file for import.
    - confirm=false → preview only (parse + validate, return row count & errors)
    - confirm=true  → actually insert into DB
    """
    if data_type not in VALID_TYPES:
        raise HTTPException(400, f"Invalid data type: {data_type}")
    if data_type in _MATERIAL_SCOPED and material_id is None:
        raise HTTPException(400, "material_id is required for this import type")

    # Read file
    file_bytes = file.file.read()
    if not file_bytes:
        raise HTTPException(400, "Empty file")

    # Parse
    try:
        rows = parse_excel(data_type, file_bytes)
    except Exception as e:
        raise HTTPException(400, f"Excelファイルの読み込みに失敗しました: {e}")

    if not rows:
        raise HTTPException(400, "データ行がありません")

    # Validate
    if data_type == "lots":
        errors = validate_lots(rows, db, material_id)
    elif data_type == "lot_properties":
        errors = validate_lot_properties(rows, db, material_id)
    else:
        validator = VALIDATORS[data_type]
        errors = validator(rows, db)

    # Preview mode
    if not confirm:
        return {
            "total_rows": len(rows),
            "errors": errors,
            "valid": len(errors) == 0,
        }

    # Confirm mode – must have no errors
    if errors:
        raise HTTPException(
            422,
            detail={
                "message": "バリデーションエラーがあります",
                "errors": errors,
            },
        )

    # Import
    if data_type == "lots":
        count = import_lots(rows, db, material_id)
    elif data_type == "lot_properties":
        count = import_lot_properties(rows, db, material_id)
    else:
        importer = IMPORTERS[data_type]
        count = importer(rows, db)

    label = TYPE_LABELS.get(data_type, data_type)
    return {"imported": count, "message": f"{count}件の{label}をインポートしました"}
