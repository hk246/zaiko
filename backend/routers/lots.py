"""Lot management endpoints."""

from datetime import datetime

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.lot import Lot
from backend.models.master import Tare
from backend.models.material import RawMaterial
from backend.models.property_field import LotProperty, PropertyField
from backend.schemas.lot import LotCreate, LotPropertyRead, LotPropertyWrite, LotRead, LotUpdate

router = APIRouter(prefix="/materials/{material_id}/lots", tags=["lots"])


def _lot_to_read(lot: Lot, db: Session) -> dict:
    props = []
    for p in lot.properties:
        field = db.get(PropertyField, p.field_id)
        if field:
            props.append(LotPropertyRead(
                field_id=field.id,
                field_name=field.name,
                field_type=field.field_type,
                unit=field.unit,
                value_string=p.value_string,
                value_number=p.value_number,
            ))
    return {
        **{c.name: getattr(lot, c.name) for c in lot.__table__.columns},
        "properties": props,
    }


@router.get("", response_model=list[LotRead])
def list_lots(material_id: int, db: Session = Depends(get_db)):
    m = db.get(RawMaterial, material_id)
    if not m or m.deleted_at:
        raise HTTPException(404, "Material not found")
    lots = (
        db.query(Lot)
        .filter(Lot.material_id == material_id, Lot.deleted_at == None)  # noqa: E711
        .order_by(Lot.created_at.desc())
        .all()
    )
    return [_lot_to_read(l, db) for l in lots]


@router.post("", response_model=LotRead, status_code=201)
def create_lot(material_id: int, body: LotCreate, db: Session = Depends(get_db)):
    m = db.get(RawMaterial, material_id)
    if not m or m.deleted_at:
        raise HTTPException(404, "Material not found")

    weight = body.weight
    if body.tare_id:
        tare = db.get(Tare, body.tare_id)
        if tare:
            weight = max(0, weight - tare.weight)

    lot = Lot(
        material_id=material_id,
        lot_name=body.lot_name,
        weight=weight,
        is_fraction=body.is_fraction,
    )
    db.add(lot)
    db.commit()
    db.refresh(lot)
    return _lot_to_read(lot, db)


@router.put("/{lot_id}", response_model=LotRead)
def update_lot(material_id: int, lot_id: int, body: LotUpdate, db: Session = Depends(get_db)):
    lot = db.get(Lot, lot_id)
    if not lot or lot.deleted_at or lot.material_id != material_id:
        raise HTTPException(404, "Lot not found")
    data = body.model_dump(exclude_unset=True)
    tare_id = data.pop("tare_id", None)
    for k, v in data.items():
        setattr(lot, k, v)
    if tare_id and body.weight is not None:
        tare = db.get(Tare, tare_id)
        if tare:
            lot.weight = max(0, body.weight - tare.weight)
    db.commit()
    db.refresh(lot)
    return _lot_to_read(lot, db)


@router.delete("/{lot_id}", status_code=204)
def delete_lot(material_id: int, lot_id: int, db: Session = Depends(get_db)):
    lot = db.get(Lot, lot_id)
    if not lot or lot.deleted_at or lot.material_id != material_id:
        raise HTTPException(404, "Lot not found")
    lot.deleted_at = datetime.utcnow()
    db.commit()


@router.patch("/{lot_id}/fraction", response_model=LotRead)
def toggle_fraction(material_id: int, lot_id: int, db: Session = Depends(get_db)):
    lot = db.get(Lot, lot_id)
    if not lot or lot.deleted_at or lot.material_id != material_id:
        raise HTTPException(404, "Lot not found")
    lot.is_fraction = not lot.is_fraction
    db.commit()
    db.refresh(lot)
    return _lot_to_read(lot, db)


# ── Lot Properties ──

@router.get("/{lot_id}/properties", response_model=list[LotPropertyRead])
def get_lot_properties(material_id: int, lot_id: int, db: Session = Depends(get_db)):
    lot = db.get(Lot, lot_id)
    if not lot or lot.deleted_at or lot.material_id != material_id:
        raise HTTPException(404, "Lot not found")
    result = []
    for p in lot.properties:
        field = db.get(PropertyField, p.field_id)
        if field:
            result.append(LotPropertyRead(
                field_id=field.id,
                field_name=field.name,
                field_type=field.field_type,
                unit=field.unit,
                value_string=p.value_string,
                value_number=p.value_number,
            ))
    return result


@router.put("/{lot_id}/properties")
def update_lot_properties(
    material_id: int,
    lot_id: int,
    body: list[LotPropertyWrite],
    db: Session = Depends(get_db),
):
    lot = db.get(Lot, lot_id)
    if not lot or lot.deleted_at or lot.material_id != material_id:
        raise HTTPException(404, "Lot not found")

    for item in body:
        existing = (
            db.query(LotProperty)
            .filter(LotProperty.lot_id == lot_id, LotProperty.field_id == item.field_id)
            .first()
        )
        if existing:
            existing.value_string = item.value_string
            existing.value_number = item.value_number
            existing.updated_at = datetime.utcnow()
        else:
            db.add(LotProperty(
                lot_id=lot_id,
                field_id=item.field_id,
                value_string=item.value_string,
                value_number=item.value_number,
            ))
    db.commit()
    return {"status": "ok"}
