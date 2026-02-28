"""PropertyField management endpoints."""

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.property_field import PropertyField
from backend.schemas.property_field import (
    PropertyFieldCreate,
    PropertyFieldRead,
    PropertyFieldReorder,
    PropertyFieldUpdate,
)

router = APIRouter(prefix="/property-fields", tags=["property-fields"])


@router.get("", response_model=list[PropertyFieldRead])
def list_fields(label_id: int | None = Query(default=None), db: Session = Depends(get_db)):
    query = db.query(PropertyField)
    if label_id is not None:
        query = query.filter(PropertyField.label_id == label_id)
    return query.order_by(PropertyField.order_index).all()


@router.post("", response_model=PropertyFieldRead, status_code=201)
def create_field(body: PropertyFieldCreate, db: Session = Depends(get_db)):
    field = PropertyField(**body.model_dump())
    db.add(field)
    db.commit()
    db.refresh(field)
    return field


@router.put("/{field_id}", response_model=PropertyFieldRead)
def update_field(field_id: int, body: PropertyFieldUpdate, db: Session = Depends(get_db)):
    field = db.get(PropertyField, field_id)
    if not field:
        raise HTTPException(404, "Field not found")
    for k, v in body.model_dump(exclude_unset=True).items():
        setattr(field, k, v)
    db.commit()
    db.refresh(field)
    return field


@router.delete("/{field_id}", status_code=204)
def delete_field(field_id: int, db: Session = Depends(get_db)):
    field = db.get(PropertyField, field_id)
    if not field:
        raise HTTPException(404, "Field not found")
    db.delete(field)
    db.commit()


@router.post("/reorder")
def reorder_fields(body: PropertyFieldReorder, db: Session = Depends(get_db)):
    for idx, fid in enumerate(body.field_ids):
        field = db.get(PropertyField, fid)
        if field:
            field.order_index = idx
    db.commit()
    return {"status": "ok"}
