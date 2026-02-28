"""Material CRUD endpoints."""

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.material import RawMaterial
from backend.schemas.material import MaterialCreate, MaterialListRead, MaterialRead, MaterialUpdate
from backend.services import stock_calculator

router = APIRouter(prefix="/materials", tags=["materials"])


def _enrich(db: Session, m: RawMaterial) -> dict:
    """Add computed stock fields to a material."""
    cs = stock_calculator.current_stock(db, m.id)
    fs = stock_calculator.fraction_stock(db, m.id)
    ps = stock_calculator.predicted_stock(db, m.id)
    low = stock_calculator.is_low_stock(db, m.id, m.min_weight)
    lot_count = len([l for l in m.lots if not l.deleted_at])
    return {
        **{c.name: getattr(m, c.name) for c in m.__table__.columns},
        "label": m.label,
        "current_stock": cs,
        "fraction_stock": fs,
        "predicted_stock": ps,
        "is_low_stock": low,
        "lot_count": lot_count,
        "contact_name": m.contact.name if m.contact else None,
        "contact_email": m.contact.email if m.contact else None,
    }


@router.get("", response_model=list[MaterialListRead])
def list_materials(
    q: str = Query(default="", description="Search by name"),
    label_id: int | None = Query(default=None),
    low_stock_only: bool = Query(default=False),
    db: Session = Depends(get_db),
):
    query = db.query(RawMaterial).filter(RawMaterial.deleted_at == None)  # noqa: E711
    if q:
        query = query.filter(RawMaterial.name.contains(q))
    if label_id:
        query = query.filter(RawMaterial.label_id == label_id)

    materials = query.order_by(RawMaterial.name).all()
    results = []
    for m in materials:
        data = _enrich(db, m)
        if low_stock_only and not data["is_low_stock"]:
            continue
        results.append(data)
    return results


@router.get("/{material_id}", response_model=MaterialRead)
def get_material(material_id: int, db: Session = Depends(get_db)):
    m = db.get(RawMaterial, material_id)
    if not m or m.deleted_at:
        raise HTTPException(404, "Material not found")
    return _enrich(db, m)


@router.post("", response_model=MaterialRead, status_code=201)
def create_material(body: MaterialCreate, db: Session = Depends(get_db)):
    m = RawMaterial(**body.model_dump())
    db.add(m)
    db.commit()
    db.refresh(m)
    return _enrich(db, m)


@router.put("/{material_id}", response_model=MaterialRead)
def update_material(material_id: int, body: MaterialUpdate, db: Session = Depends(get_db)):
    m = db.get(RawMaterial, material_id)
    if not m or m.deleted_at:
        raise HTTPException(404, "Material not found")
    for k, v in body.model_dump(exclude_unset=True).items():
        setattr(m, k, v)
    db.commit()
    db.refresh(m)
    return _enrich(db, m)


@router.delete("/{material_id}", status_code=204)
def delete_material(material_id: int, db: Session = Depends(get_db)):
    from datetime import datetime
    m = db.get(RawMaterial, material_id)
    if not m or m.deleted_at:
        raise HTTPException(404, "Material not found")
    m.deleted_at = datetime.utcnow()
    db.commit()
