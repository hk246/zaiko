"""Label CRUD endpoints."""

from datetime import datetime

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.label import MaterialLabel
from backend.schemas.label import LabelCreate, LabelRead, LabelUpdate

router = APIRouter(prefix="/labels", tags=["labels"])


def _to_read(label: MaterialLabel) -> dict:
    active = [m for m in label.materials if not m.deleted_at]
    return {
        **{c.name: getattr(label, c.name) for c in label.__table__.columns},
        "material_count": len(active),
    }


@router.get("", response_model=list[LabelRead])
def list_labels(db: Session = Depends(get_db)):
    labels = (
        db.query(MaterialLabel)
        .filter(MaterialLabel.deleted_at == None)  # noqa: E711
        .order_by(MaterialLabel.name)
        .all()
    )
    return [_to_read(l) for l in labels]


@router.post("", response_model=LabelRead, status_code=201)
def create_label(body: LabelCreate, db: Session = Depends(get_db)):
    existing = db.query(MaterialLabel).filter(MaterialLabel.name == body.name, MaterialLabel.deleted_at == None).first()  # noqa: E711
    if existing:
        raise HTTPException(400, "Label name already exists")
    label = MaterialLabel(**body.model_dump())
    db.add(label)
    db.commit()
    db.refresh(label)
    return _to_read(label)


@router.put("/{label_id}", response_model=LabelRead)
def update_label(label_id: int, body: LabelUpdate, db: Session = Depends(get_db)):
    label = db.get(MaterialLabel, label_id)
    if not label or label.deleted_at:
        raise HTTPException(404, "Label not found")
    for k, v in body.model_dump(exclude_unset=True).items():
        setattr(label, k, v)
    db.commit()
    db.refresh(label)
    return _to_read(label)


@router.delete("/{label_id}", status_code=204)
def delete_label(label_id: int, db: Session = Depends(get_db)):
    label = db.get(MaterialLabel, label_id)
    if not label or label.deleted_at:
        raise HTTPException(404, "Label not found")
    label.deleted_at = datetime.utcnow()
    db.commit()
