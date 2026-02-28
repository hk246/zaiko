"""Tare & Contact CRUD endpoints."""

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.master import Contact, Tare
from backend.schemas.master import (
    ContactCreate,
    ContactRead,
    ContactUpdate,
    TareCreate,
    TareRead,
    TareUpdate,
)

router = APIRouter(tags=["master"])

# ── Tares ──


@router.get("/tares", response_model=list[TareRead])
def list_tares(db: Session = Depends(get_db)):
    return db.query(Tare).order_by(Tare.name).all()


@router.post("/tares", response_model=TareRead, status_code=201)
def create_tare(body: TareCreate, db: Session = Depends(get_db)):
    t = Tare(**body.model_dump())
    db.add(t)
    db.commit()
    db.refresh(t)
    return t


@router.put("/tares/{tare_id}", response_model=TareRead)
def update_tare(tare_id: int, body: TareUpdate, db: Session = Depends(get_db)):
    t = db.get(Tare, tare_id)
    if not t:
        raise HTTPException(404, "Tare not found")
    for k, v in body.model_dump(exclude_unset=True).items():
        setattr(t, k, v)
    db.commit()
    db.refresh(t)
    return t


@router.delete("/tares/{tare_id}", status_code=204)
def delete_tare(tare_id: int, db: Session = Depends(get_db)):
    t = db.get(Tare, tare_id)
    if not t:
        raise HTTPException(404, "Tare not found")
    db.delete(t)
    db.commit()


# ── Contacts ──


@router.get("/contacts", response_model=list[ContactRead])
def list_contacts(db: Session = Depends(get_db)):
    return db.query(Contact).order_by(Contact.name).all()


@router.post("/contacts", response_model=ContactRead, status_code=201)
def create_contact(body: ContactCreate, db: Session = Depends(get_db)):
    c = Contact(**body.model_dump())
    db.add(c)
    db.commit()
    db.refresh(c)
    return c


@router.put("/contacts/{contact_id}", response_model=ContactRead)
def update_contact(contact_id: int, body: ContactUpdate, db: Session = Depends(get_db)):
    c = db.get(Contact, contact_id)
    if not c:
        raise HTTPException(404, "Contact not found")
    for k, v in body.model_dump(exclude_unset=True).items():
        setattr(c, k, v)
    db.commit()
    db.refresh(c)
    return c


@router.delete("/contacts/{contact_id}", status_code=204)
def delete_contact(contact_id: int, db: Session = Depends(get_db)):
    c = db.get(Contact, contact_id)
    if not c:
        raise HTTPException(404, "Contact not found")
    db.delete(c)
    db.commit()
