"""Reservation CRUD + execution endpoints."""

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.reservation import Reservation
from backend.schemas.reservation import (
    ReservationCreate,
    ReservationExecute,
    ReservationRead,
    ReservationUpdate,
)
from backend.services.reservation_executor import (
    InsufficientStockError,
    ReservationNotFoundError,
    execute_reservation,
)

router = APIRouter(prefix="/reservations", tags=["reservations"])


def _to_read(r: Reservation) -> dict:
    return {
        **{c.name: getattr(r, c.name) for c in r.__table__.columns},
        "material_name": r.material.name if r.material else "",
        "is_overdue": r.is_overdue,
    }


@router.get("", response_model=list[ReservationRead])
def list_reservations(
    status: str = Query(default="pending"),
    type: str | None = Query(default=None),
    material_id: int | None = Query(default=None),
    db: Session = Depends(get_db),
):
    query = db.query(Reservation)
    if status:
        query = query.filter(Reservation.status == status)
    if type:
        query = query.filter(Reservation.type == type)
    if material_id:
        query = query.filter(Reservation.material_id == material_id)
    reservations = query.order_by(Reservation.scheduled_date).all()
    return [_to_read(r) for r in reservations]


@router.get("/{reservation_id}", response_model=ReservationRead)
def get_reservation(reservation_id: int, db: Session = Depends(get_db)):
    r = db.get(Reservation, reservation_id)
    if not r:
        raise HTTPException(404, "Reservation not found")
    return _to_read(r)


@router.post("", response_model=ReservationRead, status_code=201)
def create_reservation(body: ReservationCreate, db: Session = Depends(get_db)):
    r = Reservation(**body.model_dump())
    db.add(r)
    db.commit()
    db.refresh(r)
    return _to_read(r)


@router.put("/{reservation_id}", response_model=ReservationRead)
def update_reservation(reservation_id: int, body: ReservationUpdate, db: Session = Depends(get_db)):
    r = db.get(Reservation, reservation_id)
    if not r:
        raise HTTPException(404, "Reservation not found")
    if r.status != "pending":
        raise HTTPException(400, "Can only edit pending reservations")
    for k, v in body.model_dump(exclude_unset=True).items():
        setattr(r, k, v)
    db.commit()
    db.refresh(r)
    return _to_read(r)


@router.delete("/{reservation_id}", status_code=204)
def delete_reservation(reservation_id: int, db: Session = Depends(get_db)):
    r = db.get(Reservation, reservation_id)
    if not r:
        raise HTTPException(404, "Reservation not found")
    if r.status == "executed":
        raise HTTPException(400, "Cannot delete executed reservation")
    db.delete(r)
    db.commit()


@router.post("/{reservation_id}/execute", response_model=ReservationRead)
def execute(reservation_id: int, body: ReservationExecute, db: Session = Depends(get_db)):
    try:
        r = execute_reservation(db, reservation_id, body)
        return _to_read(r)
    except ReservationNotFoundError as e:
        raise HTTPException(404, str(e))
    except InsufficientStockError as e:
        raise HTTPException(400, str(e))
    except ValueError as e:
        raise HTTPException(400, str(e))


@router.post("/{reservation_id}/cancel", response_model=ReservationRead)
def cancel_reservation(reservation_id: int, db: Session = Depends(get_db)):
    r = db.get(Reservation, reservation_id)
    if not r:
        raise HTTPException(404, "Reservation not found")
    if r.status != "pending":
        raise HTTPException(400, "Can only cancel pending reservations")
    r.status = "cancelled"
    db.commit()
    db.refresh(r)
    return _to_read(r)
