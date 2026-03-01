"""Reservation execution logic – handles stock updates and audit trail."""

from datetime import datetime

from sqlalchemy.orm import Session

from backend.models.lot import Lot
from backend.models.reservation import Reservation, ReservationUsage
from backend.models.stock_transaction import StockTransaction
from backend.schemas.reservation import ReservationExecute


class InsufficientStockError(Exception):
    pass


class ReservationNotFoundError(Exception):
    pass


def execute_reservation(db: Session, reservation_id: int, payload: ReservationExecute) -> Reservation:
    """Execute a pending reservation, updating lot weights and creating audit records."""
    reservation = db.get(Reservation, reservation_id)
    if not reservation:
        raise ReservationNotFoundError(f"Reservation {reservation_id} not found")
    if reservation.status != "pending":
        raise ValueError(f"Reservation {reservation_id} is already {reservation.status}")

    actual_qty = payload.actual_quantity

    if reservation.type == "use":
        _execute_use(db, reservation, payload)
    else:
        _execute_replenish(db, reservation, payload)

    # Update reservation
    reservation.status = "executed"
    reservation.actual_quantity = actual_qty
    reservation.executed_at = datetime.utcnow()

    # Create audit transaction
    txn = StockTransaction(
        material_id=reservation.material_id,
        lot_id=reservation.lot_id,
        type=reservation.type,
        quantity=actual_qty,
        reservation_id=reservation.id,
        executed_by=payload.executed_by,
        note=payload.note,
    )
    db.add(txn)
    db.commit()
    db.refresh(reservation)
    return reservation


def _execute_use(db: Session, reservation: Reservation, payload: ReservationExecute) -> None:
    """Subtract stock from lots."""
    if payload.lot_usages:
        # Multi-lot usage
        for usage in payload.lot_usages:
            lot = db.get(Lot, usage.lot_id)
            if not lot or lot.deleted_at:
                raise ReservationNotFoundError(f"Lot {usage.lot_id} not found")
            if lot.weight < usage.quantity:
                raise InsufficientStockError(
                    f"Lot {lot.lot_name}: {lot.weight}g available, {usage.quantity}g requested"
                )
            lot.weight -= usage.quantity
            db.add(ReservationUsage(
                reservation_id=reservation.id,
                lot_id=lot.id,
                quantity=usage.quantity,
            ))
    elif reservation.lot_id:
        # Single lot
        lot = db.get(Lot, reservation.lot_id)
        if not lot or lot.deleted_at:
            raise ReservationNotFoundError("Target lot not found")
        if lot.weight < payload.actual_quantity:
            raise InsufficientStockError(
                f"Lot {lot.lot_name}: {lot.weight}g available, {payload.actual_quantity}g requested"
            )
        lot.weight -= payload.actual_quantity
    else:
        # Auto-deduct from oldest non-fraction lots (FIFO)
        _auto_deduct_fifo(db, reservation.material_id, payload.actual_quantity, reservation.id)


def _execute_replenish(db: Session, reservation: Reservation, payload: ReservationExecute) -> None:
    """Add stock to a lot (existing or new)."""
    from backend.models.master import Tare

    # Payload lot_id overrides reservation lot_id
    target_lot_id = payload.lot_id if payload.lot_id is not None else reservation.lot_id

    if target_lot_id:
        lot = db.get(Lot, target_lot_id)
        if lot and not lot.deleted_at:
            lot.weight += payload.actual_quantity
            reservation.lot_id = lot.id
            return

    # Create new lot
    lot_name = payload.lot_name or reservation.lot_name or f"補充-{datetime.utcnow().strftime('%Y%m%d-%H%M')}"
    weight = payload.actual_quantity
    if payload.tare_id:
        tare = db.get(Tare, payload.tare_id)
        if tare:
            weight = max(0, weight - tare.weight)

    new_lot = Lot(
        material_id=reservation.material_id,
        lot_name=lot_name,
        weight=weight,
    )
    db.add(new_lot)
    db.flush()
    reservation.lot_id = new_lot.id


def _auto_deduct_fifo(db: Session, material_id: int, quantity: float, reservation_id: int) -> None:
    """FIFO deduction across lots."""
    lots = (
        db.query(Lot)
        .filter(
            Lot.material_id == material_id,
            Lot.is_fraction == False,  # noqa: E712
            Lot.deleted_at == None,  # noqa: E711
            Lot.weight > 0,
        )
        .order_by(Lot.created_at)
        .all()
    )

    remaining = quantity
    for lot in lots:
        if remaining <= 0:
            break
        deduct = min(lot.weight, remaining)
        lot.weight -= deduct
        remaining -= deduct
        db.add(ReservationUsage(
            reservation_id=reservation_id,
            lot_id=lot.id,
            quantity=deduct,
        ))

    if remaining > 0:
        raise InsufficientStockError(f"在庫不足: {remaining}g 足りません")
