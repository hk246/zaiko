"""Stock calculation engine – single source of truth for inventory levels."""

from datetime import date, timedelta

from sqlalchemy import func, and_
from sqlalchemy.orm import Session

from backend.models.lot import Lot
from backend.models.reservation import Reservation
from backend.models.stock_transaction import StockTransaction
from backend.schemas.dashboard import CriticalPeriod, MaterialStats, StockTimelineEvent


def current_stock(db: Session, material_id: int) -> float:
    """Sum of lot weights, excluding fraction lots."""
    result = db.query(func.coalesce(func.sum(Lot.weight), 0.0)).filter(
        Lot.material_id == material_id,
        Lot.is_fraction == False,  # noqa: E712
        Lot.deleted_at == None,  # noqa: E711
    ).scalar()
    return float(result)


def fraction_stock(db: Session, material_id: int) -> float:
    """Sum of fraction lot weights."""
    result = db.query(func.coalesce(func.sum(Lot.weight), 0.0)).filter(
        Lot.material_id == material_id,
        Lot.is_fraction == True,  # noqa: E712
        Lot.deleted_at == None,  # noqa: E711
    ).scalar()
    return float(result)


def predicted_stock(db: Session, material_id: int) -> float:
    """Current stock + pending replenish - pending use."""
    stock = current_stock(db, material_id)

    pending_use = db.query(func.coalesce(func.sum(Reservation.quantity), 0.0)).filter(
        Reservation.material_id == material_id,
        Reservation.type == "use",
        Reservation.status == "pending",
    ).scalar()

    pending_replenish = db.query(func.coalesce(func.sum(Reservation.quantity), 0.0)).filter(
        Reservation.material_id == material_id,
        Reservation.type == "replenish",
        Reservation.status == "pending",
    ).scalar()

    return stock + float(pending_replenish) - float(pending_use)


def critical_periods(db: Session, material_id: int, min_weight: float) -> list[CriticalPeriod]:
    """Find date ranges where stock falls below minimum.

    Covers two cases:
    1. Current stock is ALREADY below min_weight (start_date = today)
    2. Future reservations will cause stock to drop below min_weight
    """
    from backend.models.material import RawMaterial

    material = db.get(RawMaterial, material_id)
    if not material:
        return []

    stock = current_stock(db, material_id)

    # Contact info for email alerts
    _contact_id = material.contact_id if hasattr(material, "contact_id") else None
    _contact_name = material.contact.name if (hasattr(material, "contact") and material.contact) else None
    _contact_email = material.contact.email if (hasattr(material, "contact") and material.contact) else None
    _action_type = material.action_type if hasattr(material, "action_type") else "none"

    # Get pending reservations ordered by date
    reservations = (
        db.query(Reservation)
        .filter(
            Reservation.material_id == material_id,
            Reservation.status == "pending",
        )
        .order_by(Reservation.scheduled_date)
        .all()
    )

    periods: list[CriticalPeriod] = []
    running_stock = stock
    # Already below minimum before any reservation?
    in_critical = stock < min_weight
    critical_start: date | None = date.today() if in_critical else None

    if not reservations:
        # No reservations but already below minimum → immediate alert
        if in_critical:
            periods.append(CriticalPeriod(
                material_id=material_id,
                material_name=material.name,
                start_date=date.today(),
                end_date=None,
                shortage_amount=min_weight - stock,
                contact_id=_contact_id,
                contact_name=_contact_name,
                contact_email=_contact_email,
                action_type=_action_type,
                current_stock=stock,
            ))
        return periods

    for res in reservations:
        if res.type == "use":
            running_stock -= res.quantity
        else:
            running_stock += res.quantity

        if running_stock < min_weight and not in_critical:
            in_critical = True
            critical_start = res.scheduled_date
        elif running_stock >= min_weight and in_critical:
            in_critical = False
            periods.append(CriticalPeriod(
                material_id=material_id,
                material_name=material.name,
                start_date=critical_start,  # type: ignore[arg-type]
                end_date=res.scheduled_date,
                shortage_amount=min_weight - running_stock,
                contact_id=_contact_id,
                contact_name=_contact_name,
                contact_email=_contact_email,
                action_type=_action_type,
                current_stock=stock,
            ))

    # Still in critical at end of reservations
    if in_critical:
        periods.append(CriticalPeriod(
            material_id=material_id,
            material_name=material.name,
            start_date=critical_start,  # type: ignore[arg-type]
            end_date=None,
            shortage_amount=min_weight - running_stock,
            contact_id=_contact_id,
            contact_name=_contact_name,
            contact_email=_contact_email,
            action_type=_action_type,
            current_stock=stock,
        ))

    return periods


def is_low_stock(db: Session, material_id: int, min_weight: float) -> bool:
    """Quick check: is stock currently below minimum, or will it drop?"""
    if min_weight <= 0:
        return False
    # Fast path: check current stock directly
    stock = current_stock(db, material_id)
    if stock < min_weight:
        return True
    # Check future via reservations
    return len(critical_periods(db, material_id, min_weight)) > 0


def stock_timeline(db: Session, material_id: int) -> list[StockTimelineEvent]:
    """Build a timeline of stock changes for graphing."""
    stock = current_stock(db, material_id)

    events: list[StockTimelineEvent] = [
        StockTimelineEvent(date=date.today(), stock=stock, description="現在の在庫")
    ]

    reservations = (
        db.query(Reservation)
        .filter(
            Reservation.material_id == material_id,
            Reservation.status == "pending",
        )
        .order_by(Reservation.scheduled_date)
        .all()
    )

    running = stock
    for res in reservations:
        if res.type == "use":
            running -= res.quantity
        else:
            running += res.quantity
        events.append(StockTimelineEvent(
            date=res.scheduled_date,
            stock=running,
            event_type=res.type,
            description=f"{'使用' if res.type == 'use' else '補充'}: {res.quantity}",
        ))

    return events


def material_stats(db: Session, material_id: int, days: int) -> MaterialStats:
    """Usage/replenishment stats for a given period."""
    since = date.today() - timedelta(days=days)

    base = db.query(StockTransaction).filter(
        StockTransaction.material_id == material_id,
        StockTransaction.executed_at >= since,
    )

    use_txns = base.filter(StockTransaction.type == "use").all()
    replenish_txns = base.filter(StockTransaction.type == "replenish").all()

    total_used = sum(t.quantity for t in use_txns)
    total_replenished = sum(t.quantity for t in replenish_txns)

    period_labels = {1: "1日", 7: "7日", 30: "30日", 90: "90日", 180: "180日", 365: "1年"}

    return MaterialStats(
        period_label=period_labels.get(days, f"{days}日"),
        total_used=total_used,
        total_replenished=total_replenished,
        net_change=total_replenished - total_used,
        use_count=len(use_txns),
        replenish_count=len(replenish_txns),
    )
