"""Dashboard & statistics endpoints."""

from fastapi import APIRouter, Depends, Query
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.material import RawMaterial
from backend.models.lot import Lot
from backend.models.reservation import Reservation
from backend.models.work_order import WorkOrder
from backend.schemas.dashboard import (
    CriticalPeriod,
    DashboardSummary,
    MaterialStats,
    StockTimelineEvent,
)
from backend.services import stock_calculator

router = APIRouter(prefix="/dashboard", tags=["dashboard"])


@router.get("/summary", response_model=DashboardSummary)
def get_summary(db: Session = Depends(get_db)):
    materials = db.query(RawMaterial).filter(RawMaterial.deleted_at == None).all()  # noqa: E711
    total_lots = db.query(Lot).filter(Lot.deleted_at == None).count()  # noqa: E711

    low_count = 0
    for m in materials:
        if stock_calculator.is_low_stock(db, m.id, m.min_weight):
            low_count += 1

    pending_res = db.query(Reservation).filter(Reservation.status == "pending").count()
    overdue_res = sum(1 for r in db.query(Reservation).filter(Reservation.status == "pending").all() if r.is_overdue)

    active_wo = db.query(WorkOrder).filter(
        WorkOrder.status.in_(["pending", "in_progress"]),
        WorkOrder.archived == False,  # noqa: E712
    ).count()
    overdue_wo = sum(
        1 for wo in db.query(WorkOrder).filter(
            WorkOrder.status.in_(["pending", "in_progress"]),
            WorkOrder.archived == False,  # noqa: E712
        ).all()
        if wo.is_overdue
    )

    return DashboardSummary(
        total_materials=len(materials),
        total_lots=total_lots,
        low_stock_count=low_count,
        pending_reservations=pending_res,
        overdue_reservations=overdue_res,
        active_work_orders=active_wo,
        overdue_work_orders=overdue_wo,
    )


@router.get("/alerts", response_model=list[CriticalPeriod])
def get_alerts(db: Session = Depends(get_db)):
    materials = db.query(RawMaterial).filter(RawMaterial.deleted_at == None).all()  # noqa: E711
    all_periods: list[CriticalPeriod] = []
    for m in materials:
        if m.min_weight > 0:
            periods = stock_calculator.critical_periods(db, m.id, m.min_weight)
            all_periods.extend(periods)
    return all_periods


@router.get("/materials/{material_id}/stats", response_model=list[MaterialStats])
def get_material_stats(material_id: int, db: Session = Depends(get_db)):
    periods = [1, 7, 30, 90, 180, 365]
    return [stock_calculator.material_stats(db, material_id, d) for d in periods]


@router.get("/materials/{material_id}/timeline", response_model=list[StockTimelineEvent])
def get_material_timeline(material_id: int, db: Session = Depends(get_db)):
    return stock_calculator.stock_timeline(db, material_id)
