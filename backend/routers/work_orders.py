"""Work order CRUD + status management endpoints."""

from datetime import datetime

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.work_order import WorkOrder
from backend.schemas.work_order import WorkOrderCreate, WorkOrderRead, WorkOrderUpdate

router = APIRouter(prefix="/work-orders", tags=["work-orders"])


def _to_read(wo: WorkOrder) -> dict:
    return {
        **{c.name: getattr(wo, c.name) for c in wo.__table__.columns},
        "material_name": wo.material.name if wo.material else None,
        "is_overdue": wo.is_overdue,
        "progress_pct": wo.progress_pct,
    }


@router.get("", response_model=list[WorkOrderRead])
def list_work_orders(
    status: str | None = Query(default=None),
    priority: str | None = Query(default=None),
    archived: bool = Query(default=False),
    db: Session = Depends(get_db),
):
    query = db.query(WorkOrder).filter(WorkOrder.archived == archived)
    if status:
        query = query.filter(WorkOrder.status == status)
    if priority:
        query = query.filter(WorkOrder.priority == priority)
    work_orders = query.order_by(WorkOrder.planned_start.desc().nullslast()).all()

    # Auto-compute status
    for wo in work_orders:
        new_status = wo.compute_status()
        if new_status != wo.status:
            wo.status = new_status
    db.commit()

    return [_to_read(wo) for wo in work_orders]


@router.get("/{wo_id}", response_model=WorkOrderRead)
def get_work_order(wo_id: int, db: Session = Depends(get_db)):
    wo = db.get(WorkOrder, wo_id)
    if not wo:
        raise HTTPException(404, "Work order not found")
    return _to_read(wo)


@router.post("", response_model=WorkOrderRead, status_code=201)
def create_work_order(body: WorkOrderCreate, db: Session = Depends(get_db)):
    wo = WorkOrder(**body.model_dump())
    wo.status = wo.compute_status()
    db.add(wo)
    db.commit()
    db.refresh(wo)
    return _to_read(wo)


@router.put("/{wo_id}", response_model=WorkOrderRead)
def update_work_order(wo_id: int, body: WorkOrderUpdate, db: Session = Depends(get_db)):
    wo = db.get(WorkOrder, wo_id)
    if not wo:
        raise HTTPException(404, "Work order not found")
    for k, v in body.model_dump(exclude_unset=True).items():
        setattr(wo, k, v)
    wo.status = wo.compute_status()
    db.commit()
    db.refresh(wo)
    return _to_read(wo)


@router.delete("/{wo_id}", status_code=204)
def delete_work_order(wo_id: int, db: Session = Depends(get_db)):
    wo = db.get(WorkOrder, wo_id)
    if not wo:
        raise HTTPException(404, "Work order not found")
    db.delete(wo)
    db.commit()


@router.patch("/{wo_id}/status", response_model=WorkOrderRead)
def update_status(wo_id: int, status: str = Query(...), db: Session = Depends(get_db)):
    wo = db.get(WorkOrder, wo_id)
    if not wo:
        raise HTTPException(404, "Work order not found")
    wo.status = status
    db.commit()
    db.refresh(wo)
    return _to_read(wo)


@router.post("/{wo_id}/archive", response_model=WorkOrderRead)
def archive(wo_id: int, db: Session = Depends(get_db)):
    wo = db.get(WorkOrder, wo_id)
    if not wo:
        raise HTTPException(404, "Work order not found")
    wo.archived = True
    db.commit()
    db.refresh(wo)
    return _to_read(wo)


@router.post("/{wo_id}/unarchive", response_model=WorkOrderRead)
def unarchive(wo_id: int, db: Session = Depends(get_db)):
    wo = db.get(WorkOrder, wo_id)
    if not wo:
        raise HTTPException(404, "Work order not found")
    wo.archived = False
    db.commit()
    db.refresh(wo)
    return _to_read(wo)


@router.post("/archive-completed")
def archive_completed(db: Session = Depends(get_db)):
    work_orders = (
        db.query(WorkOrder)
        .filter(WorkOrder.status == "completed", WorkOrder.archived == False)  # noqa: E712
        .all()
    )
    for wo in work_orders:
        wo.archived = True
    db.commit()
    return {"archived_count": len(work_orders)}
