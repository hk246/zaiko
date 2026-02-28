"""Pydantic schemas for WorkOrder."""

from datetime import date, datetime

from pydantic import BaseModel, Field


class WorkOrderCreate(BaseModel):
    process_name: str = Field(max_length=200)
    lot_name: str | None = None
    material_id: int | None = None
    priority: str = Field(default="none")
    invoice_issued: bool = False
    worker_name: str | None = None
    experiment_type: str = Field(default="standard")
    planned_start: date | None = None
    planned_end: date | None = None
    actual_start: date | None = None
    actual_end: date | None = None
    notes: str | None = None


class WorkOrderUpdate(BaseModel):
    process_name: str | None = Field(default=None, max_length=200)
    lot_name: str | None = None
    material_id: int | None = None
    status: str | None = None
    priority: str | None = None
    invoice_issued: bool | None = None
    worker_name: str | None = None
    experiment_type: str | None = None
    planned_start: date | None = None
    planned_end: date | None = None
    actual_start: date | None = None
    actual_end: date | None = None
    notes: str | None = None


class WorkOrderRead(BaseModel):
    id: int
    process_name: str
    lot_name: str | None
    material_id: int | None
    material_name: str | None = None
    status: str
    priority: str
    invoice_issued: bool
    worker_name: str | None
    experiment_type: str
    planned_start: date | None
    planned_end: date | None
    actual_start: date | None
    actual_end: date | None
    notes: str | None
    archived: bool
    created_at: datetime
    is_overdue: bool = False
    progress_pct: int = 0

    model_config = {"from_attributes": True}
