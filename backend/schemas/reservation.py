"""Pydantic schemas for Reservation."""

from datetime import date, datetime

from pydantic import BaseModel, Field


class ReservationCreate(BaseModel):
    material_id: int
    lot_id: int | None = None
    lot_name: str | None = None
    type: str  # 'use' | 'replenish'
    quantity: float = Field(gt=0)
    user_name: str | None = None
    purpose: str | None = None
    scheduled_date: date


class ReservationUpdate(BaseModel):
    lot_id: int | None = None
    lot_name: str | None = None
    quantity: float | None = Field(default=None, gt=0)
    user_name: str | None = None
    purpose: str | None = None
    scheduled_date: date | None = None


class ReservationExecute(BaseModel):
    """Payload for executing a reservation."""
    actual_quantity: float = Field(gt=0)
    lot_usages: list["LotUsageItem"] = []
    executed_by: str | None = None
    note: str | None = None


class LotUsageItem(BaseModel):
    lot_id: int
    quantity: float = Field(gt=0)


class ReservationRead(BaseModel):
    id: int
    material_id: int
    material_name: str = ""
    lot_id: int | None
    lot_name: str | None
    recipe_id: int | None
    type: str
    quantity: float
    actual_quantity: float | None
    user_name: str | None
    purpose: str | None
    scheduled_date: date
    status: str
    executed_at: datetime | None
    created_at: datetime
    is_overdue: bool = False

    model_config = {"from_attributes": True}
