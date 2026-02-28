"""Pydantic schemas for Dashboard & Statistics."""

from datetime import date

from pydantic import BaseModel


class DashboardSummary(BaseModel):
    total_materials: int
    total_lots: int
    low_stock_count: int
    pending_reservations: int
    overdue_reservations: int
    active_work_orders: int
    overdue_work_orders: int


class CriticalPeriod(BaseModel):
    material_id: int
    material_name: str
    start_date: date
    end_date: date | None
    shortage_amount: float
    contact_id: int | None = None
    contact_name: str | None = None
    contact_email: str | None = None
    action_type: str = "none"
    current_stock: float = 0.0


class MaterialStats(BaseModel):
    period_label: str
    total_used: float
    total_replenished: float
    net_change: float
    use_count: int
    replenish_count: int


class StockTimelineEvent(BaseModel):
    date: date
    stock: float
    event_type: str | None = None  # 'use' | 'replenish' | None
    description: str = ""
