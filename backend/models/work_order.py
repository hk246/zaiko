"""WorkOrder model – process / manufacturing tracking."""

from datetime import date, datetime

from sqlalchemy import Boolean, Date, DateTime, ForeignKey, String, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class WorkOrder(Base):
    __tablename__ = "work_orders"

    id: Mapped[int] = mapped_column(primary_key=True)
    process_name: Mapped[str] = mapped_column(String(200))
    lot_name: Mapped[str | None] = mapped_column(String(100), nullable=True)
    material_id: Mapped[int | None] = mapped_column(ForeignKey("raw_materials.id"), nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="pending")  # pending | in_progress | completed
    priority: Mapped[str] = mapped_column(String(20), default="none")  # critical | high | low | none
    invoice_issued: Mapped[bool] = mapped_column(Boolean, default=False)
    worker_name: Mapped[str | None] = mapped_column(String(100), nullable=True)
    experiment_type: Mapped[str] = mapped_column(String(20), default="standard")  # standard | experiment
    planned_start: Mapped[date | None] = mapped_column(Date, nullable=True)
    planned_end: Mapped[date | None] = mapped_column(Date, nullable=True)
    actual_start: Mapped[date | None] = mapped_column(Date, nullable=True)
    actual_end: Mapped[date | None] = mapped_column(Date, nullable=True)
    notes: Mapped[str | None] = mapped_column(Text, nullable=True)
    archived: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    material = relationship("RawMaterial")

    def compute_status(self) -> str:
        today = date.today()
        if self.actual_end and self.actual_end <= today:
            return "completed"
        if self.actual_start:
            return "in_progress"
        return "pending"

    @property
    def is_overdue(self) -> bool:
        if self.status == "completed" or not self.planned_end:
            return False
        return date.today() > self.planned_end

    @property
    def progress_pct(self) -> int:
        if self.status == "completed":
            return 100
        if not self.actual_start or not self.planned_end:
            return 0
        today = date.today()
        total = (self.planned_end - self.actual_start).days
        if total <= 0:
            return 0
        elapsed = (today - self.actual_start).days
        return min(100, max(0, int(elapsed / total * 100)))
