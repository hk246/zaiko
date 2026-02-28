"""StockTransaction model – immutable audit trail of all stock changes."""

from datetime import datetime

from sqlalchemy import DateTime, Float, ForeignKey, String, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class StockTransaction(Base):
    __tablename__ = "stock_transactions"

    id: Mapped[int] = mapped_column(primary_key=True)
    material_id: Mapped[int] = mapped_column(ForeignKey("raw_materials.id"))
    lot_id: Mapped[int | None] = mapped_column(ForeignKey("lots.id"), nullable=True)
    type: Mapped[str] = mapped_column(String(10))  # 'use' | 'replenish'
    quantity: Mapped[float] = mapped_column(Float)
    reservation_id: Mapped[int | None] = mapped_column(ForeignKey("reservations.id"), nullable=True)
    executed_by: Mapped[str | None] = mapped_column(String(100), nullable=True)
    note: Mapped[str | None] = mapped_column(Text, nullable=True)
    executed_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    material = relationship("RawMaterial", back_populates="transactions")
    lot = relationship("Lot", back_populates="transactions")
