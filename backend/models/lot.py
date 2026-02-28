"""Lot model – batch tracking for materials."""

from datetime import datetime

from sqlalchemy import Boolean, DateTime, Float, ForeignKey, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class Lot(Base):
    __tablename__ = "lots"

    id: Mapped[int] = mapped_column(primary_key=True)
    material_id: Mapped[int] = mapped_column(ForeignKey("raw_materials.id"))
    lot_name: Mapped[str] = mapped_column(String(100))
    weight: Mapped[float] = mapped_column(Float, default=0.0)
    is_fraction: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    deleted_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    # relationships
    material = relationship("RawMaterial", back_populates="lots")
    properties = relationship("LotProperty", back_populates="lot", cascade="all, delete-orphan")
    reservation_usages = relationship("ReservationUsage", back_populates="lot")
    transactions = relationship("StockTransaction", back_populates="lot")
