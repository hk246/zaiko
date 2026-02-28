"""Reservation & ReservationUsage models."""

from datetime import date, datetime

from sqlalchemy import Boolean, Date, DateTime, Float, ForeignKey, String, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class Reservation(Base):
    __tablename__ = "reservations"

    id: Mapped[int] = mapped_column(primary_key=True)
    material_id: Mapped[int] = mapped_column(ForeignKey("raw_materials.id"))
    lot_id: Mapped[int | None] = mapped_column(ForeignKey("lots.id"), nullable=True)
    lot_name: Mapped[str | None] = mapped_column(String(100), nullable=True)
    recipe_id: Mapped[int | None] = mapped_column(ForeignKey("recipes.id"), nullable=True)
    type: Mapped[str] = mapped_column(String(10))  # 'use' | 'replenish'
    quantity: Mapped[float] = mapped_column(Float)
    actual_quantity: Mapped[float | None] = mapped_column(Float, nullable=True)
    user_name: Mapped[str | None] = mapped_column(String(100), nullable=True)
    purpose: Mapped[str | None] = mapped_column(Text, nullable=True)
    scheduled_date: Mapped[date] = mapped_column(Date)
    status: Mapped[str] = mapped_column(String(20), default="pending")  # pending | executed | cancelled
    executed_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    # relationships
    material = relationship("RawMaterial", back_populates="reservations")
    lot = relationship("Lot", foreign_keys=[lot_id])
    recipe = relationship("Recipe", back_populates="reservations")
    usages = relationship("ReservationUsage", back_populates="reservation", cascade="all, delete-orphan")

    @property
    def is_overdue(self) -> bool:
        return self.status == "pending" and self.scheduled_date < date.today()


class ReservationUsage(Base):
    """Tracks which lots were consumed when executing a multi-lot reservation."""

    __tablename__ = "reservation_usages"

    id: Mapped[int] = mapped_column(primary_key=True)
    reservation_id: Mapped[int] = mapped_column(ForeignKey("reservations.id"))
    lot_id: Mapped[int] = mapped_column(ForeignKey("lots.id"))
    quantity: Mapped[float] = mapped_column(Float)

    reservation = relationship("Reservation", back_populates="usages")
    lot = relationship("Lot", back_populates="reservation_usages")
