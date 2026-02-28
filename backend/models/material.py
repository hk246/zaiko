"""RawMaterial model – core inventory item."""

from datetime import datetime

from sqlalchemy import DateTime, Float, ForeignKey, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class RawMaterial(Base):
    __tablename__ = "raw_materials"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100))
    unit: Mapped[str] = mapped_column(String(10), default="g")
    min_weight: Mapped[float] = mapped_column(Float, default=0.0)
    email: Mapped[str | None] = mapped_column(String(120), nullable=True)
    excel_path: Mapped[str | None] = mapped_column(String(500), nullable=True)
    action_type: Mapped[str] = mapped_column(String(20), default="none")  # none | email | excel
    contact_id: Mapped[int | None] = mapped_column(ForeignKey("contacts.id", ondelete="SET NULL"), nullable=True)
    label_id: Mapped[int | None] = mapped_column(ForeignKey("material_labels.id"), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    deleted_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    # relationships
    label = relationship("MaterialLabel", back_populates="materials")
    contact = relationship("Contact", foreign_keys=[contact_id])
    lots = relationship("Lot", back_populates="material", cascade="all, delete-orphan")
    reservations = relationship("Reservation", back_populates="material", cascade="all, delete-orphan")
    transactions = relationship("StockTransaction", back_populates="material")
