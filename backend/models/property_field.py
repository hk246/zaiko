"""PropertyField & LotProperty models – custom attributes per lot."""

from datetime import datetime

from sqlalchemy import DateTime, Float, ForeignKey, Integer, String, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class PropertyField(Base):
    __tablename__ = "property_fields"

    id: Mapped[int] = mapped_column(primary_key=True)
    label_id: Mapped[int | None] = mapped_column(ForeignKey("material_labels.id"), nullable=True)
    name: Mapped[str] = mapped_column(String(100))
    field_type: Mapped[str] = mapped_column(String(10), default="string")  # string | number
    unit: Mapped[str | None] = mapped_column(String(30), nullable=True)
    order_index: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    label = relationship("MaterialLabel", back_populates="property_fields")
    values = relationship("LotProperty", back_populates="field", cascade="all, delete-orphan")


class LotProperty(Base):
    __tablename__ = "lot_properties"
    __table_args__ = (UniqueConstraint("lot_id", "field_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    lot_id: Mapped[int] = mapped_column(ForeignKey("lots.id"))
    field_id: Mapped[int] = mapped_column(ForeignKey("property_fields.id"))
    value_string: Mapped[str | None] = mapped_column(String(500), nullable=True)
    value_number: Mapped[float | None] = mapped_column(Float, nullable=True)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    lot = relationship("Lot", back_populates="properties")
    field = relationship("PropertyField", back_populates="values")
