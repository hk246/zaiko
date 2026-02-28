"""MaterialLabel model – categories for materials."""

from datetime import datetime

from sqlalchemy import DateTime, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class MaterialLabel(Base):
    __tablename__ = "material_labels"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(50), unique=True)
    color: Mapped[str] = mapped_column(String(7), default="#6c757d")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    deleted_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    # relationships
    materials = relationship("RawMaterial", back_populates="label")
    property_fields = relationship("PropertyField", back_populates="label")
