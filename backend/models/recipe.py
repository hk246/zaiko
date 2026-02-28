"""Recipe & RecipeItem models."""

from datetime import datetime

from sqlalchemy import DateTime, Float, ForeignKey, String, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from backend.database import Base


class Recipe(Base):
    __tablename__ = "recipes"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100))
    description: Mapped[str | None] = mapped_column(Text, nullable=True)
    type: Mapped[str] = mapped_column(String(10), default="use")  # 'use' | 'replenish'
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    deleted_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    items = relationship("RecipeItem", back_populates="recipe", cascade="all, delete-orphan")
    reservations = relationship("Reservation", back_populates="recipe")


class RecipeItem(Base):
    __tablename__ = "recipe_items"

    id: Mapped[int] = mapped_column(primary_key=True)
    recipe_id: Mapped[int] = mapped_column(ForeignKey("recipes.id"))
    material_id: Mapped[int] = mapped_column(ForeignKey("raw_materials.id"))
    quantity: Mapped[float] = mapped_column(Float)
    lot_name: Mapped[str | None] = mapped_column(String(100), nullable=True)

    recipe = relationship("Recipe", back_populates="items")
    material = relationship("RawMaterial")
