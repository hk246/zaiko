"""Pydantic schemas for Recipe."""

from datetime import datetime

from pydantic import BaseModel, Field


class RecipeItemCreate(BaseModel):
    material_id: int
    quantity: float = Field(gt=0)
    lot_name: str | None = None


class RecipeCreate(BaseModel):
    name: str = Field(max_length=100)
    description: str | None = None
    type: str = Field(default="use")
    items: list[RecipeItemCreate]


class RecipeUpdate(BaseModel):
    name: str | None = Field(default=None, max_length=100)
    description: str | None = None
    type: str | None = None
    items: list[RecipeItemCreate] | None = None


class RecipeItemRead(BaseModel):
    id: int
    material_id: int
    material_name: str = ""
    quantity: float
    lot_name: str | None

    model_config = {"from_attributes": True}


class RecipeRead(BaseModel):
    id: int
    name: str
    description: str | None
    type: str
    items: list[RecipeItemRead] = []
    created_at: datetime

    model_config = {"from_attributes": True}


class RecipeExecute(BaseModel):
    """Execute all items in a recipe as reservations."""
    scheduled_date: str  # ISO date
    user_name: str | None = None
    purpose: str | None = None
