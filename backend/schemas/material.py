"""Pydantic schemas for RawMaterial."""

from datetime import datetime

from pydantic import BaseModel, Field

from backend.schemas.label import LabelRead


class MaterialCreate(BaseModel):
    name: str = Field(max_length=100)
    unit: str = Field(default="g", max_length=10)
    min_weight: float = Field(default=0.0, ge=0)
    email: str | None = None
    excel_path: str | None = None
    action_type: str = Field(default="none")
    contact_id: int | None = None
    label_id: int | None = None


class MaterialUpdate(BaseModel):
    name: str | None = Field(default=None, max_length=100)
    unit: str | None = None
    min_weight: float | None = Field(default=None, ge=0)
    email: str | None = None
    excel_path: str | None = None
    action_type: str | None = None
    contact_id: int | None = None
    label_id: int | None = None


class MaterialRead(BaseModel):
    id: int
    name: str
    unit: str
    min_weight: float
    email: str | None
    excel_path: str | None
    action_type: str
    contact_id: int | None = None
    label_id: int | None
    label: LabelRead | None = None
    created_at: datetime
    # computed fields
    current_stock: float = 0.0
    fraction_stock: float = 0.0
    predicted_stock: float = 0.0
    is_low_stock: bool = False
    lot_count: int = 0
    contact_name: str | None = None
    contact_email: str | None = None

    model_config = {"from_attributes": True}


class MaterialListRead(BaseModel):
    """Lighter version for list views."""
    id: int
    name: str
    unit: str
    min_weight: float
    label: LabelRead | None = None
    current_stock: float = 0.0
    predicted_stock: float = 0.0
    is_low_stock: bool = False
    lot_count: int = 0

    model_config = {"from_attributes": True}
