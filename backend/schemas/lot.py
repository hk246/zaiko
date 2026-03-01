"""Pydantic schemas for Lot."""

from datetime import datetime

from pydantic import BaseModel, Field


class LotCreate(BaseModel):
    lot_name: str = Field(max_length=100)
    weight: float = Field(ge=0)
    is_fraction: bool = False
    tare_id: int | None = None  # optional: subtract tare weight


class LotUpdate(BaseModel):
    lot_name: str | None = Field(default=None, max_length=100)
    weight: float | None = Field(default=None, ge=0)
    is_fraction: bool | None = None
    tare_id: int | None = None  # optional: subtract tare weight from weight


class LotRead(BaseModel):
    id: int
    material_id: int
    lot_name: str
    weight: float
    is_fraction: bool
    created_at: datetime
    properties: list["LotPropertyRead"] = []

    model_config = {"from_attributes": True}


class LotPropertyRead(BaseModel):
    field_id: int
    field_name: str
    field_type: str
    unit: str | None
    value_string: str | None
    value_number: float | None

    model_config = {"from_attributes": True}


class LotPropertyWrite(BaseModel):
    field_id: int
    value_string: str | None = None
    value_number: float | None = None
