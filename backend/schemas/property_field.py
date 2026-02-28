"""Pydantic schemas for PropertyField."""

from datetime import datetime

from pydantic import BaseModel, Field


class PropertyFieldCreate(BaseModel):
    label_id: int | None = None
    name: str = Field(max_length=100)
    field_type: str = Field(default="string")
    unit: str | None = None
    order_index: int = 0


class PropertyFieldUpdate(BaseModel):
    name: str | None = Field(default=None, max_length=100)
    field_type: str | None = None
    unit: str | None = None
    order_index: int | None = None


class PropertyFieldRead(BaseModel):
    id: int
    label_id: int | None
    name: str
    field_type: str
    unit: str | None
    order_index: int
    created_at: datetime
    model_config = {"from_attributes": True}


class PropertyFieldReorder(BaseModel):
    field_ids: list[int]
