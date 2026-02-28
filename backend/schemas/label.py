"""Pydantic schemas for MaterialLabel."""

from datetime import datetime

from pydantic import BaseModel, Field


class LabelCreate(BaseModel):
    name: str = Field(max_length=50)
    color: str = Field(default="#6c757d", max_length=7)


class LabelUpdate(BaseModel):
    name: str | None = Field(default=None, max_length=50)
    color: str | None = Field(default=None, max_length=7)


class LabelRead(BaseModel):
    id: int
    name: str
    color: str
    created_at: datetime
    material_count: int = 0

    model_config = {"from_attributes": True}
