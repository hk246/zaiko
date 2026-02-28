"""Pydantic schemas for Tare & Contact."""

from datetime import datetime

from pydantic import BaseModel, Field


# ── Tare ──
class TareCreate(BaseModel):
    name: str = Field(max_length=100)
    weight: float
    description: str | None = None


class TareUpdate(BaseModel):
    name: str | None = Field(default=None, max_length=100)
    weight: float | None = None
    description: str | None = None


class TareRead(BaseModel):
    id: int
    name: str
    weight: float
    description: str | None
    created_at: datetime
    model_config = {"from_attributes": True}


# ── Contact ──
class ContactCreate(BaseModel):
    name: str = Field(max_length=100)
    email: str = Field(max_length=120)
    description: str | None = None


class ContactUpdate(BaseModel):
    name: str | None = Field(default=None, max_length=100)
    email: str | None = Field(default=None, max_length=120)
    description: str | None = None


class ContactRead(BaseModel):
    id: int
    name: str
    email: str
    description: str | None
    created_at: datetime
    model_config = {"from_attributes": True}
