"""Email draft endpoint – opens Outlook with pre-filled email for low-stock alerts."""

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.material import RawMaterial
from backend.models.master import Contact
from backend.services import stock_calculator

router = APIRouter(prefix="/email", tags=["email"])


class EmailDraftRequest(BaseModel):
    material_id: int
    contact_id: int


class EmailDraftResponse(BaseModel):
    status: str
    message: str


@router.post("/draft", response_model=EmailDraftResponse)
def create_email_draft(body: EmailDraftRequest, db: Session = Depends(get_db)):
    """材料の低在庫アラートメールの下書きをOutlookで開く。"""
    material = db.get(RawMaterial, body.material_id)
    if not material or material.deleted_at:
        raise HTTPException(404, "材料が見つかりません")

    contact = db.get(Contact, body.contact_id)
    if not contact:
        raise HTTPException(404, "連絡先が見つかりません")

    current = stock_calculator.current_stock(db, material.id)
    periods = stock_calculator.critical_periods(db, material.id, material.min_weight)

    # Build email
    subject = f"【在庫アラート】{material.name} - 在庫不足の見込み"

    lines = [
        f"{contact.name} 様",
        "",
        "以下の材料について、在庫不足の見込みをお知らせいたします。",
        "",
        f"■ 材料名: {material.name}",
        f"■ 現在の在庫: {current:.1f} {material.unit}",
        f"■ 最低在庫量: {material.min_weight:.1f} {material.unit}",
    ]

    if periods:
        p = periods[0]
        lines.append(f"■ 不足予測日: {p.start_date}")
        lines.append(f"■ 不足量: {abs(p.shortage_amount):.1f} {material.unit}")

    lines.extend([
        "",
        "ご確認の上、補充のご手配をお願いいたします。",
        "",
        "よろしくお願いいたします。",
    ])

    email_body = "\n".join(lines)

    from backend.services.email_service import create_outlook_draft
    try:
        result = create_outlook_draft(
            to_email=contact.email,
            subject=subject,
            body=email_body,
        )
        return result
    except RuntimeError as e:
        raise HTTPException(500, str(e))
