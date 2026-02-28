"""Recipe execution – create batch reservations from a recipe."""

from datetime import date, datetime

from sqlalchemy.orm import Session

from backend.models.recipe import Recipe
from backend.models.reservation import Reservation


def execute_recipe(
    db: Session,
    recipe_id: int,
    scheduled_date: date,
    user_name: str | None = None,
    purpose: str | None = None,
) -> list[Reservation]:
    """Create one reservation per recipe item, all linked by recipe_id."""
    recipe = db.get(Recipe, recipe_id)
    if not recipe or recipe.deleted_at:
        raise ValueError(f"Recipe {recipe_id} not found")

    reservations: list[Reservation] = []
    for item in recipe.items:
        res = Reservation(
            material_id=item.material_id,
            lot_name=item.lot_name,
            recipe_id=recipe.id,
            type=recipe.type,
            quantity=item.quantity,
            user_name=user_name,
            purpose=purpose or f"レシピ: {recipe.name}",
            scheduled_date=scheduled_date,
            status="pending",
        )
        db.add(res)
        reservations.append(res)

    db.commit()
    for r in reservations:
        db.refresh(r)
    return reservations
