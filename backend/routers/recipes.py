"""Recipe CRUD + execution endpoints."""

from datetime import date

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from backend.database import get_db
from backend.models.recipe import Recipe, RecipeItem
from backend.schemas.recipe import RecipeCreate, RecipeExecute, RecipeItemRead, RecipeRead, RecipeUpdate
from backend.services.recipe_executor import execute_recipe

router = APIRouter(prefix="/recipes", tags=["recipes"])


def _to_read(r: Recipe) -> dict:
    items = []
    for item in r.items:
        items.append(RecipeItemRead(
            id=item.id,
            material_id=item.material_id,
            material_name=item.material.name if item.material else "",
            quantity=item.quantity,
            lot_name=item.lot_name,
        ))
    return {
        "id": r.id,
        "name": r.name,
        "description": r.description,
        "type": r.type,
        "items": items,
        "created_at": r.created_at,
    }


@router.get("", response_model=list[RecipeRead])
def list_recipes(db: Session = Depends(get_db)):
    recipes = (
        db.query(Recipe)
        .filter(Recipe.deleted_at == None)  # noqa: E711
        .order_by(Recipe.name)
        .all()
    )
    return [_to_read(r) for r in recipes]


@router.get("/{recipe_id}", response_model=RecipeRead)
def get_recipe(recipe_id: int, db: Session = Depends(get_db)):
    r = db.get(Recipe, recipe_id)
    if not r or r.deleted_at:
        raise HTTPException(404, "Recipe not found")
    return _to_read(r)


@router.post("", response_model=RecipeRead, status_code=201)
def create_recipe(body: RecipeCreate, db: Session = Depends(get_db)):
    recipe = Recipe(name=body.name, description=body.description, type=body.type)
    db.add(recipe)
    db.flush()
    for item in body.items:
        db.add(RecipeItem(
            recipe_id=recipe.id,
            material_id=item.material_id,
            quantity=item.quantity,
            lot_name=item.lot_name,
        ))
    db.commit()
    db.refresh(recipe)
    return _to_read(recipe)


@router.put("/{recipe_id}", response_model=RecipeRead)
def update_recipe(recipe_id: int, body: RecipeUpdate, db: Session = Depends(get_db)):
    r = db.get(Recipe, recipe_id)
    if not r or r.deleted_at:
        raise HTTPException(404, "Recipe not found")

    if body.name is not None:
        r.name = body.name
    if body.description is not None:
        r.description = body.description
    if body.type is not None:
        r.type = body.type

    if body.items is not None:
        # Replace all items
        for old in r.items:
            db.delete(old)
        db.flush()
        for item in body.items:
            db.add(RecipeItem(
                recipe_id=r.id,
                material_id=item.material_id,
                quantity=item.quantity,
                lot_name=item.lot_name,
            ))

    db.commit()
    db.refresh(r)
    return _to_read(r)


@router.delete("/{recipe_id}", status_code=204)
def delete_recipe(recipe_id: int, db: Session = Depends(get_db)):
    from datetime import datetime
    r = db.get(Recipe, recipe_id)
    if not r or r.deleted_at:
        raise HTTPException(404, "Recipe not found")
    r.deleted_at = datetime.utcnow()
    db.commit()


@router.post("/{recipe_id}/execute")
def execute(recipe_id: int, body: RecipeExecute, db: Session = Depends(get_db)):
    try:
        scheduled = date.fromisoformat(body.scheduled_date)
        reservations = execute_recipe(db, recipe_id, scheduled, body.user_name, body.purpose)
        return {"created_reservations": len(reservations)}
    except ValueError as e:
        raise HTTPException(400, str(e))
