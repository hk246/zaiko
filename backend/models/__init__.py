"""SQLAlchemy models – import all so Base.metadata sees them."""

from backend.models.label import MaterialLabel  # noqa: F401
from backend.models.lot import Lot  # noqa: F401
from backend.models.master import Contact, Tare  # noqa: F401
from backend.models.material import RawMaterial  # noqa: F401
from backend.models.property_field import LotProperty, PropertyField  # noqa: F401
from backend.models.recipe import Recipe, RecipeItem  # noqa: F401
from backend.models.reservation import Reservation, ReservationUsage  # noqa: F401
from backend.models.stock_transaction import StockTransaction  # noqa: F401
from backend.models.work_order import WorkOrder  # noqa: F401
