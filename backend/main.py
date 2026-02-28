"""FastAPI application entry point."""

import socket
import sys
import threading
import webbrowser
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from backend.config import settings
from backend.database import init_db
from backend.routers import (
    backup,
    dashboard,
    email_router,
    export_router,
    import_router,
    labels,
    lots,
    master,
    materials,
    properties,
    recipes,
    reservations,
    settings as settings_router,
    work_orders,
)


def _resolve_frontend_dir() -> Path:
    """Resolve frontend dist directory for both dev and PyInstaller."""
    if getattr(sys, "frozen", False):
        # PyInstaller bundle: frontend/dist is packed alongside the exe
        return Path(sys._MEIPASS) / "frontend" / "dist"
    return Path(__file__).resolve().parent.parent / "frontend" / "dist"


FRONTEND_DIR = _resolve_frontend_dir()


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Initialize DB on startup."""
    init_db()
    yield


app = FastAPI(
    title=settings.app_name,
    version=settings.app_version,
    lifespan=lifespan,
)

# CORS for development (Vite dev server)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── API routes ──
api_prefix = "/api"
app.include_router(materials.router, prefix=api_prefix)
app.include_router(lots.router, prefix=api_prefix)
app.include_router(reservations.router, prefix=api_prefix)
app.include_router(recipes.router, prefix=api_prefix)
app.include_router(work_orders.router, prefix=api_prefix)
app.include_router(labels.router, prefix=api_prefix)
app.include_router(properties.router, prefix=api_prefix)
app.include_router(master.router, prefix=api_prefix)
app.include_router(dashboard.router, prefix=api_prefix)
app.include_router(backup.router, prefix=api_prefix)
app.include_router(settings_router.router, prefix=api_prefix)
app.include_router(import_router.router, prefix=api_prefix)
app.include_router(email_router.router, prefix=api_prefix)
app.include_router(export_router.router, prefix=api_prefix)

# ── Serve frontend (production) ──
if FRONTEND_DIR.is_dir():
    app.mount("/assets", StaticFiles(directory=FRONTEND_DIR / "assets"), name="assets")

    @app.get("/{full_path:path}")
    async def serve_spa(full_path: str):
        """Serve index.html for all non-API routes (SPA client-side routing)."""
        from fastapi.responses import FileResponse
        index = FRONTEND_DIR / "index.html"
        if index.exists():
            return FileResponse(index)
        return {"detail": "Frontend not built. Run: cd frontend && npm run build"}


def _find_free_port(host: str, preferred: int) -> int:
    """Return *preferred* if available, otherwise find a free port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind((host, preferred))
            return preferred
        except OSError:
            pass
    # preferred port is busy – let the OS assign a free one
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind((host, 0))
        return s.getsockname()[1]


def _open_browser_delayed(url: str, delay: float = 1.5):
    """Open the default browser after a short delay so the server is ready."""
    def _open():
        import time
        time.sleep(delay)
        webbrowser.open(url)
    t = threading.Thread(target=_open, daemon=True)
    t.start()


def start():
    """Entry point for production (called from run.py or PyInstaller)."""
    import uvicorn

    host = settings.host
    port = _find_free_port(host, settings.port)

    url = f"http://{host}:{port}"
    print(f"\n  {settings.app_name}")
    print(f"  URL: {url}")
    print(f"  API docs: {url}/docs\n")

    _open_browser_delayed(url)
    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    start()
