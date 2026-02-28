"""Database engine, session, and base model.

ネットワークドライブ上のSQLiteを複数ユーザーで共有する構成に対応。
- WALモードはネットワークドライブ非対応のため DELETE モードを使用
- busy_timeout で他ユーザーのロック待ちに対応
- NullPool で接続を溜め込まず都度接続/切断
"""

from collections.abc import Generator

from sqlalchemy import create_engine, event
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker
from sqlalchemy.pool import NullPool

from backend.config import settings

engine = create_engine(
    settings.database_url,
    connect_args={"check_same_thread": False},
    poolclass=NullPool,
    echo=False,
)


@event.listens_for(engine, "connect")
def _set_sqlite_pragma(dbapi_conn, _connection_record):
    """SQLiteプラグマ設定（ネットワークドライブ対応）"""
    cursor = dbapi_conn.cursor()
    # DELETEジャーナルモード: ネットワークドライブで安全に動作
    # (WALは共有メモリを使うためネットワーク非対応)
    cursor.execute("PRAGMA journal_mode=DELETE")
    # 他ユーザーがロック中の場合、最大10秒待機してリトライ
    cursor.execute("PRAGMA busy_timeout=10000")
    # 外部キー制約を有効化
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()


SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)


class Base(DeclarativeBase):
    pass


def get_db() -> Generator[Session, None, None]:
    """FastAPI dependency that yields a DB session."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db() -> None:
    """Create all tables (for development / first run)."""
    import backend.models  # noqa: F401 – ensure models are registered

    Base.metadata.create_all(bind=engine)
    _run_migrations()


def _run_migrations() -> None:
    """既存DBに新カラムを追加（バージョンアップ対応）"""
    import sqlite3
    from pathlib import Path
    from backend.config import settings

    db_path = str(Path(settings.database_folder) / "inventory.db")
    if not Path(db_path).exists():
        return
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute("PRAGMA table_info(raw_materials)")
        columns = [col[1] for col in cursor.fetchall()]
        if "contact_id" not in columns:
            cursor.execute(
                "ALTER TABLE raw_materials ADD COLUMN contact_id INTEGER REFERENCES contacts(id)"
            )
            conn.commit()
    finally:
        conn.close()
