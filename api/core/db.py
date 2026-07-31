"""
SQLite Database Connection and Initialization for Cashier Intelligence.
"""

from __future__ import annotations

import sqlite3
import logging
import threading
from contextlib import contextmanager
from typing import Iterator

from .config import DATA_DIR

log = logging.getLogger(__name__)

DB_PATH = DATA_DIR / "cashiers.db"
_lock = threading.Lock()


@contextmanager
def _connect() -> Iterator[sqlite3.Connection]:
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    """Инициализирует таблицы БД для кассиров."""
    from .cashier_analytics import init_cashier_tables
    init_cashier_tables()
    log.info("БД инициализирована: %s", DB_PATH)
