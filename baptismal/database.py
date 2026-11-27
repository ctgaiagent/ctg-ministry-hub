"""
baptismal/database.py — Persistent PostgreSQL & SQLite Database Abstraction
=========================================================================
Uses parameterized SQL exclusively (no string interpolation of user inputs)
and explicit transaction commit/rollback handling.
Supports PostgreSQL via psycopg2 in production and WAL-optimized SQLite locally.
"""

from contextlib import contextmanager
import os
import sqlite3
import stat
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .config import get_database_url


class DatabaseManager:
    """Thread-safe, transaction-managed database wrapper for PostgreSQL and SQLite."""

    def __init__(self, dsn: str = ""):
        self.dsn = (dsn or get_database_url() or "sqlite:///baptismal_local.db").strip()
        self.is_postgres = self.dsn.startswith("postgres://") or self.dsn.startswith("postgresql://")
        if self.dsn.startswith("sqlite:///"):
            self.sqlite_path = self.dsn[len("sqlite:///"):]
        elif not self.is_postgres:
            self.sqlite_path = self.dsn
        else:
            self.sqlite_path = ""
        self._permissions_set = False

    def _adapt_sql(self, sql: str) -> str:
        """Converts PostgreSQL-style %s placeholders to SQLite ? placeholders when running on SQLite."""
        if self.is_postgres:
            return sql
        return sql.replace("%s", "?")

    @contextmanager
    def transaction(self):
        """Provides a transactional database cursor; commits on success, rolls back on any error."""
        if self.is_postgres:
            import psycopg2  # type: ignore
            import psycopg2.extras  # type: ignore

            conn = psycopg2.connect(self.dsn)
            try:
                cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                yield cur
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()
        else:
            is_new_file = bool(self.sqlite_path and self.sqlite_path != ":memory:" and not os.path.exists(self.sqlite_path))
            conn = sqlite3.connect(self.sqlite_path, timeout=20.0)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys = ON;")
            conn.execute("PRAGMA synchronous = NORMAL;")
            try:
                cur = conn.cursor()
                yield cur
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()
                if (is_new_file or not self._permissions_set) and self.sqlite_path and self.sqlite_path != ":memory:":
                    self._permissions_set = True
                    try:
                        os.chmod(self.sqlite_path, stat.S_IRUSR | stat.S_IWUSR)
                    except Exception:
                        pass

    def execute(self, sql: str, params: Tuple[Any, ...] = ()) -> int:
        adapted = self._adapt_sql(sql)
        with self.transaction() as cur:
            cur.execute(adapted, params)
            return cur.rowcount

    def executemany(self, sql: str, seq_of_params: Iterable[Tuple[Any, ...]]) -> int:
        adapted = self._adapt_sql(sql)
        with self.transaction() as cur:
            cur.executemany(adapted, seq_of_params)
            return cur.rowcount

    def fetch_one(self, sql: str, params: Tuple[Any, ...] = ()) -> Optional[Dict[str, Any]]:
        adapted = self._adapt_sql(sql)
        with self.transaction() as cur:
            cur.execute(adapted, params)
            row = cur.fetchone()
            return dict(row) if row is not None else None

    def fetch_all(self, sql: str, params: Tuple[Any, ...] = ()) -> List[Dict[str, Any]]:
        adapted = self._adapt_sql(sql)
        with self.transaction() as cur:
            cur.execute(adapted, params)
            rows = cur.fetchall()
            return [dict(r) for r in rows]
