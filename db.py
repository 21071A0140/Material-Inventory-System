"""
db.py — Postgres-backed storage layer.

This module is a drop-in replacement for the JSON-file load_*/save_* functions
that used to live in main.py. Every function here has the EXACT same name,
signature, and return shape as its file-based counterpart, so main.py's
endpoints don't need to change — only the import + a few file-only helpers
(create/delete project, list projects) change.

Connection: reads DATABASE_URL from the environment. On Render, set this to
the Internal Database URL of your Postgres instance (Environment tab of your
web service). Locally, you can point it at any Postgres you have running.

Concurrency: every save_* is a single UPSERT statement (INSERT ... ON CONFLICT
DO UPDATE), so two users saving the same project at the same time never
corrupt each other's write — Postgres serializes it at the row level.
"""

import os
import json
import threading
from contextlib import contextmanager

import psycopg2
import psycopg2.extras
from psycopg2.pool import ThreadedConnectionPool

DATABASE_URL = os.environ.get("DATABASE_URL", "")

_pool = None
_pool_lock = threading.Lock()


def _get_pool():
    global _pool
    if _pool is None:
        with _pool_lock:
            if _pool is None:
                if not DATABASE_URL:
                    raise RuntimeError(
                        "DATABASE_URL environment variable is not set. "
                        "Set it to your Render Postgres Internal Database URL."
                    )
                # Render internal URLs don't need SSL; external ones do.
                # Detect by hostname: internal = *.render.com suffix absent
                ssl = "require" if ".oregon-postgres.render.com" in DATABASE_URL or \
                                   ".render.com" in DATABASE_URL.split("@")[-1] and \
                                   "oregon-postgres" in DATABASE_URL else "prefer"
                _pool = ThreadedConnectionPool(1, 10, DATABASE_URL, sslmode=ssl)
    return _pool


@contextmanager
def _conn():
    pool = _get_pool()
    conn = pool.getconn()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        pool.putconn(conn)


def init_db():
    """Run once at startup — creates all tables if they don't exist yet.
    Reads schema.sql from the same directory as this file."""
    schema_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "schema.sql")
    with open(schema_path, "r", encoding="utf-8") as f:
        schema_sql = f.read()
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(schema_sql)


# ── Project registry ──────────────────────────────────────────────────────

def list_projects():
    """Return list of project names — replaces the old
    [d.name for d in PROJECTS.iterdir() if d.is_dir()] pattern."""
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT name FROM projects ORDER BY name")
            return [r[0] for r in cur.fetchall()]


def project_exists(project: str) -> bool:
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM projects WHERE name = %s", (project,))
            return cur.fetchone() is not None


def create_project(project: str):
    """Idempotent — safe to call even if the project already exists."""
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO projects (name) VALUES (%s) ON CONFLICT (name) DO NOTHING",
                (project,),
            )


def delete_project(project: str):
    """Deletes the project row — CASCADE wipes every domain table's row
    for this project automatically (see schema.sql FK definitions)."""
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM projects WHERE name = %s", (project,))


def rename_project(old_name: str, new_name: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE projects SET name = %s WHERE name = %s",
                (new_name, old_name),
            )


def project_last_updated(project: str):
    """Returns the latest updated_at timestamp across all of this project's
    domain tables — used by the polling endpoint so the frontend can ask
    'has anything changed since I last checked?' with one cheap query."""
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT last_updated FROM project_last_updated WHERE project_name = %s",
                (project,),
            )
            row = cur.fetchone()
            return row[0].isoformat() if row and row[0] else None


def all_projects_last_updated():
    """Bulk version — one query for every project's last-updated time.
    Used by the dashboard poll so checking all open projects costs 1 round trip."""
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT project_name, last_updated FROM project_last_updated")
            return {row[0]: row[1].isoformat() if row[1] else None for row in cur.fetchall()}


# ── Generic JSONB domain table read/write ────────────────────────────────
# Every domain (meta, items, schedule_v2, etc.) follows the identical
# pattern: SELECT data FROM <table> WHERE project_name = %s, defaulting to
# `default` if no row exists yet (mirrors the old "file doesn't exist yet"
# behavior of the JSON-file functions).

def _load(table: str, project: str, default):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(f"SELECT data FROM {table} WHERE project_name = %s", (project,))
            row = cur.fetchone()
            if row is None:
                return default
            return row[0]


def _save(table: str, project: str, data):
    # Ensure the project row exists first (FK requirement)
    create_project(project)
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                f"""
                INSERT INTO {table} (project_name, data, updated_at)
                VALUES (%s, %s, now())
                ON CONFLICT (project_name)
                DO UPDATE SET data = EXCLUDED.data, updated_at = now()
                """,
                (project, psycopg2.extras.Json(data)),
            )


# ── load_*/save_* — EXACT same names + signatures as the old file-based
#    functions in main.py. Only the import line + these definitions change;
#    every endpoint that calls load_items(project) etc. works unmodified. ──

def load_meta(project):
    return _load("meta", project, {})

def save_meta(project, meta):
    _save("meta", project, meta)


def load_items(project):
    return _load("items", project, [])

def save_items(project, items):
    _save("items", project, items)


def load_schedule(project):
    return _load("schedule_legacy", project, {})

def save_schedule(project, data):
    _save("schedule_legacy", project, data)


def load_labor(project: str):
    return _load("labor", project, {})

def save_labor(project: str, data: dict):
    _save("labor", project, data)


def load_bt_estimate(project: str):
    return _load("bt_estimate", project, None)

def save_bt_estimate(project: str, data):
    _save("bt_estimate", project, data)


def load_bt_pos(project: str):
    return _load("bt_pos", project, [])

def save_bt_pos(project: str, data):
    _save("bt_pos", project, data)


def load_sched_v2(project: str):
    return _load("schedule_v2", project, {})

def save_sched_v2(project: str, data: dict):
    _save("schedule_v2", project, data)


def load_baselines(project: str):
    return _load("baselines", project, [])

def save_baselines(project: str, data: dict):
    _save("baselines", project, data)


def load_calendar(project: str):
    return _load("calendar", project, {})

def save_calendar(project: str, data: dict):
    _save("calendar", project, data)


def load_ma_results(project: str):
    return _load("ma_results", project, {})

def save_ma_results(project: str, data: dict):
    _save("ma_results", project, data)