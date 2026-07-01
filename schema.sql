-- ════════════════════════════════════════════════════════════════════════
-- MATERIAL INVENTORY SYSTEM — POSTGRES SCHEMA
-- ════════════════════════════════════════════════════════════════════════
-- Design: each domain that used to be a JSON file (projects/<name>/X.json)
-- becomes one row in a dedicated table, keyed by project name, storing the
-- exact same JSON shape in a JSONB column. This means:
--   - No data-shape redesign needed — existing main.py code that does
--     json.loads()/json.dumps() on these structures keeps working untouched
--     once load_*/save_* read from Postgres instead of disk.
--   - Every save is a single UPSERT (INSERT ... ON CONFLICT DO UPDATE),
--     so concurrent users never corrupt each other's data the way two
--     people writing the same JSON file at once could.
--   - updated_at lets the frontend polling layer ask "did anything change
--     since I last checked?" with one cheap indexed query per project.
-- ════════════════════════════════════════════════════════════════════════

-- Master project registry — every project must have a row here.
-- Deleting a project deletes this row, which CASCADEs to every domain table.
CREATE TABLE IF NOT EXISTS projects (
    name          TEXT PRIMARY KEY,
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- One table per JSON-file domain. Pattern is identical for all nine:
--   project_name TEXT  → FK to projects.name, ON DELETE CASCADE
--   data         JSONB → the exact JSON blob that used to be the file content
--   updated_at   TIMESTAMPTZ → bumped on every save, used by the poll endpoint

CREATE TABLE IF NOT EXISTS meta (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '{}'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS items (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '[]'::jsonb,   -- list of item dicts
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS schedule_legacy (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '{}'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS schedule_v2 (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '{}'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS baselines (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '[]'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS calendar (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '{}'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS labor (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '{}'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS bt_estimate (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '{}'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS bt_pos (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '[]'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS ma_results (
    project_name  TEXT PRIMARY KEY REFERENCES projects(name) ON DELETE CASCADE,
    data          JSONB NOT NULL DEFAULT '{}'::jsonb,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Index used by the "has anything changed?" poll the frontend calls every
-- ~15s. One cheap lookup per project tells us the single latest update
-- across ALL of that project's domains, so the client knows whether to
-- pull fresh data — without us needing to diff or compare payloads.
CREATE OR REPLACE VIEW project_last_updated AS
    SELECT project_name, MAX(updated_at) AS last_updated FROM (
        SELECT project_name, updated_at FROM meta
        UNION ALL SELECT project_name, updated_at FROM items
        UNION ALL SELECT project_name, updated_at FROM schedule_legacy
        UNION ALL SELECT project_name, updated_at FROM schedule_v2
        UNION ALL SELECT project_name, updated_at FROM baselines
        UNION ALL SELECT project_name, updated_at FROM calendar
        UNION ALL SELECT project_name, updated_at FROM labor
        UNION ALL SELECT project_name, updated_at FROM bt_estimate
        UNION ALL SELECT project_name, updated_at FROM bt_pos
        UNION ALL SELECT project_name, updated_at FROM ma_results
    ) AS all_updates
    GROUP BY project_name;