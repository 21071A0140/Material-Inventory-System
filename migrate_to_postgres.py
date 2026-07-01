"""
migrate_to_postgres.py — one-time migration of existing JSON project data
into Postgres.

Usage:
    1. Set DATABASE_URL env var to your Render Postgres connection string.
    2. Place your existing `projects/` folder next to this script
       (the same folder structure as on your laptop / old Render disk).
    3. Run:  python migrate_to_postgres.py

Safe to re-run: every project + domain is upserted, so running this twice
just overwrites with the same data (no duplicates).
"""

import os
import sys
import json
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import db

PROJECTS_DIR = Path(os.environ.get("PROJECTS_DIR", "projects"))

# Maps each JSON filename to the db.py save_* function that should store it.
FILE_TO_SAVER = {
    "meta.json":         db.save_meta,
    "items.json":        db.save_items,
    "schedule.json":     db.save_schedule,
    "schedule_v2.json":  db.save_sched_v2,
    "baselines.json":    db.save_baselines,
    "calendar.json":     db.save_calendar,
    "labor.json":        db.save_labor,
    "bt_estimate.json":  db.save_bt_estimate,
    "bt_pos.json":       db.save_bt_pos,
    "ma_results.json":   db.save_ma_results,
}


def migrate():
    if not PROJECTS_DIR.exists():
        print(f"ERROR: {PROJECTS_DIR} not found. Set PROJECTS_DIR env var "
              f"or place a 'projects' folder next to this script.")
        sys.exit(1)

    db.init_db()
    print("✓ Schema ensured (tables created if not already present).\n")

    project_dirs = [d for d in PROJECTS_DIR.iterdir() if d.is_dir()]
    if not project_dirs:
        print("No project folders found — nothing to migrate.")
        return

    total_files = 0
    for pdir in sorted(project_dirs):
        project_name = pdir.name
        print(f"── {project_name} ──")
        db.create_project(project_name)

        found_any = False
        for fname, saver in FILE_TO_SAVER.items():
            fpath = pdir / fname
            if not fpath.exists():
                continue
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                print(f"   ✗ {fname}: could not parse JSON ({e}) — skipped")
                continue
            saver(project_name, data)
            print(f"   ✓ {fname} → migrated")
            found_any = True
            total_files += 1

        if not found_any:
            print(f"   (no recognized JSON files found in this folder)")
        print()

    print(f"Done. {len(project_dirs)} project(s), {total_files} data file(s) migrated.")
    print("\nVerifying — projects now in Postgres:")
    for name in db.list_projects():
        print(f"   • {name}")


if __name__ == "__main__":
    migrate()