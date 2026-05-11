# website/backend/backup_restore.py
import json, sqlite3, os, sys
from pathlib import Path

# ------------------------------------------------------------
# Detect PyInstaller environment & resolve correct base folder
# ------------------------------------------------------------
def _base_dir():
    if getattr(sys, "frozen", False):
        # Installed EXE: always store DB in LocalAppData\BrewInsPOS
        return Path(os.environ["LOCALAPPDATA"]) / "BrewInsPOS"
    else:
        # Running from source
        return Path(__file__).resolve().parents[1]

BASE_DIR = _base_dir()

DB_FILE     = BASE_DIR / "backend" / "pos.db"
MODELS_SQL  = BASE_DIR / "backend" / "models.sql"


CATALOG_TABLES = ["items", "components", "inventory", "recipes"]  # order matters on restore

def _connect(db_path=DB_FILE):
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON;")
    return con

def backup_catalog(out_dir: str):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    with _connect() as c:
        for t in CATALOG_TABLES:
            rows = [dict(r) for r in c.execute(f"SELECT * FROM {t}")]
            (out / f"{t}.json").write_text(json.dumps(rows, indent=2), encoding="utf-8")
    print(f"Catalog backup complete → {out.resolve()}")

def restore_catalog(backup_dir: str, new_db: str | None = None):
    """Rebuild schema from models.sql, then load JSON data into tables."""
    backup = Path(backup_dir)
    if not backup.exists():
        raise FileNotFoundError(f"Backup dir not found: {backup}")

    target_db = Path(new_db) if new_db else DB_FILE
    # 1) Recreate DB from schema
    if target_db.exists():
        target_db.unlink()
    schema_sql = MODELS_SQL.read_text(encoding="utf-8")
    with _connect(target_db) as c:
        c.executescript(schema_sql)
        c.commit()

    # 2) Insert data in dependency-safe order
    with _connect(target_db) as c:
        c.execute("PRAGMA foreign_keys = OFF;")
        c.commit()

        def load_json(name):
            p = backup / f"{name}.json"
            return json.loads(p.read_text(encoding="utf-8")) if p.exists() else []

        # items
        items = load_json("items")
        if items:
            cols = ["id","sku","name","category","price","active"]
            c.executemany(
                f"INSERT INTO items ({','.join(cols)}) VALUES (?,?,?,?,?,?)",
                [(r.get("id"), r.get("sku"), r.get("name"), r.get("category"),
                  float(r.get("price",0)), int(r.get("active",1))) for r in items]
            )

        # components
        comps = load_json("components")
        if comps:
            # handle either old/new component shapes
            # new shape: id,name,qty_used,unit_cost,pos_track_sellout
            # legacy shape: may include unit, qty_on_hand,.. ignored here
            cols_present = set(comps[0].keys())
            if {"qty_used","unit_cost"} <= cols_present:
                c.executemany(
                    "INSERT INTO components (id,name,qty_used,unit_cost,pos_track_sellout) VALUES (?,?,?,?,?)",
                    [(r.get("id"), r.get("name"), float(r.get("qty_used",0)),
                      float(r.get("unit_cost",0)), int(r.get("pos_track_sellout",0))) for r in comps]
                )
            else:
                # minimal: id,name,pos_track_sellout
                c.executemany(
                    "INSERT INTO components (id,name,pos_track_sellout) VALUES (?,?,?)",
                    [(r.get("id"), r.get("name"), int(r.get("pos_track_sellout",0))) for r in comps]
                )

        # inventory
        inv = load_json("inventory")
        if inv:
            cols = ["id","name","unit","case_cost","units_per_case","qty_on_hand","reorder_point"]
            c.executemany(
                f"INSERT INTO inventory ({','.join(cols)}) VALUES (?,?,?,?,?,?,?)",
                [(r.get("id"), r.get("name"), r.get("unit"),
                  float(r.get("case_cost",0)), float(r.get("units_per_case",0)),
                  float(r.get("qty_on_hand",0)), float(r.get("reorder_point",0))) for r in inv]
            )

        # recipes (depends on items + components)
        recs = load_json("recipes")
        if recs:
            cols = ["id","item_id","component_id","qty_per_item"]
            c.executemany(
                f"INSERT INTO recipes ({','.join(cols)}) VALUES (?,?,?,?)",
                [(r.get("id"), r.get("item_id"), r.get("component_id"),
                  float(r.get("qty_per_item",0))) for r in recs]
            )

        # Clean slate for operational tables (leave empty)
        for t in ("orders","order_lines","shifts","expenses"):
            c.execute(f"DELETE FROM {t}")

        c.execute("PRAGMA foreign_keys = ON;")
        c.commit()
    print(f"Catalog restore complete → {target_db.resolve()}")

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Backup/Restore POS catalog tables.")
    ap.add_argument("action", choices=["backup","restore"])
    ap.add_argument("--dir", required=True, help="Directory to write/read JSON files")
    ap.add_argument("--db", help="Optional path for new DB when restoring")
    args = ap.parse_args()

    if args.action == "backup":
        backup_catalog(args.dir)
    else:
        restore_catalog(args.dir, new_db=args.db)
