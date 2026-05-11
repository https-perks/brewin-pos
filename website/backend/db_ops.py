"""
Coffee Shop POS — Database Operations
-------------------------------------
Clean rewrite for the Flask web version.

Responsibilities:
• Handle all database reads/writes
• Track inventory (components)
• Record orders and automatically decrement inventory
• Provide summaries for reports
"""

import sqlite3
import sys
from pathlib import Path
from datetime import datetime
import os
import io, csv, zipfile, time, math
import shutil
from openpyxl import Workbook

# ---------------------------------------------------------------------
# SAFE LOGGING (NO EMOJI / NO print())
# ---------------------------------------------------------------------

def _log_path():
    """Return path to LocalAppData\BrewInsPOS\pos.log"""
    local = os.environ.get("LOCALAPPDATA", "")
    log_dir = Path(local) / "BrewInsPOS"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / "pos.log"

LOG_FILE = _log_path()

def log(msg: str):
    """Append UTF-8 log messages safely (never crashes)."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except:
        pass


# ---------------------------------------------------------------------
# Connection (PyInstaller-safe paths)
# ---------------------------------------------------------------------

# Determine base directory depending on whether we're frozen
if getattr(sys, 'frozen', False):
    # Running inside EXE — store DB in writable LocalAppData
    BASE_DIR = Path(os.environ["LOCALAPPDATA"]) / "BrewInsPOS"
else:
    # Source mode — project root, because this file is backend/db_ops.py
    BASE_DIR = Path(__file__).resolve().parent.parent

DB_FILE = BASE_DIR / "backend" / "pos.db"

# Alias for backup/restore helpers that expect DB_PATH
DB_PATH = str(DB_FILE)


def connect():
    DB_FILE.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def _debug_db_paths():
    log(f"---- DB PATH DEBUG ----")
    log(f"sys.frozen = {getattr(sys, 'frozen', False)}")
    log(f"sys._MEIPASS = {getattr(sys, '_MEIPASS', None)}")
    log(f"BASE_DIR = {BASE_DIR}")
    log(f"DB_FILE (exists?) = {DB_FILE}  /  {DB_FILE.exists()}")
    log(f"DB_PATH = {DB_PATH}")
    log(f"------------------------")


_debug_db_paths()

# ---------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------
def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")

def initialize_database():
    """
    Create database schema if tables do not exist, using models.sql.
    This only runs if pos.db exists but required tables are missing.
    """
    try:
        # Required tables
        required = {"items", "components", "recipes", "inventory", "orders", "order_lines"}
        
        with connect() as c:
            existing = {row[0] for row in c.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()}

            missing = required - existing
            if not missing:
                log(f"DB Init: All required tables already present.")
                return

            log(f"DB Init: Missing tables detected: {missing}")

            # Load models.sql from backend
            model_path = BASE_DIR / "backend" / "models.sql"
            if not model_path.exists():
                log(f"DB Init ERROR: models.sql not found at {model_path}")
                return

            sql_text = model_path.read_text(encoding="utf-8")
            c.executescript(sql_text)
            c.commit()

            log(f"DB Init: models.sql executed successfully — schema created.")

    except Exception as e:
        log(f"DB Init FAILED: {e}")

def make_safety_db_backup() -> str:
    """
    Copy the current database before a restore.
    Returns the path to the safety backup.
    """
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = _safety_backups_dir()

    src = DB_PATH
    filename = f"pos_safety_backup_{stamp}.db"
    dest = os.path.join(backup_dir, filename)

    shutil.copy2(src, dest)

    return dest

# ---------------------------------------------------------------------
# Migrations (run on app start)
# ---------------------------------------------------------------------
def _migrate_receipt_seq():
    """Ensure orders table has receipt_seq and receipt_date."""
    with connect() as c:
        cur = c.execute("PRAGMA table_info(orders)")
        cols = {row[1] for row in cur.fetchall()}
        if "receipt_seq" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN receipt_seq INTEGER")
        if "receipt_date" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN receipt_date TEXT")
        c.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS ux_orders_receipt
            ON orders(receipt_date, receipt_seq)
        """)
        c.commit()
    log(f"Migration: receipt sequence verified.")
        
def _migrate_orders_table():
    """Ensure the orders table exists."""
    with connect() as c:
        c.execute("""
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT NOT NULL,
            cashier TEXT NOT NULL,
            method TEXT NOT NULL,
            discount REAL NOT NULL DEFAULT 0,
            tax REAL NOT NULL DEFAULT 0,
            subtotal REAL NOT NULL DEFAULT 0,
            total REAL NOT NULL DEFAULT 0,
            cash_given REAL DEFAULT 0,
            change_given REAL DEFAULT 0,
            receipt_date TEXT,
            receipt_seq INTEGER,
            shift_id INTEGER,  -- REQUIRED so FK works
            FOREIGN KEY (shift_id) REFERENCES shifts(id)
            );
        """)
        c.commit()
    log(f"Migration: orders table verified.")
    
def _migrate_shifts_table():
    """Ensure the shifts table exists."""
    with connect() as c:
        c.execute("""
        CREATE TABLE IF NOT EXISTS shifts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts_start TEXT NOT NULL,
            ts_end TEXT,
            cashier TEXT NOT NULL,
            opening_float REAL NOT NULL DEFAULT 0,
            closing_amount REAL,
            over_short REAL,
            is_active INTEGER NOT NULL DEFAULT 1
        );
        """)
        c.commit()
    log(f"Migration: shifts table verified.")
        
def _migrate_order_cash_columns():
    """Ensure cash tracking columns exist."""
    with connect() as c:
        # shifts
        cols = [r["name"] for r in c.execute("PRAGMA table_info(shifts)").fetchall()]
        if "cash_sales" not in cols:
            c.execute("ALTER TABLE shifts ADD COLUMN cash_sales REAL DEFAULT 0")
        if "cash_given" not in cols:
            c.execute("ALTER TABLE shifts ADD COLUMN cash_given REAL DEFAULT 0")
        if "change_given" not in cols:
            c.execute("ALTER TABLE shifts ADD COLUMN change_given REAL DEFAULT 0")

        # orders
        cols = [r["name"] for r in c.execute("PRAGMA table_info(orders)").fetchall()]
        if "cash_given" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN cash_given REAL DEFAULT 0")
        if "change_given" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN change_given REAL DEFAULT 0")

        c.commit()
    log(f"Migration: cash tracking columns verified.")
        
def _migrate_shift_last_sale():
    """Ensure last_sale_at exists."""
    with connect() as c:
        cols = [r["name"] for r in c.execute("PRAGMA table_info(shifts)").fetchall()]
        if "last_sale_at" not in cols:
            c.execute("ALTER TABLE shifts ADD COLUMN last_sale_at TEXT")
        c.commit()
    log(f"Migration: last_sale_at verified.")
    
def _migrate_discount_type():
    with connect() as c:
        cols = [r["name"] for r in c.execute("PRAGMA table_info(orders)").fetchall()]
        if "discount_type" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN discount_type TEXT")
            c.commit()
            log(f"Migration: discount_type column added.")
            
def _migrate_settings_tables():
    with connect() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS admin_pins (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pin TEXT NOT NULL,
                label TEXT,
                active INTEGER NOT NULL DEFAULT 1
            )
        """)

        for key in [
            "require_pin_discounts",
            "require_pin_admin_access",
            "require_pin_exit_program",
            "tax_rate"
        ]:
            c.execute("""
                INSERT OR IGNORE INTO settings(key, value)
                VALUES (?, '0')
            """, (key,))
            
        c.commit()

def _migrate_sales_table():
    """Create sales/promotions table and ensure BOGO fields exist."""
    with connect() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                sale_type TEXT NOT NULL,
                amount REAL NOT NULL DEFAULT 0,
                item_id INTEGER,
                category TEXT,
                buy_qty INTEGER NOT NULL DEFAULT 1,
                free_qty INTEGER NOT NULL DEFAULT 1,
                auto_apply INTEGER NOT NULL DEFAULT 0,
                active INTEGER NOT NULL DEFAULT 1,
                starts_on TEXT,
                ends_on TEXT,
                FOREIGN KEY (item_id) REFERENCES items(id)
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS sale_requirements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_id INTEGER NOT NULL,
                requirement_type TEXT NOT NULL DEFAULT 'CATEGORY',
                item_id INTEGER,
                category TEXT,
                qty INTEGER NOT NULL DEFAULT 1,
                FOREIGN KEY (sale_id) REFERENCES sales(id) ON DELETE CASCADE,
                FOREIGN KEY (item_id) REFERENCES items(id)
            )
        """)

        cols = [r["name"] for r in c.execute("PRAGMA table_info(sales)").fetchall()]

        if "category" not in cols:
            c.execute("ALTER TABLE sales ADD COLUMN category TEXT")

        if "buy_qty" not in cols:
            c.execute("ALTER TABLE sales ADD COLUMN buy_qty INTEGER NOT NULL DEFAULT 1")

        if "free_qty" not in cols:
            c.execute("ALTER TABLE sales ADD COLUMN free_qty INTEGER NOT NULL DEFAULT 1")
        
        if "auto_apply" not in cols:
            c.execute("ALTER TABLE sales ADD COLUMN auto_apply INTEGER NOT NULL DEFAULT 0")

        c.commit()

    log("Migration: sales table verified.")

def _migrate_order_discounts_table():
    """Track detailed discount records per order."""
    with connect() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS order_discounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                sale_id INTEGER,
                sale_name TEXT,
                sale_type TEXT,
                source TEXT NOT NULL, -- AUTO or MANUAL
                amount REAL NOT NULL DEFAULT 0,
                authorized_pin_id INTEGER,
                authorized_pin_label TEXT,
                created_at TEXT NOT NULL DEFAULT (datetime('now', '-6 hours')),
                FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE,
                FOREIGN KEY (sale_id) REFERENCES sales(id),
                FOREIGN KEY (authorized_pin_id) REFERENCES admin_pins(id)
            )
        """)
        c.commit()

    log("Migration: order_discounts table verified.")
    
# ---------------------------------------------------------------------`
# Shifts (daily cashier sessions)
# ---------------------------------------------------------------------`
def start_shift(cashier:str, opening_float:float):
    with connect() as c:
        active = c.execute("SELECT id FROM shifts WHERE is_active=1").fetchone()
        if active:
            raise ValueError("Shift already active.")
        c.execute("""
            INSERT INTO shifts(ts_start, cashier, opening_float, is_active)
            VALUES(datetime('now', '-6 hours'), ?, ?, 1)
        """, (cashier, opening_float))
        c.commit()
    log(f"Shift started for {cashier}")
    return True


def get_active_shift():
    with connect() as c:
        row = c.execute("SELECT * FROM shifts WHERE is_active=1").fetchone()
        return dict(row) if row else None


def get_cash_summary(ts_start: str | None = None) -> dict:
    with connect() as c:
        if ts_start is None:
            shift = c.execute("SELECT * FROM shifts WHERE is_active=1").fetchone()
            if not shift:
                return {"opening": 0, "cash_sales": 0, "net_cash": 0, "expected": 0}
            opening = float(shift["opening_float"] or 0)
            ts_start = shift["ts_start"]
        else:
            sh = c.execute("SELECT opening_float FROM shifts WHERE ts_start=? LIMIT 1", (ts_start,)).fetchone()
            opening = float(sh["opening_float"] or 0) if sh else 0

        row = c.execute("""
            SELECT
                COALESCE(SUM(total), 0) AS cash_sales,
                COALESCE(SUM(cash_given - change_given), 0) AS net_cash
            FROM orders
            WHERE method='Cash' AND ts >= ?
        """, (ts_start,)).fetchone()

        cash_sales = float(row["cash_sales"] or 0)
        net_cash   = float(row["net_cash"] or 0)
        expected   = round(opening + net_cash, 2)

        return {
            "opening": round(opening, 2),
            "cash_sales": round(cash_sales, 2),
            "net_cash": round(net_cash, 2),
            "expected": expected
        }

# ---------------------------------------------------------------------
# Close shift
# ---------------------------------------------------------------------
def close_shift(actual_cash: float):
    """
    Close the current active shift.
    Calculates expected cash, computes over/short, and finalizes the record.
    """
    with connect() as c:
        shift = c.execute("SELECT * FROM shifts WHERE is_active=1").fetchone()
        if not shift:
            raise ValueError("No active shift found.")

        shift_id = shift["id"]
        ts_start = shift["ts_start"]
        opening = float(shift["opening_float"] or 0)

        # --- Calculate net cash movement since shift start ---
        row = c.execute("""
            SELECT 
                COALESCE(SUM(cash_given - change_given), 0) AS net_cash
            FROM orders
            WHERE method='Cash' AND ts >= ?
        """, (ts_start,)).fetchone()

        net_cash = float(row["net_cash"] or 0)
        expected = round(opening + net_cash, 2)
        diff = round(float(actual_cash) - expected, 2)

        # --- Update shift record ---
        c.execute("""
            UPDATE shifts
            SET ts_end = datetime('now', '-6 hours'),
                closing_amount = ?,
                over_short = ?,
                is_active = 0
            WHERE id = ?
        """, (actual_cash, diff, shift_id))
        c.commit()

        return {
            "cashier": shift["cashier"],
            "expected": expected,
            "actual": round(actual_cash, 2),
            "over_short": diff
        }

# ---------------------------------------------------------------------
# Close stale shift and log expected till
# ---------------------------------------------------------------------
def _auto_close_stale_shifts():
    """Auto-close any left-open shifts from previous days."""
    with connect() as c:
        rows = c.execute("""
            SELECT id, cashier, ts_start, opening_float
            FROM shifts
            WHERE is_active = 1
        """).fetchall()

        if not rows:
            log("No stale shifts found.")
            return

        closed = 0

        # Use same local-time offset as the rest of the POS
        today = c.execute("SELECT date('now', '-6 hours')").fetchone()[0]

        for r in rows:
            # Works for timestamps like '2025-10-11 14:22:00'
            start_date = (r["ts_start"] or "")[:10]

            if start_date != today:
                net_cash = c.execute("""
                    SELECT COALESCE(SUM(cash_given - change_given), 0)
                    FROM orders
                    WHERE method='Cash' AND ts >= ?
                """, (r["ts_start"],)).fetchone()[0]

                expected = float(r["opening_float"] or 0) + float(net_cash or 0)

                log(
                    f"Auto-close stale shift: "
                    f"cashier={r['cashier']} "
                    f"start={r['ts_start']} "
                    f"expected_till={expected}"
                )

                c.execute("""
                    UPDATE shifts
                    SET is_active = 0,
                        ts_end = datetime('now', '-6 hours'),
                        closing_amount = ?,
                        over_short = NULL
                    WHERE id = ?
                """, (expected, r["id"]))

                closed += 1

        c.commit()

        if closed:
            log(f"Auto-closed {closed} stale shift(s).")
        else:
            log("All active shifts are from today; no auto-close needed.")
            
# ---------------------------------------------------------------------
# Items (what appears on the POS)
# ---------------------------------------------------------------------
def list_items():
    """Return all active sellable items with calculated recipe cost."""
    with connect() as c:
        rows = c.execute("""
            SELECT
                i.id,
                i.sku,
                i.name,
                i.category,
                i.price,
                i.active,

                ROUND(
                    COALESCE(SUM(
                        r.qty_per_item *
                        CASE
                            WHEN inv.units_per_case > 0
                            THEN inv.case_cost / inv.units_per_case
                            ELSE 0
                        END
                    ), 0),
                2) AS item_cost

            FROM items i
            LEFT JOIN recipes r ON r.item_id = i.id
            LEFT JOIN components comp ON comp.id = r.component_id
            LEFT JOIN inventory inv ON inv.id = comp.inventory_id

            WHERE i.active = 1

            GROUP BY
                i.id,
                i.sku,
                i.name,
                i.category,
                i.price,
                i.active

            ORDER BY i.category, i.name
        """).fetchall()

        items = []

        for r in rows:
            item = dict(r)

            price = float(item.get("price") or 0)
            cost = float(item.get("item_cost") or 0)

            item["gross_profit"] = round(price - cost, 2)

            if price > 0:
                item["margin_percent"] = round(((price - cost) / price) * 100, 1)
            else:
                item["margin_percent"] = 0

            items.append(item)

        return items

def get_item_by_sku(conn, sku:str):
    row = conn.execute("SELECT id, price FROM items WHERE sku=? OR name=?", (sku, sku)).fetchone()
    return dict(row) if row else None

def add_item(data):
    with connect() as c:
        c.execute("""
            INSERT INTO items (sku, name, category, price, active)
            VALUES (?, ?, ?, ?, 1)
        """, (data['sku'], data['name'], data['category'], data['price']))
        c.commit()

def add_component(data):
    """Add or update a component linked to inventory by inventory_id."""
    inventory_id = int(data.get("inventory_id") or 0)
    display_name = (data.get("display_name") or "").strip() or None

    if inventory_id <= 0:
        raise ValueError("inventory_id is required.")

    with connect() as c:
        inv = c.execute("""
            SELECT id, name
            FROM inventory
            WHERE id = ?
            LIMIT 1
        """, (inventory_id,)).fetchone()

        if not inv:
            raise ValueError("Inventory item not found.")

        inventory_name = inv["name"]

        c.execute("""
            INSERT INTO components (
                name,
                display_name,
                inventory_id,
                pos_track_sellout
            )
            VALUES (?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
                display_name = COALESCE(excluded.display_name, components.display_name),
                inventory_id = excluded.inventory_id,
                pos_track_sellout = excluded.pos_track_sellout
        """, (
            inventory_name,
            display_name,
            inventory_id,
            int(data.get("pos_track_sellout", 0))
        ))

        c.commit()

# ---------------------------------------------------------------------
# Availability / Stock checks
# ---------------------------------------------------------------------
def available_qty(item_id: int) -> int:
    """Return how many of this item can be sold before components run out."""
    with connect() as c:
        rows = c.execute("""
            SELECT inv.qty_on_hand AS onhand, r.qty_per_item AS per
            FROM recipes r
            JOIN components comp ON comp.id = r.component_id
            JOIN inventory inv ON inv.name = comp.name
            WHERE r.item_id=? AND comp.pos_track_sellout=1
        """, (item_id,)).fetchall()

        if not rows:
            return 10**9  # no limiting components → unlimited

        limits = []
        for r in rows:
            per = float(r["per"] or 0)
            if per <= 0:
                continue
            onhand = float(r["onhand"] or 0)
            limits.append(int(onhand // per))

        return min(limits) if limits else 10**9

def sold_out_items():
    """List of items whose limiting component = 0."""
    with connect() as c:
        items = c.execute("SELECT id, name FROM items WHERE active=1").fetchall()
        sold_out = []
        for it in items:
            if available_qty(it["id"]) <= 0:
                sold_out.append(dict(it))
        return sold_out
    

# ---------------------------------------------------------------------
# Orders & Sales
# ---------------------------------------------------------------------
def record_order(cashier:str, method:str, discount:float, lines:list, discount_type:str=None,
                 discount_details:list=None,
                 tax_rate:float=0.0, cash_given:float=0.0, change_given:float=0.0):
    """Record a sale, update inventory, and update shift cash totals."""
    with connect() as c:
        ts = c.execute("SELECT datetime('now', '-6 hours')").fetchone()[0]

        shift_row = c.execute("SELECT id FROM shifts WHERE is_active=1").fetchone()
        shift_id = shift_row["id"] if shift_row else None
        
        # --- Build cart ---
        cart = []
        for l in lines:
            sku = l.get("item_sku")
            it = get_item_by_sku(c, sku)
            if not it:
                raise ValueError(f"Unknown item: {sku}")
            cart.append({
                "item_id": it["id"],
                "qty": float(l["qty"]),
                "unit_price": float(it["price"])
            })

        subtotal = sum(x["qty"] * x["unit_price"] for x in cart)
        taxable  = max(0.0, subtotal - float(discount or 0))
        tax      = round(taxable * tax_rate, 2)
        total    = taxable + tax
        rdate    = _today_date(ts)
        rseq     = _next_receipt_seq(rdate)

        # --- Create order ---
        cur = c.execute("""
            INSERT INTO orders(
                ts, cashier, method,
                discount, discount_type,
                tax, subtotal, total,
                cash_given, change_given,
                receipt_date, receipt_seq,
                shift_id
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ts,
            cashier,
            method,
            discount,
            discount_type,
            tax,
            subtotal,
            total,
            cash_given,
            change_given,
            rdate,
            rseq,
            shift_id
        ))

        order_id = cur.lastrowid
        
        # --- Record discount details ---
        for d in discount_details or []:
            amount = float(d.get("amount") or 0)

            if amount <= 0:
                continue

            c.execute("""
                INSERT INTO order_discounts (
                    order_id,
                    sale_id,
                    sale_name,
                    sale_type,
                    source,
                    amount,
                    authorized_pin_id,
                    authorized_pin_label
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                order_id,
                int(d["sale_id"]) if d.get("sale_id") else None,
                d.get("sale_name"),
                d.get("sale_type"),
                d.get("source", "MANUAL"),
                amount,
                int(d["authorized_pin_id"]) if d.get("authorized_pin_id") else None,
                d.get("authorized_pin_label")
            ))

        # --- Order lines ---
        for x in cart:
            c.execute("""
                INSERT INTO order_lines(order_id, item_id, qty, unit_price)
                VALUES (?, ?, ?, ?)
            """, (order_id, x["item_id"], x["qty"], x["unit_price"]))

        # --- Update inventory for each component ---
        rows = c.execute("""
            SELECT ol.item_id,
                   SUM(ol.qty) AS qty_sold,
                   inv.id      AS inv_id,
                   inv.name    AS inv_name,
                   r.qty_per_item AS per_item
            FROM order_lines ol
            JOIN recipes r ON r.item_id = ol.item_id
            JOIN components c2 ON c2.id = r.component_id
            JOIN inventory inv ON inv.name = c2.name
            WHERE ol.order_id = ?
            GROUP BY ol.item_id, inv.id, r.qty_per_item
        """, (order_id,)).fetchall()

        touched_items = set()
        for r in rows:
            touched_items.add(r["item_id"])
            used_qty = float(r["qty_sold"] or 0) * float(r["per_item"] or 0)
            if used_qty <= 0:
                continue
            c.execute("""
                UPDATE inventory
                SET qty_on_hand = MAX(qty_on_hand - ?, 0)
                WHERE id = ?
            """, (used_qty, r["inv_id"]))

        # --- Fallback for items with no recipe link ---
        leftovers = c.execute("""
            SELECT item_id, SUM(qty) AS qty
            FROM order_lines
            WHERE order_id = ?
            GROUP BY item_id
        """, (order_id,)).fetchall()

        for lf in leftovers:
            if lf["item_id"] in touched_items:
                continue
            qty = float(lf["qty"] or 0)
            if qty <= 0:
                continue
            lim = c.execute("""
                SELECT inv.id
                FROM recipes r
                JOIN components c3 ON c3.id = r.component_id
                JOIN inventory inv ON inv.name = c3.name
                WHERE r.item_id=? AND c3.pos_track_sellout=1
                ORDER BY c3.id LIMIT 1
            """, (lf["item_id"],)).fetchone()
            if lim:
                c.execute("""
                    UPDATE inventory
                    SET qty_on_hand = MAX(qty_on_hand - ?, 0)
                    WHERE id = ?
                """, (qty, lim["id"]))

        # --- Commit order + inventory changes ---
        c.commit()

        # --- Update shift totals if this was a cash order ---
        if method.lower() == "cash":
            try:
                c.execute("""
                    UPDATE shifts
                    SET
                        cash_sales   = COALESCE(cash_sales,0) + ?,
                        cash_given   = COALESCE(cash_given,0) + ?,
                        change_given = COALESCE(change_given,0) + ?,
                        last_sale_at = datetime('now', '-6 hours')
                    WHERE is_active = 1
                """, (total, cash_given, change_given))
                c.commit()
            except Exception as e:
                log(f"Could not update active shift totals: {e}")

    return {
        "order_id": order_id,
        "receipt_no": rseq,
        "date": rdate, 
        "total": total,
        "subtotal": subtotal,
        "tax": tax,
        "cash_given": cash_given,
        "change_given": change_given
    }

# ---------------------------------------------------------------------
# Sales / Promotions
# ---------------------------------------------------------------------

def list_sales(include_inactive: bool = False):
    """Return sales/promotions."""
    with connect() as c:
        where = "" if include_inactive else "WHERE s.active = 1"
        rows = c.execute(f"""
            SELECT
                s.id,
                s.name,
                s.sale_type,
                s.amount,
                s.item_id,
                s.category,
                s.buy_qty,
                s.free_qty,
                s.auto_apply,
                i.name AS item_name,
                s.active,
                s.starts_on,
                s.ends_on
            FROM sales s
            LEFT JOIN items i ON i.id = s.item_id
            {where}
            ORDER BY s.active DESC, s.name
        """).fetchall()

        sales = [dict(r) for r in rows]

        for s in sales:
            s["requirements"] = list_sale_requirements(s["id"])

        return sales


def list_active_sales():
    """Return only currently active sales."""
    today = datetime.now().date().isoformat()

    with connect() as c:
        rows = c.execute("""
            SELECT
                s.id,
                s.name,
                s.sale_type,
                s.amount,
                s.item_id,
                s.category,
                s.buy_qty,
                s.free_qty,
                s.auto_apply,
                i.name AS item_name,
                s.active,
                s.starts_on,
                s.ends_on
            FROM sales s
            LEFT JOIN items i ON i.id = s.item_id
            WHERE s.active = 1
              AND (s.starts_on IS NULL OR s.starts_on = '' OR s.starts_on <= ?)
              AND (s.ends_on IS NULL OR s.ends_on = '' OR s.ends_on >= ?)
            ORDER BY s.name
        """, (today, today)).fetchall()

        sales = [dict(r) for r in rows]

        for s in sales:
            s["requirements"] = list_sale_requirements(s["id"])

        return sales

def list_sale_requirements(sale_id: int):
    with connect() as c:
        rows = c.execute("""
            SELECT
                sr.id,
                sr.sale_id,
                sr.requirement_type,
                sr.item_id,
                i.name AS item_name,
                sr.category,
                sr.qty
            FROM sale_requirements sr
            LEFT JOIN items i ON i.id = sr.item_id
            WHERE sr.sale_id = ?
            ORDER BY sr.id
        """, (sale_id,)).fetchall()

        return [dict(r) for r in rows]
    
def replace_sale_requirements(sale_id: int, requirements: list):
    with connect() as c:
        c.execute("DELETE FROM sale_requirements WHERE sale_id=?", (sale_id,))

        for req in requirements or []:
            requirement_type = req.get("requirement_type", "CATEGORY")
            qty = int(req.get("qty") or 1)

            if qty < 1:
                qty = 1

            c.execute("""
                INSERT INTO sale_requirements (
                    sale_id,
                    requirement_type,
                    item_id,
                    category,
                    qty
                )
                VALUES (?, ?, ?, ?, ?)
            """, (
                sale_id,
                requirement_type,
                int(req["item_id"]) if req.get("item_id") else None,
                (req.get("category") or "").strip() or None,
                qty
            ))

        c.commit()
    
def add_sale(data):
    """Create a sale/promotion."""
    sale_type = data.get("sale_type", "DOLLAR_OFF")

    buy_qty = int(data.get("buy_qty") or 1)
    free_qty = int(data.get("free_qty") or 1)

    if buy_qty < 1:
        buy_qty = 1

    if free_qty < 1:
        free_qty = 1

    with connect() as c:
        cur = c.execute("""
            INSERT INTO sales (
                name,
                sale_type,
                amount,
                item_id,
                category,
                buy_qty,
                free_qty,
                auto_apply,
                active,
                starts_on,
                ends_on
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            (data.get("name") or "").strip(),
            sale_type,
            float(data.get("amount") or 0),
            int(data["item_id"]) if data.get("item_id") else None,
            (data.get("category") or "").strip() or None,
            buy_qty,
            free_qty,
            int(data.get("auto_apply", 0)),
            int(data.get("active", 1)),
            data.get("starts_on") or None,
            data.get("ends_on") or None
        ))

        sale_id = cur.lastrowid
        c.commit()

    if sale_type == "COMBO_PRICE":
        replace_sale_requirements(sale_id, data.get("requirements", []))
    
def update_sale(data):
    """Update an existing sale/promotion."""
    sale_type = data.get("sale_type", "DOLLAR_OFF")

    buy_qty = int(data.get("buy_qty") or 1)
    free_qty = int(data.get("free_qty") or 1)

    if buy_qty < 1:
        buy_qty = 1

    if free_qty < 1:
        free_qty = 1

    with connect() as c:
        c.execute("""
            UPDATE sales
            SET name=?,
                sale_type=?,
                amount=?,
                item_id=?,
                category=?,
                buy_qty=?,
                free_qty=?,
                auto_apply=?,
                active=?,
                starts_on=?,
                ends_on=?
            WHERE id=?
        """, (
            (data.get("name") or "").strip(),
            sale_type,
            float(data.get("amount") or 0),
            int(data["item_id"]) if data.get("item_id") else None,
            (data.get("category") or "").strip() or None,
            buy_qty,
            free_qty,
            int(data.get("auto_apply", 0)),
            int(data.get("active", 1)),
            data.get("starts_on") or None,
            data.get("ends_on") or None,
            int(data.get("id"))
        ))
        c.commit()
        
    if sale_type == "COMBO_PRICE":
        replace_sale_requirements(int(data.get("id")), data.get("requirements", []))
    else:
        replace_sale_requirements(int(data.get("id")), [])

def delete_sale(sale_id: int):
    """Delete a sale/promotion."""
    with connect() as c:
        c.execute("DELETE FROM sales WHERE id=?", (sale_id,))
        c.commit()
         
# ---------------------------------------------------------------------
# Admin helpers for adding new records
# ---------------------------------------------------------------------
def add_item(data):
    """Create or update a POS item safely (parentheses, special chars allowed)."""
    sku  = (data.get("sku") or "").strip()
    name = (data.get("name") or "").strip()
    cat  = (data.get("category") or "").strip()
    price = float(data.get("price") or 0)

    if not name:
        raise ValueError("Item name cannot be blank.")

    with connect() as c:
        # If SKU empty, auto-generate one from name
        if not sku:
            sku = name.replace(" ", "_").replace("(", "").replace(")", "").upper()[:12]

        c.execute("""
            INSERT INTO items (sku, name, category, price, active)
            VALUES (?, ?, ?, ?, 1)
            ON CONFLICT(sku) DO UPDATE SET
              name=excluded.name,
              category=excluded.category,
              price=excluded.price,
              active=1
        """, (sku, name, cat, price))
        c.commit()
        
def update_item(data):
    """Edit an existing item (name, category, price)."""
    with connect() as c:
        c.execute("""
            UPDATE items
            SET name=?, category=?, price=?
            WHERE id=?
        """, (
            data.get("name", "").strip(),
            data.get("category", "").strip(),
            float(data.get("price", 0)),
            int(data.get("id"))
        ))
        c.commit()

def delete_item(item_id:int):
    """Delete an item and its linked recipes."""
    with connect() as c:
        c.execute("DELETE FROM recipes WHERE item_id=?", (item_id,))
        c.execute("DELETE FROM items WHERE id=?", (item_id,))
        c.commit()

# ---------------------------------------------------------------------
# COMPONENTS (Usage Tracking)
# ---------------------------------------------------------------------
def list_components():
    """Return components with stable inventory link, unit cost, nickname, and recipe usage details."""
    with connect() as c:
        rows = c.execute("""
            SELECT 
                comp.id,
                comp.name,
                comp.display_name,
                comp.inventory_id,
                inv.name AS inventory_name,

                COALESCE(
                    CASE 
                        WHEN inv.units_per_case > 0
                        THEN inv.case_cost / inv.units_per_case
                        ELSE 0
                    END,
                0) AS unit_cost,

                comp.pos_track_sellout,

                COUNT(DISTINCT r.item_id) AS used_in_count

            FROM components comp
            LEFT JOIN inventory inv ON inv.id = comp.inventory_id
            LEFT JOIN recipes r ON r.component_id = comp.id

            GROUP BY
                comp.id,
                comp.name,
                comp.display_name,
                comp.inventory_id,
                inv.name,
                inv.case_cost,
                inv.units_per_case,
                comp.pos_track_sellout

            ORDER BY COALESCE(comp.display_name, inv.name, comp.name), comp.name
        """).fetchall()

        comps = []

        for r in rows:
            comp = dict(r)

            inventory_name = comp.get("inventory_name") or comp.get("name") or ""
            nickname = comp.get("display_name") or ""

            if nickname and nickname != inventory_name:
                comp["label"] = f"{inventory_name} — {nickname}"
            else:
                comp["label"] = inventory_name

            used_rows = c.execute("""
                SELECT 
                    i.id AS item_id,
                    i.name AS item_name,
                    r.qty_per_item
                FROM recipes r
                JOIN items i ON i.id = r.item_id
                WHERE r.component_id = ?
                ORDER BY i.name
            """, (comp["id"],)).fetchall()

            comp["used_in"] = [dict(x) for x in used_rows]
            comps.append(comp)

        return comps

def add_component(data):
    """Add or update a component linked to inventory by name."""
    name = (data.get("name") or "").strip()
    display_name = (data.get("display_name") or "").strip() or None

    with connect() as c:
        c.execute("""
            INSERT INTO components (name, display_name, pos_track_sellout)
            VALUES (?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
              display_name=COALESCE(excluded.display_name, components.display_name),
              pos_track_sellout=excluded.pos_track_sellout
        """, (
            name,
            display_name,
            int(data.get("pos_track_sellout", 0))
        ))
        c.commit()

def update_component(data):
    """Update one component's nickname and sold-out tracking."""
    with connect() as c:
        c.execute("""
            UPDATE components
            SET display_name=?,
                pos_track_sellout=?
            WHERE id=?
        """, (
            (data.get("display_name") or "").strip() or None,
            int(data.get("pos_track_sellout", 0)),
            int(data.get("id"))
        ))
        c.commit()

def delete_component(comp_id:int):
    """Delete a component and any linked recipes."""
    with connect() as c:
        c.execute("DELETE FROM recipes WHERE component_id=?", (comp_id,))
        c.execute("DELETE FROM components WHERE id=?", (comp_id,))
        c.commit()

def batch_update_components(components: list):
    """Update multiple component nicknames and sold-out tracking at once."""
    with connect() as c:
        for comp in components:
            comp_id = int(comp.get("id"))
            display_name = (comp.get("display_name") or "").strip()
            track = int(comp.get("pos_track_sellout") or 0)

            c.execute("""
                UPDATE components
                SET display_name = ?,
                    pos_track_sellout = ?
                WHERE id = ?
            """, (
                display_name or None,
                track,
                comp_id
            ))

        c.commit()

def _migrate_component_display_name():
    """Add display_name/nickname field to components."""
    with connect() as c:
        cols = [r["name"] for r in c.execute("PRAGMA table_info(components)").fetchall()]

        if "display_name" not in cols:
            c.execute("ALTER TABLE components ADD COLUMN display_name TEXT")

        c.commit()

    log("Migration: component display_name verified.")
    
# ---------------------------------------------------------------------
# RECIPES
# ---------------------------------------------------------------------
def add_recipe(data):
    """Link an item to a component using stable component_id."""
    with connect() as c:
        item = c.execute(
            "SELECT id FROM items WHERE sku=? OR name=? LIMIT 1",
            (data.get("item_key"), data.get("item_key"))
        ).fetchone()

        if data.get("component_id"):
            comp = c.execute(
                "SELECT id FROM components WHERE id=? LIMIT 1",
                (int(data.get("component_id")),)
            ).fetchone()
        else:
            # fallback for older frontend
            comp = c.execute(
                "SELECT id FROM components WHERE name=? LIMIT 1",
                (data.get("component_name"),)
            ).fetchone()

        if not item or not comp:
            raise ValueError("Unknown item or component")

        c.execute("""
            INSERT INTO recipes (item_id, component_id, qty_per_item)
            VALUES (?, ?, ?)
        """, (item["id"], comp["id"], float(data.get("qty_per_item", 1))))

        c.commit()

def update_recipe(data):
    """Update recipe quantity per item."""
    with connect() as c:
        c.execute("""
            UPDATE recipes
            SET qty_per_item=?
            WHERE id=?
        """, (float(data.get("qty_per_item", 1)), int(data.get("id"))))
        c.commit()

def delete_recipe(recipe_id:int):
    """Delete a recipe link."""
    with connect() as c:
        c.execute("DELETE FROM recipes WHERE id=?", (recipe_id,))
        c.commit()
        
def list_recipes():
    with connect() as c:
        rows = c.execute("""
            SELECT 
                r.id,
                i.name AS item_name,
                COALESCE(NULLIF(c.display_name, ''), c.name) AS component_name,
                c.name AS component_inventory_name,
                c.display_name AS component_display_name,
                r.qty_per_item,

                ROUND(
                    CASE 
                        WHEN inv.units_per_case > 0 
                        THEN inv.case_cost / inv.units_per_case 
                        ELSE 0 
                    END,
                4) AS unit_cost,

                ROUND(
                    r.qty_per_item *
                    CASE 
                        WHEN inv.units_per_case > 0 
                        THEN inv.case_cost / inv.units_per_case 
                        ELSE 0 
                    END,
                4) AS line_cost

            FROM recipes r
            JOIN items i ON i.id = r.item_id
            JOIN components c ON c.id = r.component_id
            LEFT JOIN inventory inv ON inv.id = c.inventory_id

            ORDER BY i.name, c.name
        """).fetchall()

        return [dict(r) for r in rows]

# ---------------------------------------------------------------------
# INVENTORY (Warehouse Stock)
# ---------------------------------------------------------------------
def list_inventory():
    """Return all inventory rows with calculated unit cost."""
    with connect() as c:
        rows = c.execute("""
            SELECT id,
                   name,
                   unit,
                   case_cost,
                   units_per_case,
                   CASE WHEN units_per_case > 0 
                        THEN ROUND(case_cost / units_per_case, 2)
                        ELSE 0 END AS unit_cost,
                   qty_on_hand,
                   reorder_point
            FROM inventory
            ORDER BY name
        """).fetchall()
        return [dict(r) for r in rows]

def add_inventory_item(data):
    """Add or update an inventory item."""
    with connect() as c:
        c.execute("""
            INSERT INTO inventory (name, unit, case_cost, units_per_case,
                                   qty_on_hand, reorder_point)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
              unit=excluded.unit,
              case_cost=excluded.case_cost,
              units_per_case=excluded.units_per_case,
              qty_on_hand=excluded.qty_on_hand,
              reorder_point=excluded.reorder_point
        """, (
            data.get("name", "").strip(),
            data.get("unit", "").strip(),
            float(data.get("case_cost", 0)),
            float(data.get("units_per_case", 0)),
            float(data.get("qty_on_hand", 0)),
            float(data.get("reorder_point", 0))
        ))
        c.commit()

def update_inventory_item(data):
    """Edit an existing inventory record."""
    with connect() as c:
        c.execute("""
            UPDATE inventory
            SET name=?,
                unit=?,
                case_cost=?,
                units_per_case=?,
                qty_on_hand=?,
                reorder_point=?
            WHERE id=?
        """, (
            data.get("name", "").strip(),
            data.get("unit", "").strip(),
            float(data.get("case_cost", 0)),
            float(data.get("units_per_case", 0)),
            float(data.get("qty_on_hand", 0)),
            float(data.get("reorder_point", 0)),
            int(data.get("id"))
        ))
        c.commit()

def delete_inventory_item(item_id:int):
    """Remove an inventory record."""
    with connect() as c:
        c.execute("DELETE FROM inventory WHERE id=?", (item_id,))
        c.commit()

def _migrate_component_inventory_id():
    """Add stable inventory_id link to components and backfill existing rows by name."""
    with connect() as c:
        cols = [r["name"] for r in c.execute("PRAGMA table_info(components)").fetchall()]

        if "inventory_id" not in cols:
            c.execute("ALTER TABLE components ADD COLUMN inventory_id INTEGER")

        # Backfill existing components by matching old component name to inventory name.
        c.execute("""
            UPDATE components
            SET inventory_id = (
                SELECT inventory.id
                FROM inventory
                WHERE inventory.name = components.name
                LIMIT 1
            )
            WHERE inventory_id IS NULL
        """)

        c.commit()

    log("Migration: component inventory_id verified.")

# --- EXPORT HELPERS (add to db_ops.py) ---

_ALLOWED_TABLES = {"items", "components", "recipes", "inventory"}

def _query_all(conn, table):
    return conn.execute(f"SELECT * FROM {table}").fetchall(), [d[0] for d in conn.execute(f"PRAGMA table_info({table})")]

def export_selected_as_zip_csv(tables: list[str]) -> bytes:
    """
    Return a ZIP (bytes) containing one CSV per selected table.
    """
    with connect() as c:
        mem = io.BytesIO()
        with zipfile.ZipFile(mem, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for t in tables:
                if t not in _ALLOWED_TABLES: 
                    continue
                rows = c.execute(f"SELECT * FROM {t}").fetchall()
                cols = [col[1] for col in c.execute(f"PRAGMA table_info({t})")]
                # Write CSV to an in-memory buffer
                buf = io.StringIO()
                writer = csv.writer(buf)
                writer.writerow(cols)
                for r in rows:
                    writer.writerow([r[k] if isinstance(r, sqlite3.Row) else r[i] for i, k in enumerate(cols)])
                zf.writestr(f"{t}.csv", buf.getvalue())
        return mem.getvalue()

def export_selected_as_sql(tables: list[str]) -> str:
    """
    Return a SQL script that recreates ONLY the selected tables' schema + data.
    (This uses sqlite's schema + INSERTs for selected tables.)
    """
    with connect() as c:
        out = io.StringIO()
        # Write schema for selected tables
        for t in tables:
            if t not in _ALLOWED_TABLES:
                continue
            schema = c.execute(
                "SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (t,)
            ).fetchone()
            if schema and schema["sql"]:
                out.write(schema["sql"] + ";\n\n")
        # Write INSERTs
        for t in tables:
            if t not in _ALLOWED_TABLES:
                continue
            rows = c.execute(f"SELECT * FROM {t}").fetchall()
            cols = [col[1] for col in c.execute(f"PRAGMA table_info({t})")]
            for r in rows:
                vals = []
                for k in cols:
                    v = r[k]
                    if v is None:
                        vals.append("NULL")
                    elif isinstance(v, (int, float)):
                        vals.append(str(v))
                    else:
                        # escape single quotes
                        s = str(v).replace("'", "''")
                        vals.append(f"'{s}'")
                out.write(f"INSERT INTO {t} ({', '.join(cols)}) VALUES ({', '.join(vals)});\n")
            out.write("\n")
        return out.getvalue()

CATALOG_TABLES = [
    "inventory",
    "items",
    "components",
    "recipes",
    "sales",
    "sale_requirements",
    "settings",
]

FULL_BACKUP_TABLES = [
    "inventory",
    "items",
    "components",
    "recipes",
    "sales",
    "sale_requirements",
    "settings",
    "admin_pins",
    "orders",
    "order_lines",
    "order_discounts",
    "shifts",
    "expenses",
]


def _table_exists(c, table_name: str) -> bool:
    row = c.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type='table' AND name=?
    """, (table_name,)).fetchone()

    return row is not None


def _get_export_tables(export_type: str) -> list[str]:
    if export_type == "full":
        return FULL_BACKUP_TABLES

    return CATALOG_TABLES


def _exports_dir() -> str:
    base = os.path.join(os.getcwd(), "exports")
    os.makedirs(base, exist_ok=True)
    return base


def export_database(export_type: str = "catalog", fmt: str = "csv") -> str:
    """
    Export catalog setup or full database backup.

    export_type:
      catalog = setup data only
      full    = setup + operational records

    fmt:
      csv = ZIP of CSV files
      sql = one SQL dump file
    """
    export_type = export_type if export_type in ("catalog", "full") else "catalog"
    fmt = fmt if fmt in ("csv", "sql") else "csv"

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = _exports_dir()

    if fmt == "sql":
        filename = f"brewins_{export_type}_backup_{stamp}.sql"
        path = os.path.join(out_dir, filename)
        export_database_sql(path, export_type)
        return path

    filename = f"brewins_{export_type}_backup_{stamp}.zip"
    path = os.path.join(out_dir, filename)
    export_database_csv_zip(path, export_type)
    return path


def export_database_csv_zip(path: str, export_type: str = "catalog"):
    tables = _get_export_tables(export_type)

    with connect() as c:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            exported_tables = []

            for table in tables:
                if not _table_exists(c, table):
                    continue

                exported_tables.append(table)

                rows = c.execute(f'SELECT * FROM "{table}"').fetchall()
                cols = [r["name"] for r in c.execute(f'PRAGMA table_info("{table}")').fetchall()]

                csv_buffer = io.StringIO()
                writer = csv.writer(csv_buffer)

                writer.writerow(cols)

                for row in rows:
                    writer.writerow([row[col] for col in cols])

                zf.writestr(f"{table}.csv", csv_buffer.getvalue())

            manifest_buffer = io.StringIO()
            writer = csv.writer(manifest_buffer)

            writer.writerow(["field", "value"])
            writer.writerow(["export_type", export_type])
            writer.writerow(["created_at", datetime.now().isoformat(timespec="seconds")])
            writer.writerow(["format", "csv_zip"])
            writer.writerow(["tables", ", ".join(exported_tables)])

            zf.writestr("_manifest.csv", manifest_buffer.getvalue())


def export_database_sql(path: str, export_type: str = "catalog"):
    tables = _get_export_tables(export_type)

    with connect() as c:
        with open(path, "w", encoding="utf-8") as f:
            f.write("-- BrewIns POS Export\n")
            f.write(f"-- Export type: {export_type}\n")
            f.write(f"-- Created at: {datetime.now().isoformat(timespec='seconds')}\n\n")

            f.write("PRAGMA foreign_keys=OFF;\n")
            f.write("BEGIN TRANSACTION;\n\n")

            for table in tables:
                if not _table_exists(c, table):
                    continue

                schema = c.execute("""
                    SELECT sql
                    FROM sqlite_master
                    WHERE type='table' AND name=?
                """, (table,)).fetchone()

                if schema and schema["sql"]:
                    f.write(f'DROP TABLE IF EXISTS "{table}";\n')
                    f.write(schema["sql"] + ";\n\n")

                rows = c.execute(f'SELECT * FROM "{table}"').fetchall()
                cols = [r["name"] for r in c.execute(f'PRAGMA table_info("{table}")').fetchall()]

                for row in rows:
                    values = [row[col] for col in cols]
                    sql = _make_insert_sql(table, cols, values)
                    f.write(sql + "\n")

                f.write("\n")

            f.write("COMMIT;\n")
            f.write("PRAGMA foreign_keys=ON;\n")


def _make_insert_sql(table: str, cols: list[str], values: list) -> str:
    col_list = ", ".join([f'"{col}"' for col in cols])
    value_list = ", ".join([_sql_literal(v) for v in values])

    return f'INSERT INTO "{table}" ({col_list}) VALUES ({value_list});'


def _sql_literal(value):
    if value is None:
        return "NULL"

    if isinstance(value, (int, float)):
        return str(value)

    text = str(value).replace("'", "''")
    return f"'{text}'"

def _restore_uploads_dir() -> str:
    base = os.path.join(os.getcwd(), "restore_uploads")
    os.makedirs(base, exist_ok=True)
    return base


def _safety_backups_dir() -> str:
    base = os.path.join(os.getcwd(), "safety_backups")
    os.makedirs(base, exist_ok=True)
    return base

def restore_database_from_sql(sql_path: str, restore_type: str = "catalog") -> dict:
    """
    Restore database content from a SQL export.

    restore_type:
      catalog = restore setup/catalog tables only from SQL export
      full    = restore full SQL backup

    This assumes the SQL file came from our export system.
    """
    restore_type = restore_type if restore_type in ("catalog", "full") else "catalog"

    safety_backup = make_safety_db_backup()

    with open(sql_path, "r", encoding="utf-8") as f:
        sql_text = f.read()

    with connect() as c:
        if restore_type == "full":
            # Full restore runs the SQL dump as-is.
            c.executescript(sql_text)
            c.commit()

        else:
            # Catalog restore should only wipe/rebuild catalog tables.
            # Since our catalog SQL export only contains catalog tables,
            # running it is safe as long as the uploaded file is truly catalog export.
            c.executescript(sql_text)
            c.commit()

    return {
        "ok": True,
        "restore_type": restore_type,
        "restored_from": sql_path,
        "safety_backup": safety_backup
    }
# ---------------------------------------------------------------------
# Reports / KPIs
# ---------------------------------------------------------------------
def sales_summary():
    """Daily gross totals for /reports screen."""
    with connect() as c:
        rows = c.execute("""
            SELECT DATE(ts) AS date,
                   COUNT(*) AS orders,
                   ROUND(SUM(total), 2) AS gross
            FROM orders
            GROUP BY DATE(ts)
            ORDER BY date DESC
        """).fetchall()
        return [dict(r) for r in rows]

def kpis(start_date:str, end_date:str):
    """Detailed KPI block (gross, tax, cogs, net)."""
    with connect() as c:
        orders = c.execute("""
            SELECT id, subtotal, discount, tax
            FROM orders
            WHERE DATE(ts) BETWEEN ? AND ?
        """, (start_date, end_date)).fetchall()

        gross = sum(o["subtotal"] - float(o["discount"] or 0) for o in orders)
        tax   = sum(o["tax"] for o in orders)
        cogs  = 0.0
        for o in orders:
            lines = c.execute("SELECT id, item_id, qty FROM order_lines WHERE order_id=?", (o["id"],)).fetchall()
            for ln in lines:
                recs = c.execute("""
                    SELECT r.qty_per_item, c.unit_cost
                    FROM recipes r
                    JOIN components c ON c.id=r.component_id
                    WHERE r.item_id=?
                """, (ln["item_id"],)).fetchall()
                for rr in recs:
                    cogs += float(ln["qty"]) * float(rr["qty_per_item"]) * float(rr["unit_cost"])

        gp = gross - cogs
        return dict(gross=round(gross,2), tax=round(tax,2), cogs=round(cogs,2), gp=round(gp,2))
    
# --- REPORTS: Orders + Receipts ---

def list_orders_by_date(date_str: str):
    with connect() as c:
        rows = c.execute("""
            SELECT id, ts, cashier, method, subtotal, discount, tax, total,
                   COALESCE(receipt_seq, id) AS receipt_no
            FROM orders
            WHERE DATE(ts) = ?
            ORDER BY receipt_no ASC, id ASC
        """, (date_str,)).fetchall()
        return [dict(r) for r in rows]

def get_receipt(order_id: int):
    with connect() as c:
        order = c.execute("""
            SELECT id, ts, cashier, method,
                   subtotal,
                   discount,
                   discount_type,
                   tax,
                   total,
                   COALESCE(receipt_seq, id) AS receipt_no,
                   COALESCE(receipt_date, substr(ts,1,10)) AS receipt_date
            FROM orders
            WHERE id = ?
        """, (order_id,)).fetchone()

        if not order:
            return None
        lines = c.execute("""
            SELECT ol.id, ol.item_id, i.name AS item_name, ol.qty, ol.unit_price,
                   ROUND(ol.qty * ol.unit_price, 2) AS line_total
            FROM order_lines ol
            JOIN items i ON i.id = ol.item_id
            WHERE ol.order_id = ?
            ORDER BY ol.id ASC
        """, (order_id,)).fetchall()    
        discounts = c.execute("""
            SELECT id,
                sale_id,
                sale_name,
                sale_type,
                source,
                amount,
                authorized_pin_id,
                authorized_pin_label,
                created_at
            FROM order_discounts
            WHERE order_id = ?
            ORDER BY id ASC
        """, (order_id,)).fetchall()
        return {
            "order": dict(order),
            "lines": [dict(r) for r in lines],
            "discounts": [dict(r) for r in discounts]
        }
        


def _today_date(ts_iso: str | None = None) -> str:
    # ts_iso like '2025-10-11T14:22:00'; default now
    if not ts_iso:
        return datetime.now().date().isoformat()
    return ts_iso[:10]

def _next_receipt_seq(date_str: str) -> int:
    with connect() as c:
        row = c.execute("SELECT COALESCE(MAX(receipt_seq),0)+1 AS n FROM orders WHERE receipt_date = ?", (date_str,)).fetchone()
        return int(row["n"] or 1)


# ------- Delete Order and Restore Inventory -------
def delete_order_and_restore(order_id: int) -> bool:
    """
    Restores inventory quantities consumed by the order (per recipes),
    then deletes the order (order_lines are ON DELETE CASCADE).
    """
    with connect() as c:
        # 1) Gather usage by inventory.name for this order
        rows = c.execute("""
            SELECT inv.name AS inv_name,
                   SUM(ol.qty * r.qty_per_item) AS used_qty
            FROM order_lines ol
            JOIN recipes r       ON r.item_id = ol.item_id
            JOIN components comp ON comp.id   = r.component_id
            JOIN inventory inv   ON inv.name  = comp.name
            WHERE ol.order_id = ?
            GROUP BY inv.name
        """, (order_id,)).fetchall()

        # 2) Restore inventory (add back used qty)
        for row in rows:
            used = float(row["used_qty"] or 0)
            if used <= 0:
                continue
            c.execute("""
                UPDATE inventory
                SET qty_on_hand = qty_on_hand + ?
                WHERE name = ?
            """, (used, row["inv_name"]))

        # 3) Delete order (order_lines will be removed via FK cascade)
        c.execute("DELETE FROM orders WHERE id=?", (order_id,))
        c.commit()
        return True

# ------- Daily Sales Summary + Export to XLSX -------


def daily_sales_summary(date_str: str):
    """Return item sales, discounts, COGS, and profit summary for a given date."""
    with connect() as c:
        # Item-level sales before discounts
        lines = c.execute("""
            SELECT i.id AS item_id,
                   i.name AS item_name,
                   SUM(ol.qty) AS qty_sold,
                   i.price AS unit_price
            FROM orders o
            JOIN order_lines ol ON o.id = ol.order_id
            JOIN items i ON i.id = ol.item_id
            WHERE DATE(o.ts) = ?
            GROUP BY i.id
            ORDER BY i.name
        """, (date_str,)).fetchall()

        items = []
        pre_discount_sales = 0.0
        cogs_total = 0.0

        for l in lines:
            item = dict(l)
            item_id = item["item_id"]
            qty_sold = float(item["qty_sold"] or 0)
            unit_price = float(item["unit_price"] or 0)
            gross = qty_sold * unit_price

            comps = c.execute("""
                SELECT r.qty_per_item,
                       CASE WHEN inv.units_per_case > 0
                            THEN inv.case_cost / inv.units_per_case
                            ELSE 0 END AS unit_cost
                FROM recipes r
                JOIN components comp ON comp.id = r.component_id
                JOIN inventory inv ON inv.name = comp.name
                WHERE r.item_id = ?
            """, (item_id,)).fetchall()

            cogs = sum(
                qty_sold * float(r["qty_per_item"] or 0) * float(r["unit_cost"] or 0)
                for r in comps
            )

            items.append({
                "name": item["item_name"],
                "qty_sold": qty_sold,
                "unit_price": unit_price,
                "gross": round(gross, 2),
                "cogs": round(cogs, 2)
            })

            pre_discount_sales += gross
            cogs_total += cogs

        # Order-level totals
        order_totals = c.execute("""
            SELECT
                COALESCE(SUM(subtotal), 0) AS subtotal_total,
                COALESCE(SUM(discount), 0) AS discount_total,
                COALESCE(SUM(tax), 0) AS tax_total,
                COALESCE(SUM(total), 0) AS net_sales
            FROM orders
            WHERE DATE(ts) = ?
        """, (date_str,)).fetchone()

        subtotal_total = float(order_totals["subtotal_total"] or 0)
        discount_total = float(order_totals["discount_total"] or 0)
        tax_total = float(order_totals["tax_total"] or 0)
        net_sales = float(order_totals["net_sales"] or 0)

        # Discount details by source/sale
        discount_rows = c.execute("""
            SELECT
                COALESCE(source, 'UNKNOWN') AS source,
                COALESCE(sale_name, sale_type, 'Discount') AS sale_name,
                COALESCE(SUM(amount), 0) AS amount
            FROM order_discounts od
            JOIN orders o ON o.id = od.order_id
            WHERE DATE(o.ts) = ?
            GROUP BY source, sale_name
            ORDER BY source, sale_name
        """, (date_str,)).fetchall()

        discounts = [dict(r) for r in discount_rows]

        # Fallback if old orders have orders.discount but no order_discounts rows
        detailed_discount_total = sum(float(d["amount"] or 0) for d in discounts)
        missing_discount = round(discount_total - detailed_discount_total, 2)

        if missing_discount > 0:
            discounts.append({
                "source": "LEGACY",
                "sale_name": "Older discount record",
                "amount": missing_discount
            })

        gp = net_sales - cogs_total

        exp = c.execute("""
            SELECT COALESCE(SUM(amount),0) AS total
            FROM expenses
            WHERE date=?
        """, (date_str,)).fetchone()["total"] or 0.0

        np = gp - float(exp)

        return {
            "date": date_str,
            "items": items,

            "pre_discount_sales": round(pre_discount_sales, 2),
            "gross_total": round(pre_discount_sales, 2),  # keeps old frontend compatible

            "discount_total": round(discount_total, 2),
            "discounts": discounts,

            "tax_total": round(tax_total, 2),
            "net_sales": round(net_sales, 2),

            "cogs_total": round(cogs_total, 2),
            "gp": round(gp, 2),

            "expenses": round(float(exp), 2),
            "net_profit": round(np, 2)
        }


def export_sales_to_xlsx(date_str: str) -> bytes:
    """Return an Excel file (bytes) for that day’s sales summary."""
    data = daily_sales_summary(date_str)
    wb = Workbook()
    ws = wb.active
    ws.title = f"Sales {date_str}"

    ws.append(["Item", "Qty Sold", "Unit Price", "Pre-Discount Sales", "COGS"])
    for it in data["items"]:
        ws.append([it["name"], it["qty_sold"], it["unit_price"], it["gross"], it["cogs"]])

    ws.append([])
    ws.append(["Pre-Discount Sales", data["pre_discount_sales"]])
    ws.append(["Discounts", -data["discount_total"]])

    if data.get("tax_total", 0) > 0:
        ws.append(["Tax", data["tax_total"]])

    ws.append(["Net Sales", data["net_sales"]])
    ws.append(["COGS Total", data["cogs_total"]])
    ws.append(["Gross Profit", data["gp"]])
    ws.append(["Operating Expenses", data["expenses"]])
    ws.append(["Net Profit", data["net_profit"]])

    if data.get("discounts"):
        ws.append([])
        ws.append(["Discount Details"])
        ws.append(["Source", "Sale", "Amount"])
        for d in data["discounts"]:
            ws.append([
                d.get("source"),
                d.get("sale_name"),
                -float(d.get("amount") or 0)
            ])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

#------ Extended Sales Summary for Date Range + Export to XLSX -------


def sales_summary_range(start_date: str, end_date: str):
    """
    Aggregate item sales, discounts, COGS, and profit for DATE(ts)
    BETWEEN start_date AND end_date inclusive.
    """
    with connect() as c:
        rows = c.execute("""
            SELECT i.id AS item_id,
                   i.name AS item_name,
                   i.price AS unit_price,
                   SUM(ol.qty) AS qty_sold
            FROM orders o
            JOIN order_lines ol ON o.id = ol.order_id
            JOIN items i ON i.id = ol.item_id
            WHERE DATE(o.ts) BETWEEN ? AND ?
            GROUP BY i.id
            ORDER BY i.name
        """, (start_date, end_date)).fetchall()

        items = []
        pre_discount_sales = 0.0
        cogs_total = 0.0

        for r in rows:
            qty_sold = float(r["qty_sold"] or 0)
            unit_price = float(r["unit_price"] or 0)
            gross = qty_sold * unit_price

            comps = c.execute("""
                SELECT r.qty_per_item,
                       CASE WHEN inv.units_per_case > 0
                            THEN inv.case_cost / inv.units_per_case
                            ELSE 0 END AS unit_cost
                FROM recipes r
                JOIN components comp ON comp.id = r.component_id
                JOIN inventory inv ON inv.name = comp.name
                WHERE r.item_id = ?
            """, (r["item_id"],)).fetchall()

            cogs = sum(
                qty_sold * float(x["qty_per_item"] or 0) * float(x["unit_cost"] or 0)
                for x in comps
            )

            items.append({
                "name": r["item_name"],
                "qty_sold": qty_sold,
                "unit_price": unit_price,
                "gross": round(gross, 2),
                "cogs": round(cogs, 2),
            })

            pre_discount_sales += gross
            cogs_total += cogs

        order_totals = c.execute("""
            SELECT
                COALESCE(SUM(subtotal), 0) AS subtotal_total,
                COALESCE(SUM(discount), 0) AS discount_total,
                COALESCE(SUM(tax), 0) AS tax_total,
                COALESCE(SUM(total), 0) AS net_sales
            FROM orders
            WHERE DATE(ts) BETWEEN ? AND ?
        """, (start_date, end_date)).fetchone()

        subtotal_total = float(order_totals["subtotal_total"] or 0)
        discount_total = float(order_totals["discount_total"] or 0)
        tax_total = float(order_totals["tax_total"] or 0)
        net_sales = float(order_totals["net_sales"] or 0)

        discount_rows = c.execute("""
            SELECT
                COALESCE(source, 'UNKNOWN') AS source,
                COALESCE(sale_name, sale_type, 'Discount') AS sale_name,
                COALESCE(SUM(amount), 0) AS amount
            FROM order_discounts od
            JOIN orders o ON o.id = od.order_id
            WHERE DATE(o.ts) BETWEEN ? AND ?
            GROUP BY source, sale_name
            ORDER BY source, sale_name
        """, (start_date, end_date)).fetchall()

        discounts = [dict(r) for r in discount_rows]

        detailed_discount_total = sum(float(d["amount"] or 0) for d in discounts)
        missing_discount = round(discount_total - detailed_discount_total, 2)

        if missing_discount > 0:
            discounts.append({
                "source": "LEGACY",
                "sale_name": "Older discount record",
                "amount": missing_discount
            })

        gp = net_sales - cogs_total

        exp = c.execute("""
            SELECT COALESCE(SUM(amount),0) AS total
            FROM expenses
            WHERE date BETWEEN ? AND ?
        """, (start_date, end_date)).fetchone()["total"] or 0.0

        np = gp - float(exp)

        return {
            "start": start_date,
            "end": end_date,
            "items": items,

            "pre_discount_sales": round(pre_discount_sales, 2),
            "gross_total": round(pre_discount_sales, 2),

            "discount_total": round(discount_total, 2),
            "discounts": discounts,

            "tax_total": round(tax_total, 2),
            "net_sales": round(net_sales, 2),

            "cogs_total": round(cogs_total, 2),
            "gp": round(gp, 2),

            "expenses": round(float(exp), 2),
            "net_profit": round(np, 2),
        }

def export_sales_range_to_xlsx(start_date: str, end_date: str) -> bytes:
    data = sales_summary_range(start_date, end_date)
    wb = Workbook()

    # Sheet 1: Summary
    ws1 = wb.active
    ws1.title = "Summary"
    ws1.append([f"Sales Summary {data['start']} to {data['end']}"])
    ws1.append([])
    ws1.append(["Pre-Discount Sales", data["pre_discount_sales"]])
    ws1.append(["Discounts", -data["discount_total"]])

    if data.get("tax_total", 0) > 0:
        ws1.append(["Tax", data["tax_total"]])

    ws1.append(["Net Sales", data["net_sales"]])
    ws1.append(["COGS Total", data["cogs_total"]])
    ws1.append(["Gross Profit", data["gp"]])
    ws1.append(["Operating Expenses", data["expenses"]])
    ws1.append(["Net Profit", data["net_profit"]])

    if data.get("discounts"):
        ws1.append([])
        ws1.append(["Discount Details"])
        ws1.append(["Source", "Sale", "Amount"])
        for d in data["discounts"]:
            ws1.append([
                d.get("source"),
                d.get("sale_name"),
                -float(d.get("amount") or 0)
            ])

    # Sheet 2: By Item
    ws2 = wb.create_sheet("By Item")
    ws2.append(["Item", "Qty Sold", "Unit Price", "Pre-Discount Sales", "COGS"])
    for it in data["items"]:
        ws2.append([it["name"], it["qty_sold"], it["unit_price"], it["gross"], it["cogs"]])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ---------------------------------------------------------------------
# Shift summaries for admin reports
# ---------------------------------------------------------------------
def list_shifts(start_date: str = None, end_date: str = None) -> list[dict]:
    """Return shift history with key financial details."""
    with connect() as c:
        if start_date and end_date:
            rows = c.execute("""
                SELECT id,
                       DATE(ts_start) AS date,
                       cashier,
                       opening_float,
                       closing_amount,
                       over_short,
                       ROUND((closing_amount - opening_float), 2) AS net_change,
                       ts_start,
                       ts_end,
                       is_active
                FROM shifts
                WHERE DATE(ts_start) BETWEEN ? AND ?
                ORDER BY ts_start DESC
            """, (start_date, end_date)).fetchall()
        else:
            rows = c.execute("""
                SELECT id,
                       DATE(ts_start) AS date,
                       cashier,
                       opening_float,
                       closing_amount,
                       over_short,
                       ROUND((closing_amount - opening_float), 2) AS net_change,
                       ts_start,
                       ts_end,
                       is_active
                FROM shifts
                ORDER BY ts_start DESC
            """).fetchall()
        return [dict(r) for r in rows]

# ---------------------------------------------------------------------
# Settings and Discount Pins
# ---------------------------------------------------------------------
def get_setting(key: str, default=None):
    with connect() as c:
        row = c.execute(
            "SELECT value FROM settings WHERE key=?",
            (key,)
        ).fetchone()

        return row["value"] if row else default


def set_setting(key: str, value):
    with connect() as c:
        c.execute("""
            INSERT INTO settings(key, value)
            VALUES (?, ?)
            ON CONFLICT(key)
            DO UPDATE SET value=excluded.value
        """, (key, str(value)))
        c.commit()


def get_pos_settings():
    return {
        "require_pin_discounts": get_setting("require_pin_discounts", "0") == "1",
        "require_pin_admin_access": get_setting("require_pin_admin_access", "0") == "1",
        "require_pin_exit_program": get_setting("require_pin_exit_program", "0") == "1",
        "tax_rate": float(get_setting("tax_rate", "0") or 0),
    }


def list_admin_pins():
    with connect() as c:
        rows = c.execute("""
            SELECT id, label, active
            FROM admin_pins
            ORDER BY active DESC, id ASC
        """).fetchall()

        return [dict(r) for r in rows]


def add_admin_pin(pin: str, label: str = ""):
    with connect() as c:
        c.execute("""
            INSERT INTO admin_pins(pin, label, active)
            VALUES (?, ?, 1)
        """, (pin.strip(), label.strip()))
        c.commit()


def delete_admin_pin(pin_id: int):
    with connect() as c:
        c.execute(
            "DELETE FROM admin_pins WHERE id=?",
            (pin_id,)
        )
        c.commit()


def validate_admin_pin(pin: str) -> bool:
    with connect() as c:
        row = c.execute("""
            SELECT id
            FROM admin_pins
            WHERE pin=? AND active=1
            LIMIT 1
        """, (pin.strip(),)).fetchone()

        return row is not None

def validate_admin_pin_info(pin: str):
    """
    Return PIN info if valid, otherwise None.
    Used when we need to record which labeled PIN authorized a manual discount.
    """
    with connect() as c:
        row = c.execute("""
            SELECT id, label
            FROM admin_pins
            WHERE pin=? AND active=1
            LIMIT 1
        """, (pin.strip(),)).fetchone()

        return dict(row) if row else None
    
def has_admin_pins() -> bool:
    """Return True if at least one active admin PIN exists."""
    with connect() as c:
        row = c.execute("""
            SELECT id
            FROM admin_pins
            WHERE active = 1
            LIMIT 1
        """).fetchone()

        return row is not None