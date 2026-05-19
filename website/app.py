from flask import Flask, render_template, jsonify, request, send_file, make_response, session, redirect, url_for
from backend import db_ops as ops
import os, sys, io, time
from datetime import datetime
from pathlib import Path
import os
from werkzeug.utils import secure_filename
import traceback

from webview import settings

# ---------------------------------------------------------------------
# SAFE UTF-8 LOGGING (NO print())
# ---------------------------------------------------------------------

def _log_path():
    local = os.environ.get("LOCALAPPDATA", "")
    log_dir = os.path.join(local, "BrewInsPOS")
    os.makedirs(log_dir, exist_ok=True)
    return os.path.join(log_dir, "pos.log")

LOG_FILE = _log_path()

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except:
        pass
    
APP_LOG = _log_path()

def app_log(msg: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(APP_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except:
        pass
    
def debug_paths():
    base = None
    if getattr(sys, 'frozen', False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent

    exe_dir = Path(sys.executable).resolve().parent
    cwd = Path(os.getcwd())

    log(f"[PATH] frozen={getattr(sys,'frozen',False)}")
    log(f"[PATH] sys._MEIPASS={getattr(sys,'_MEIPASS',None)}")
    log(f"[PATH] exe_dir={exe_dir}")
    log(f"[PATH] cwd={cwd}")
    log(f"[PATH] file_dir={Path(__file__).resolve().parent}")
    log(f"[PATH] base_chosen={base}")

debug_paths()

EXPORT_DIR = Path(os.environ["LOCALAPPDATA"]) / "BrewInsPOS" / "exports"
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------
# RUN MIGRATIONS SAFELY
# ---------------------------------------------------------------------
try:
    ops.initialize_database()
except Exception as e:
    log(f"Database initialization failed: {e}")
    
try:
    ops._migrate_receipt_seq()
    ops._migrate_shifts_table()
    ops._migrate_orders_table()
    ops._migrate_order_cash_columns()
    ops._migrate_shift_last_sale()
    ops._migrate_discount_type()
    ops._migrate_settings_tables()
    ops._migrate_sales_table()
    ops._migrate_order_discounts_table()
    ops._migrate_order_line_modifiers_table()
    ops._migrate_component_display_name()
    ops._migrate_component_inventory_id()
    ops._migrate_till_counts_table()
    ops._migrate_item_pos_group_fields()
    ops._migrate_flavors_table()
except Exception as e:
    log(f"Schema migration check failed: {e}")

try:
    ops._auto_close_stale_shifts()
except Exception as e:
    log(f"Auto-close stale shifts failed: {e}")


# ---------------------------------------------------------------------
# Flask Setup
# ---------------------------------------------------------------------
app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.secret_key = "brewins-pos-local-secret-key"

# ---------- DEBUGGING ENDPOINTS ----------
@app.route('/api/debug/components')
def api_debug_components():
    return jsonify(ops.debug_bad_components())

# ---------- ROUTES ----------
@app.route('/splash')
def splash():
    app_log("GET /splash")
    return render_template('splash.html')

@app.route('/')
def index():
    return render_template('pos.html')

@app.route('/pos')
def pos_page():
    app_log("GET /pos")
    return render_template('pos.html')


# ---------- ADMIN PAGES ----------
@app.route('/admin/inventory')
def inventory_page():
    return render_template('admin/inventory.html')

@app.route('/admin/items')
def admin_items_page():
    return render_template('admin/items.html')

@app.route('/admin/components')
def admin_components_page():
    return render_template('admin/components.html')

@app.route('/admin/recipes')
def admin_recipes_page():
    return render_template('admin/recipes.html')


@app.route('/admin/flavors')
def admin_flavors_page():
    return render_template('admin/flavors.html')

# ---------- REPORTS PAGES ----------
@app.route('/reports/receipts')
def reports_receipts():
    return render_template('reports/receipts.html')

@app.route('/reports/sales')
def reports_sales():
    return render_template('reports/sales.html')

# ---------- SHIFT API ----------
@app.route('/api/shift/start', methods=['POST'])
def api_shift_start():
    data = request.json or {}
    cashier = data.get("cashier")
    till_count = data.get("till_count")

    if till_count is not None:
        result = ops.start_shift(cashier, till_count=till_count)
    else:
        opening = float(data.get("opening_float") or 0)
        result = ops.start_shift(cashier, opening_float=opening)

    return jsonify(result)

@app.route('/api/shift/active')
def api_shift_active():
    return jsonify(ops.get_active_shift() or {})

@app.route('/api/shift/close', methods=['POST'])
def api_shift_close():
    data = request.json or {}
    till_count = data.get("till_count")

    if till_count is not None:
        result = ops.close_shift(till_count=till_count)
    else:
        result = ops.close_shift(float(data.get("actual_cash", 0)))

    return jsonify(result)

@app.route('/api/reports/cash_summary')
def api_cash_summary():
    active = ops.get_active_shift()
    if not active:
        return jsonify({"expected": 0, "cash_sales": 0, "net_cash": 0})
    # expected = opening_float + Σ(cash_given - change_given) for orders since shift start
    return jsonify(ops.get_cash_summary(active["ts_start"]))


# ---------- ADMIN API ----------
@app.route('/api/admin/add_item', methods=['POST'])
def add_item():
    data = request.json
    ops.add_item(data)
    return jsonify({'ok': True})

@app.route('/api/admin/update_item', methods=['POST'])
def update_item():
    data = request.json
    ops.update_item(data)
    return jsonify({'ok': True})

@app.route('/api/admin/delete_item', methods=['POST'])
def delete_item():
    data = request.json
    ops.delete_item(data['id'])
    return jsonify({'ok': True})

# ---------- ADMIN API: COMPONENTS ----------
@app.route('/api/admin/add_component', methods=['POST'])
def add_component():
    data = request.json
    ops.add_component(data)
    return jsonify({'ok': True})

@app.route('/api/admin/update_component', methods=['POST'])
def update_component():
    data = request.json
    ops.update_component(data)
    return jsonify({'ok': True})

@app.route('/api/admin/delete_component', methods=['POST'])
def delete_component():
    data = request.json
    ops.delete_component(data['id'])
    return jsonify({'ok': True})

@app.route('/api/admin/update_components_batch', methods=['POST'])
def update_components_batch():
    data = request.json or {}
    components = data.get("components", [])
    ops.batch_update_components(components)
    return jsonify({"ok": True})

# ---------- ADMIN API: RECIPES ----------
@app.route('/api/admin/add_recipe', methods=['POST'])
def add_recipe():
    data = request.json
    ops.add_recipe(data)
    return jsonify({'ok': True})

@app.route('/api/admin/get_recipes')
def get_recipes():
    rows = ops.list_recipes()
    return jsonify(rows)

@app.route('/api/admin/update_recipe', methods=['POST'])
def update_recipe():
    data = request.json
    ops.update_recipe(data)
    return jsonify({'ok': True})

@app.route('/api/admin/delete_recipe', methods=['POST'])
def delete_recipe():
    data = request.json
    ops.delete_recipe(data['id'])
    return jsonify({'ok': True})

@app.route('/api/admin/get_items')
def get_items():
    items = ops.list_items()
    return jsonify(items)

@app.route('/api/admin/get_components')
def get_components():
    comps = ops.list_components()
    return jsonify(comps)

@app.route('/api/admin/get_inventory')
def get_inventory():
    rows = ops.list_inventory()
    return jsonify(rows)

# ---------- ADMIN: Shift Reports ----------
@app.route('/admin/shifts')
def admin_shifts_page():
    return render_template('admin/shifts.html')

@app.route('/api/admin/get_shifts')
def api_admin_get_shifts():
    start = request.args.get('start')
    end = request.args.get('end')
    return jsonify(ops.list_shifts(start, end))

# ---------- INVENTORY ----------
@app.route('/api/admin/add_inventory', methods=['POST'])
def add_inventory():
    data = request.json
    ops.add_inventory_item(data)
    return jsonify({'ok': True})

@app.route('/api/admin/update_inventory', methods=['POST'])
def update_inventory():
    data = request.json
    ops.update_inventory_item(data)
    return jsonify({'ok': True})

@app.route('/api/admin/delete_inventory', methods=['POST'])
def delete_inventory():
    data = request.json
    ops.delete_inventory_item(data['id'])
    return jsonify({'ok': True})

@app.route('/api/inventory/delete', methods=['POST'])
def api_inventory_delete():
    data = request.json
    ops.delete_inventory_item(int(data['id']))
    return jsonify({'ok': True})

# ---------- MAIN API ----------
@app.route('/api/items')
def api_items():
    items = ops.list_items()
    for it in items:
        qty = ops.available_qty(it["id"])
        it["available_qty"] = qty
        it["in_stock"] = qty > 0
    return jsonify(items)

@app.route('/api/flavors')
def api_flavors():
    return jsonify(ops.list_flavors())

@app.route('/api/admin/flavors')
def api_admin_flavors():
    return jsonify(ops.list_flavors(include_inactive=True))

@app.route('/api/admin/flavors/add', methods=['POST'])
def api_admin_flavors_add():
    ops.add_flavor(request.json or {})
    return jsonify({"ok": True})

@app.route('/api/admin/flavors/update', methods=['POST'])
def api_admin_flavors_update():
    ops.update_flavor(request.json or {})
    return jsonify({"ok": True})

@app.route('/api/admin/flavors/delete', methods=['POST'])
def api_admin_flavors_delete():
    data = request.json or {}
    ops.delete_flavor(int(data["id"]))
    return jsonify({"ok": True})

import traceback

@app.route('/api/order', methods=['POST'])
def api_order():
    data = request.json or {}
    try:
        cashier = data.get('cashier', 'Unknown')
        method  = data.get('method', 'Cash')
        lines   = data.get('lines', [])

        discount = float(data.get("discount", 0) or 0)
        discount_type = data.get("discount_type")  # optional, UI-only
        discount_details = data.get("discount_details", [])

        cash_given   = float(data.get('cash_given', 0) or 0)
        change_given = float(data.get('change_given', 0) or 0)
        
        settings = ops.get_pos_settings()
        tax_rate = float(settings.get("tax_rate", 0) or 0)

        result = ops.record_order(
            cashier=cashier,
            method=method,
            discount=discount,
            discount_type=discount_type,
            discount_details=discount_details,
            lines=lines,
            tax_rate=tax_rate,
            cash_given=cash_given,
            change_given=change_given
        )

        # If you want the frontend to know what type was used:
        result["discount_type"] = discount_type

        return jsonify(result)

    except Exception as e:
        log("ERROR in /api/order:")
        log(str(e))
        log(traceback.format_exc())
        return jsonify({"error": str(e)}), 400


@app.route('/api/inventory')
def api_inventory():
    rows = ops.list_inventory()
    return jsonify(rows)

@app.route('/api/inventory/update', methods=['POST'])
def api_inventory_update():
    data = request.json
    log(f"Inventory update called: {data}")
    ops.update_inventory_item(data)
    return jsonify({'ok': True})


@app.route('/api/reports')
def api_reports():
    report = ops.sales_summary()
    return jsonify(report)

# ---------- ADMIN TOOLS ----------
from flask import send_file
import io, time

@app.route('/admin/tools')
def admin_tools():
    return render_template('admin/tools.html')

@app.route("/api/admin/export", methods=["POST"])
def api_admin_export():
    data = request.json or {}

    export_type = data.get("export_type", "catalog")
    fmt = data.get("format", "csv")

    try:
        saved_to = ops.export_database(
            export_type=export_type,
            fmt=fmt
        )

        return jsonify({
            "ok": True,
            "saved_to": saved_to
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e)
        }), 500

@app.route('/api/admin/open_exports_folder', methods=['POST'])
def api_open_exports():
    try:
        import subprocess, os
        folder = EXPORT_DIR
        if os.name == "nt":  # Windows
            subprocess.Popen(["explorer", str(folder)])
        else:
            # Linux/Mac fallback
            subprocess.Popen(["xdg-open", str(folder)])
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/api/admin/restore", methods=["POST"])
def api_admin_restore():
    restore_type = request.form.get("restore_type", "catalog")

    if "backup_file" not in request.files:
        return jsonify({"ok": False, "error": "No backup file uploaded."}), 400

    file = request.files["backup_file"]

    if not file.filename:
        return jsonify({"ok": False, "error": "No file selected."}), 400

    filename = secure_filename(file.filename)

    if not filename.lower().endswith(".sql"):
        return jsonify({
            "ok": False,
            "error": "Only .sql restore files are supported right now."
        }), 400

    upload_dir = ops._restore_uploads_dir()
    path = os.path.join(upload_dir, filename)

    file.save(path)

    try:
        result = ops.restore_database_from_sql(path, restore_type=restore_type)
        return jsonify(result)

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e)
        }), 500

#---------- REPORTS API ----------
from flask import send_file, make_response
import io

# --- Reports API: list orders for a date ---
@app.route('/api/reports/orders')
def api_reports_orders():
    date_str = request.args.get('date')  # expected YYYY-MM-DD
    if not date_str:
        return jsonify({"error": "Missing ?date=YYYY-MM-DD"}), 400
    return jsonify(ops.list_orders_by_date(date_str))

# --- Reports API: get single receipt (header + lines) ---
@app.route('/api/reports/receipt/<int:order_id>')
def api_reports_receipt(order_id):
    data = ops.get_receipt(order_id)
    if not data:
        return jsonify({"error": "Order not found"}), 404
    return jsonify(data)

# --- Reports API: down PDF for a receipt ---
# Uses ReportLab (pure-Python, no external binary needed)
@app.route('/api/reports/receipt/<int:order_id>/pdf')
def api_reports_receipt_pdf(order_id):
    data = ops.get_receipt(order_id)
    if not data:
        return jsonify({"error": "Order not found"}), 404

    # Create a simple PDF
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter

    # --- Centered title ---
    title = "The Gardiner Brew In"
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2.0, height - 0.6*inch, title)

    # --- Order header below title ---
    y = height - 1.0*inch
    c.setFont("Helvetica-Bold", 12)
    c.drawString(1*inch, y, f"Order #{data['order']['receipt_no']}")
    c.setFont("Helvetica", 10)

    # format ts -> 12-hour time
    from datetime import datetime
    try:
        dt = datetime.fromisoformat(data['order']['ts'].replace('Z','+00:00'))
    except Exception:
        dt = None
    time_str = data['order']['ts']
    if dt:
        time_str = dt.strftime("%Y-%m-%d %I:%M %p")

    y -= 0.22*inch
    c.drawString(1*inch, y, f"Date/Time: {time_str}")
    y -= 0.22*inch
    c.drawString(1*inch, y, f"Cashier: {data['order']['cashier']}     Method: {data['order']['method']}")
    y -= 0.3*inch

    # table headers
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1*inch, y, "Item")
    c.drawString(3.8*inch, y, "Qty")
    c.drawString(4.6*inch, y, "Price")
    c.drawString(5.4*inch, y, "Line")
    y -= 0.16*inch
    c.line(1*inch, y, 7.5*inch, y)
    y -= 0.18*inch


    c.setFont("Helvetica", 10)
    for ln in data["lines"]:
        # new page if needed
        if y < 1*inch:
            c.showPage()
            y = height - 0.75*inch
            c.setFont("Helvetica", 10)

        c.drawString(1*inch, y, ln["item_name"])
        c.drawRightString(4.3*inch, y, f"{ln['qty']:.0f}")
        c.drawRightString(5.2*inch, y, f"${float(ln['unit_price']):.2f}")
        c.drawRightString(6.2*inch, y, f"${float(ln['line_total']):.2f}")
        y -= 0.22*inch

        modifiers = ln.get("modifiers", []) or []
        if modifiers:
            c.setFont("Helvetica-Oblique", 8)
            for m in modifiers:
                if y < 1*inch:
                    c.showPage()
                    y = height - 0.75*inch
                    c.setFont("Helvetica-Oblique", 8)

                name = m.get("name") or "Modifier"
                qty = int(m.get("qty") or 0)
                unit_price = float(m.get("unit_price") or 0)
                line_total = float(m.get("line_total") or 0)

                c.drawString(1.18*inch, y, f"+ {name} x{qty}")
                c.drawRightString(5.2*inch, y, f"${unit_price:.2f}")
                c.drawRightString(6.2*inch, y, f"${line_total:.2f}")
                y -= 0.16*inch

            c.setFont("Helvetica", 10)

    y -= 0.1*inch
    c.line(4.8*inch, y, 6.2*inch, y)
    y -= 0.22*inch

    o = data["order"]
    def money(x): return f"${float(x or 0):.2f}"

    c.drawRightString(5.2*inch, y, "Subtotal:")
    c.drawRightString(6.2*inch, y, money(o["subtotal"]))
    y -= 0.18*inch

    discounts = data.get("discounts", [])

    if discounts:
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(5.2*inch, y, "Discounts:")
        c.drawRightString(6.2*inch, y, money(o["discount"]))
        y -= 0.20*inch

        c.setFont("Helvetica", 9)

        for d in discounts:
            label = d.get("sale_name") or d.get("sale_type") or "Discount"
            source = d.get("source") or ""

            c.drawRightString(
                6.2*inch,
                y,
                f"{source}: {label} -{money(d.get('amount'))}"
            )
            y -= 0.16*inch

            if d.get("source") == "MANUAL" and d.get("authorized_pin_label"):
                c.setFont("Helvetica-Oblique", 8)
                c.drawRightString(
                    6.2*inch,
                    y,
                    f"Authorized by: {d.get('authorized_pin_label')}"
                )
                c.setFont("Helvetica", 9)
                y -= 0.14*inch

        c.setFont("Helvetica", 10)

    else:
        c.drawRightString(5.2*inch, y, "Discount:")
        c.drawRightString(6.2*inch, y, money(o["discount"]))
        y -= 0.18*inch

        if o.get("discount_type"):
            c.setFont("Helvetica-Oblique", 9)
            c.drawRightString(6.2*inch, y, f"({o['discount_type']})")
            c.setFont("Helvetica", 10)
            y -= 0.18*inch


    if float(o["tax"] or 0) > 0:
        c.drawRightString(5.2*inch, y, "Tax:")
        c.drawRightString(6.2*inch, y, money(o["tax"]))
        y -= 0.18*inch

    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(5.2*inch, y, "Total:")
    c.drawRightString(6.2*inch, y, money(o["total"]))

    c.showPage()
    c.save()
    pdf = buf.getvalue()
    buf.close()

    resp = make_response(pdf)
    resp.headers['Content-Type'] = 'application/pdf'
    resp.headers['Content-Disposition'] = f'attachment; filename=receipt_{order_id}.pdf'
    return resp

# --- Reports API: delete an order (and restore inventory) ---
@app.route('/api/reports/order/<int:order_id>/delete', methods=['POST'])
def api_delete_order(order_id):
    ok = ops.delete_order_and_restore(order_id)
    return jsonify({'ok': bool(ok)})

# --- Reports API: daily sales summary + export to XLSX ---
@app.route('/api/reports/sales')
def api_reports_sales():
    date_str = request.args.get('date')
    if not date_str:
        return jsonify({'error':'missing ?date=YYYY-MM-DD'}),400
    return jsonify(ops.daily_sales_summary(date_str))

@app.route('/api/reports/sales/xlsx')
def api_reports_sales_xlsx():
    date_str = request.args.get('date')
    if not date_str:
        return jsonify({'error':'missing ?date=YYYY-MM-DD'}),400
    xlsx_bytes = ops.export_sales_to_xlsx(date_str)
    fname = f"sales_{date_str}.xlsx"
    return send_file(io.BytesIO(xlsx_bytes),
                     mimetype="application/octet-stream",
                     as_attachment=True,
                     download_name=fname)

# --- Reports API: sales summary for a date range + export to XLSX ---
from flask import send_file
import io

@app.route('/api/reports/sales_range')
def api_reports_sales_range():
    start = request.args.get('start')
    end   = request.args.get('end')
    if not start or not end:
        return jsonify({'error':'missing ?start=YYYY-MM-DD&end=YYYY-MM-DD'}), 400
    return jsonify(ops.sales_summary_range(start, end))

@app.route('/api/reports/sales_range/xlsx')
def api_reports_sales_range_xlsx():
    start = request.args.get('start')
    end   = request.args.get('end')
    if not start or not end:
        return jsonify({'error':'missing ?start=YYYY-MM-DD&end=YYYY-MM-DD'}), 400
    payload = ops.export_sales_range_to_xlsx(start, end)
    fname = f"sales_{start}_to_{end}.xlsx"
    return send_file(io.BytesIO(payload),
                     mimetype="application/octet-stream",
                     as_attachment=True,
                     download_name=fname)

# --- Admin API: get and update POS settings ---

@app.route('/api/admin/settings')
def api_admin_settings():
    return jsonify(ops.get_pos_settings())


@app.route('/api/admin/settings/update', methods=['POST'])
def api_admin_settings_update():
    data = request.json or {}

    for key in [
        "require_pin_discounts",
        "require_pin_admin_access",
        "require_pin_exit_program"
    ]:
        ops.set_setting(key, "1" if data.get(key) else "0")

    tax_rate = float(data.get("tax_rate") or 0)

    if tax_rate < 0:
        tax_rate = 0

    ops.set_setting("tax_rate", str(tax_rate))

    return jsonify({"ok": True})


@app.route('/api/admin/admin_pins')
def api_admin_admin_pins():
    return jsonify(ops.list_admin_pins())


@app.route('/api/admin/admin_pins/add', methods=['POST'])
def api_admin_admin_pin_add():
    data = request.json or {}
    ops.add_admin_pin(data.get("pin", ""), data.get("label", ""))
    return jsonify({"ok": True})


@app.route('/api/admin/admin_pins/delete', methods=['POST'])
def api_admin_admin_pin_delete():
    data = request.json or {}
    ops.delete_admin_pin(int(data["id"]))
    return jsonify({"ok": True})


@app.route('/api/discount/validate_pin', methods=['POST'])
@app.route('/api/security/validate_pin', methods=['POST'])
def api_validate_admin_pin():
    data = request.json or {}
    pin_info = ops.validate_admin_pin_info(data.get("pin", ""))

    if not pin_info:
        return jsonify({"ok": False})

    return jsonify({
        "ok": True,
        "pin_id": pin_info["id"],
        "pin_label": pin_info["label"] or f"PIN #{pin_info['id']}"
    })

# ---------- SALES / PROMOTIONS API ----------

@app.route('/api/admin/sales')
def api_admin_sales():
    return jsonify(ops.list_sales(include_inactive=True))


@app.route('/api/admin/sales/add', methods=['POST'])
def api_admin_sales_add():
    data = request.json or {}
    ops.add_sale(data)
    return jsonify({"ok": True})


@app.route('/api/admin/sales/update', methods=['POST'])
def api_admin_sales_update():
    data = request.json or {}
    ops.update_sale(data)
    return jsonify({"ok": True})


@app.route('/api/admin/sales/delete', methods=['POST'])
def api_admin_sales_delete():
    data = request.json or {}
    ops.delete_sale(int(data["id"]))
    return jsonify({"ok": True})


@app.route('/api/sales')
def api_sales():
    return jsonify(ops.list_active_sales())

# ---------- ADMIN PAGES ----------

@app.before_request
def require_admin_pin_for_admin_pages():
    """
    Protect /admin pages when require_pin_admin_access is enabled.
    Uses session['admin_unlocked'] after successful PIN entry.
    """

    path = request.path

    # Only protect actual admin pages
    if not path.startswith("/admin"):
        return None

    # Let the login page itself load
    if path == "/admin/login":
        return None

    # Check setting
    settings = ops.get_pos_settings()
    if not settings.get("require_pin_admin_access"):
        return None

    # Avoid accidental lockout if no PIN has been created yet
    if not ops.has_admin_pins():
        return None

    # Already unlocked in this browser session
    if session.get("admin_unlocked") == True:
        return None

    # Redirect to login and preserve where they were trying to go
    return redirect(url_for("admin_login", next=path))

# Admin login page to enter PIN

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    next_url = request.args.get("next") or "/admin/tools"

    if request.method == "GET":
        return render_template("admin/login.html", next_url=next_url)

    data = request.json or request.form
    pin = data.get("pin", "")

    if ops.validate_admin_pin(pin):
        session["admin_unlocked"] = True
        return jsonify({"ok": True, "next": next_url})

    return jsonify({"ok": False, "error": "Incorrect PIN"}), 401


@app.route('/admin/logout')
def admin_logout():
    session.pop("admin_unlocked", None)
    return redirect("/pos")

# ---------- MAIN ----------
if __name__ == '__main__':
    host = '127.0.0.1'
    port = 5000
    if '--host=0.0.0.0' in sys.argv:
        host = '0.0.0.0'
    app.run(host=host, port=port, debug=True)