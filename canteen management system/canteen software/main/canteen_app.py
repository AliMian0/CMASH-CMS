"""
CMASH Canteen Management System
--------------------------------
A single-file Streamlit web app for running a canteen:
- Login / first-time admin setup (username + password, stored locally)
- Point of Sale (scan/type a barcode OR pick a product manually)
- Inventory management (add / update / remove items, barcode support)
- Employee ledger (monthly tab accounts) with custom date-range viewing
- Profit & sales reports with quick periods AND custom date ranges
- Excel (.xlsx) export for inventory and payroll reports
- Thermal receipt printing (58mm / 80mm) straight from the browser

Run with:  streamlit run canteen_app.py
"""
import hashlib
import io
import json  # <--- ADD THIS IMPORT
import secrets
import sqlite3
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import hashlib
import io
import secrets
import sqlite3
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

DB_PATH = "canteen.db"

# Payment types that must be tied to a specific employee (deducted from their account)
EMPLOYEE_LINKED_PAYMENT_TYPES = [
    "Employee Tab (Monthly Account)",
    
]

# =====================================================================
# DATABASE LAYER
# =====================================================================

def get_conn() -> sqlite3.Connection:
    """Open a new connection with sane defaults (safe for Streamlit reruns)."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    conn = get_conn()
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    salt TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    barcode TEXT UNIQUE,
                    name TEXT UNIQUE NOT NULL,
                    cost_price REAL NOT NULL,
                    retail_price REAL NOT NULL,
                    stock_qty INTEGER NOT NULL DEFAULT 0
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    emp_code TEXT UNIQUE NOT NULL,
                    name TEXT NOT NULL,
                    dept TEXT
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS sales (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    payment_type TEXT NOT NULL,
                    emp_id INTEGER,
                    total_amount REAL NOT NULL,
                    sale_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(emp_id) REFERENCES employees(id)
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS sale_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sale_id INTEGER NOT NULL,
                    product_id INTEGER NOT NULL,
                    qty INTEGER NOT NULL,
                    cost_price REAL NOT NULL,
                    retail_price REAL NOT NULL,
                    FOREIGN KEY(sale_id) REFERENCES sales(id),
                    FOREIGN KEY(product_id) REFERENCES products(id)
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )''')

    conn.commit()
    conn.close()
    migrate_db()



def reset_all_data() -> None:
    """Permanently delete all application data from the SQLite database."""
    conn = get_conn()
    try:
        # Delete child records first because foreign-key enforcement is enabled.
        conn.execute("DELETE FROM sale_items")
        conn.execute("DELETE FROM sales")
        conn.execute("DELETE FROM products")
        conn.execute("DELETE FROM employees")
        conn.execute("DELETE FROM users")
        conn.execute("DELETE FROM settings")

        # Reset AUTOINCREMENT counters so the database starts cleanly.
        conn.execute("DELETE FROM sqlite_sequence")
        conn.commit()
    finally:
        conn.close()

# ---------------- Settings (shop name / receipt config) ----------------

DEFAULT_SETTINGS = {
    "shop_name": "CMASH Canteen",
    "shop_address": "",
    "shop_phone": "",
    "receipt_footer": "Thank you! Please come again.",
    "paper_width_mm": "80",
}


def get_setting(key: str) -> str:
    conn = get_conn()
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    conn.close()
    return row[0] if row else DEFAULT_SETTINGS.get(key, "")


def get_all_settings() -> dict:
    return {k: get_setting(k) for k in DEFAULT_SETTINGS}


def set_setting(key: str, value: str) -> None:
    conn = get_conn()
    conn.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )
    conn.commit()
    conn.close()


def migrate_db() -> None:
    """Adds columns that didn't exist in earlier versions of this app,
    without touching or deleting any existing data."""
    conn = get_conn()
    c = conn.cursor()

    user_cols = {row[1] for row in c.execute("PRAGMA table_info(users)").fetchall()}
    if "role" not in user_cols:
        c.execute("ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT 'admin'")
        conn.commit()
    user_cols = {row[1] for row in c.execute("PRAGMA table_info(users)").fetchall()}
    if "permissions" not in user_cols:
        c.execute("ALTER TABLE users ADD COLUMN permissions TEXT NOT NULL DEFAULT '{}'")
        conn.commit()

    existing_cols = {row[1] for row in c.execute("PRAGMA table_info(products)").fetchall()}
    if "barcode" not in existing_cols:
        c.execute("ALTER TABLE products ADD COLUMN barcode TEXT")
        conn.commit()
        # SQLite can't add a UNIQUE column via ALTER TABLE, so enforce
        # uniqueness with an index instead (NULLs are still allowed to repeat).
        try:
            c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_products_barcode ON products(barcode)")
            conn.commit()
        except sqlite3.IntegrityError:
            # Existing duplicate barcodes (e.g. blank strings) - skip enforcing
            # uniqueness rather than crashing; new inserts will still be checked
            # once duplicates are cleaned up.
            pass

    conn.close()


# =====================================================================
# ACCESS CONTROL
# =====================================================================

ACCESS_OPTIONS = {
    "pos": "POS / Sales",
    "returns": "Returns",
    "inventory": "Inventory",
    "products": "Add / Edit Products",
    "prices": "Change Prices",
    "stock": "Change Stock",
    "employees": "Employee Ledger",
    "hospital": "Hospital Expense",
    "ot": "OT Expense",
    "reports": "Profit & Sales Reports",
    "excel": "Excel Reports (View / Download)",
}

# These are intentionally NOT permissions:
# - Edit paid bill: never allowed
# - Edit Excel data: never allowed
# - Settings: Admin only
# - Reset All Data: Admin only

DEFAULT_SHOP_PERMISSIONS = list(ACCESS_OPTIONS.keys())


def normalize_permissions(value) -> list[str]:
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except Exception:
            value = []
    if not isinstance(value, list):
        return []
    return [p for p in value if p in ACCESS_OPTIONS]


def user_has_permission(permission: str) -> bool:
    if st.session_state.get("role", "admin") == "admin":
        return True
    return permission in normalize_permissions(st.session_state.get("permissions", []))


def get_user_permissions(username: str) -> list[str]:
    conn = get_conn()
    row = conn.execute(
        "SELECT role, COALESCE(permissions, '{}') FROM users WHERE username = ?",
        (username.strip(),)
    ).fetchone()
    conn.close()
    if not row:
        return []
    role, permissions = row
    return list(ACCESS_OPTIONS.keys()) if role == "admin" else normalize_permissions(permissions)


def save_user_permissions(username: str, permissions: list[str]) -> None:
    conn = get_conn()
    conn.execute(
        "UPDATE users SET permissions = ? WHERE username = ?",
        (json.dumps(normalize_permissions(permissions)), username.strip()),
    )
    conn.commit()
    conn.close()


# =====================================================================
# AUTH HELPERS
# =====================================================================

def hash_password(password: str, salt: str | None = None) -> tuple[str, str]:
    if salt is None:
        salt = secrets.token_hex(16)
    pwd_hash = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 100_000).hex()
    return pwd_hash, salt


def any_user_exists() -> bool:
    conn = get_conn()
    n = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    conn.close()
    return n > 0


def create_user(
    username: str,
    password: str,
    role: str = "admin",
    permissions: list[str] | None = None,
) -> tuple[bool, str]:
    pwd_hash, salt = hash_password(password)
    if permissions is None:
        permissions = list(ACCESS_OPTIONS.keys()) if role == "admin" else []
    permissions_json = json.dumps(normalize_permissions(permissions))
    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO users (username, password_hash, salt, role, permissions) VALUES (?, ?, ?, ?, ?)",
            (username.strip(), pwd_hash, salt, role, permissions_json),
        )
        conn.commit()
        return True, "Account created."
    except sqlite3.IntegrityError:
        return False, "That username already exists."
    finally:
        conn.close()


def verify_login(username: str, password: str) -> tuple[bool, str, list[str]]:
    conn = get_conn()
    row = conn.execute(
        "SELECT password_hash, salt, COALESCE(role, 'admin'), COALESCE(permissions, '{}') "
        "FROM users WHERE username = ?",
        (username.strip(),)
    ).fetchone()
    conn.close()
    if not row:
        return False, "", []
    stored_hash, salt, role, permissions = row
    candidate_hash, _ = hash_password(password, salt)
    ok = secrets.compare_digest(candidate_hash, stored_hash)
    perms = list(ACCESS_OPTIONS.keys()) if role == "admin" else normalize_permissions(permissions)
    return ok, role, perms


def login_gate() -> bool:
    """Renders login / first-run setup screens. Returns True once authenticated."""
    if st.session_state.get("logged_in"):
        return True

    st.title("🍽️ Canteen Management System")

    if not any_user_exists():
        st.subheader("First-time setup: create your admin account")
        st.caption("No account exists yet. Choose a username and password to protect this system.")
        with st.form("setup_form"):
            new_user = st.text_input("Choose a username")
            new_pass = st.text_input("Choose a password", type="password")
            confirm_pass = st.text_input("Confirm password", type="password")
            submitted = st.form_submit_button("Create Account & Continue")
        if submitted:
            if not new_user or not new_pass:
                st.error("Username and password cannot be empty.")
            elif new_pass != confirm_pass:
                st.error("Passwords do not match.")
            elif len(new_pass) < 4:
                st.error("Password should be at least 4 characters.")
            else:
                ok, msg = create_user(new_user, new_pass)
                if ok:
                    st.success("Account created! Please log in below.")
                    st.rerun()
                else:
                    st.error(msg)
        return False

    st.subheader("Login")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log In")
    if submitted:
        ok, role, permissions = verify_login(username, password)
        if ok:
            st.session_state["logged_in"] = True
            st.session_state["username"] = username.strip()
            st.session_state["role"] = role
            st.session_state["permissions"] = permissions
            st.rerun()
        else:
            st.error("Invalid username or password.")

    with st.expander("Add another admin account"):
        with st.form("add_admin_form"):
            au = st.text_input("New username", key="add_admin_user")
            ap = st.text_input("New password", type="password", key="add_admin_pass")
            add_submit = st.form_submit_button("Create Additional Account")
        if add_submit:
            if not au or not ap:
                st.error("Username and password cannot be empty.")
            else:
                ok, msg = create_user(au, ap)
                st.success(msg) if ok else st.error(msg)

    return False


# =====================================================================
# GENERAL HELPERS
# =====================================================================

def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> io.BytesIO:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output


def date_range_picker(key_prefix: str, default_days_back: int = 30) -> tuple[date, date]:
    """Shared quick-period + custom-range picker. Returns (start_date, end_date)."""
    quick = st.selectbox(
        "Report Period",
        ["Today", "This Week", "This Month", "This Year", "Custom Range"],
        key=f"{key_prefix}_quick",
    )
    today = date.today()
    if quick == "Today":
        start, end = today, today
    elif quick == "This Week":
        start, end = today - timedelta(days=7), today
    elif quick == "This Month":
        start, end = today.replace(day=1), today
    elif quick == "This Year":
        start, end = today.replace(month=1, day=1), today
    else:
        c1, c2 = st.columns(2)
        with c1:
            start = st.date_input(
                "Start Date", value=today - timedelta(days=default_days_back), key=f"{key_prefix}_start"
            )
        with c2:
            end = st.date_input("End Date", value=today, key=f"{key_prefix}_end")

    if start > end:
        st.error("Start date must be before end date.")
        st.stop()

    return start, end


def sql_bounds(start: date, end: date) -> tuple[str, str]:
    return f"{start} 00:00:00", f"{end} 23:59:59"


# =====================================================================
# THERMAL RECEIPT PRINTING
# =====================================================================

def render_receipt(sale: dict) -> None:
    """Renders a printable thermal-receipt widget (58mm or 80mm) with a
    Print button. Works with any thermal printer set up as a normal
    Windows/Mac printer - clicking Print opens the browser's print dialog
    scoped to just this receipt."""

    settings = get_all_settings()
    paper_mm = settings.get("paper_width_mm", "80") or "80"

    items_rows = ""
    for it in sale["items"]:
        items_rows += (
            f"<tr>"
            f"<td style='text-align:left'>{it['name']}</td>"
            f"<td style='text-align:center'>{it['qty']}</td>"
            f"<td style='text-align:right'>{it['price']:,.2f}</td>"
            f"<td style='text-align:right'>{it['line_total']:,.2f}</td>"
            f"</tr>"
        )

    emp_line = f"<p>Employee: {sale['emp_label']}</p>" if sale.get("emp_label") else ""
    address_line = f"<p>{settings['shop_address']}</p>" if settings.get("shop_address") else ""
    phone_line = f"<p>{settings['shop_phone']}</p>" if settings.get("shop_phone") else ""

    receipt_body = f"""
    <div id="receipt-wrapper" style="display:flex; flex-direction:column; align-items:center; font-family:Arial, sans-serif;">
      <div id="receipt" style="width:{paper_mm}mm; min-width:220px; background:#fff; color:#000;
           font-family:'Courier New', monospace; font-size:12px; padding:10px;
           border:1px solid #ddd; box-shadow:0 1px 4px rgba(0,0,0,0.15);">
        <h3 style="text-align:center; margin:2px 0; font-size:15px;">{settings['shop_name']}</h3>
        <div style="text-align:center; font-size:11px;">
          {address_line}
          {phone_line}
        </div>
        <hr style="border:none; border-top:1px dashed #000; margin:6px 0;">
        <div style="font-size:11px;">
          <div>Sale #: {sale['sale_id']}</div>
          <div>Date: {sale['sale_date']}</div>
          <div>Payment: {sale['payment_type']}</div>
          {emp_line}
        </div>
        <hr style="border:none; border-top:1px dashed #000; margin:6px 0;">
        <table style="width:100%; border-collapse:collapse; font-size:11px;">
          <thead>
            <tr>
              <th style="text-align:left;">Item</th>
              <th style="text-align:center;">Qty</th>
              <th style="text-align:right;">Price</th>
              <th style="text-align:right;">Total</th>
            </tr>
          </thead>
          <tbody>
            {items_rows}
          </tbody>
        </table>
        <hr style="border:none; border-top:1px dashed #000; margin:6px 0;">
        <div style="display:flex; justify-content:space-between; font-weight:bold; font-size:13px;">
          <span>TOTAL</span><span>PKR {sale['total']:,.2f}</span>
        </div>
        <hr style="border:none; border-top:1px dashed #000; margin:6px 0;">
        <p style="text-align:center; font-size:11px;">{settings['receipt_footer']}</p>
      </div>
      <button onclick="window.print()"
              style="margin-top:12px; padding:8px 20px; font-size:14px; cursor:pointer;
                     background:#0f9d58; color:#fff; border:none; border-radius:6px;">
        🖨️ Print Receipt
      </button>
    </div>
    <style>
      @media print {{
        button {{ display: none !important; }}
        body {{ margin: 0; }}
      }}
    </style>
    """

    components.html(receipt_body, height=520, scrolling=True)


# =====================================================================
# MODULE: POINT OF SALE
# =====================================================================

def module_pos(conn: sqlite3.Connection) -> None:
    st.header("🛒 Point of Sale & Checkout")
    st.info("Paid bills cannot be edited. Use **Return Paid Sale** to process a return without changing the original bill.")

    products = pd.read_sql("SELECT * FROM products WHERE stock_qty > 0", conn)
    employees = pd.read_sql("SELECT * FROM employees", conn)

    if products.empty:
        st.warning("No items available in inventory. Please add products first.")
        return

    if "pos_selected_product_id" not in st.session_state:
        st.session_state["pos_selected_product_id"] = None

    find_mode = st.radio(
        "Find Product By", ["Scan / Enter Barcode", "Select From List"], horizontal=True
    )

    item_row = None

    if find_mode == "Scan / Enter Barcode":
        with st.form("barcode_form", clear_on_submit=True):
            barcode_val = st.text_input(
                "Scan barcode or type it and press Enter", key="barcode_input"
            )
            find_submitted = st.form_submit_button("Find Item")
        if find_submitted and barcode_val.strip():
            match = products[products["barcode"] == barcode_val.strip()]
            if match.empty:
                st.error(f"No product found with barcode '{barcode_val.strip()}'.")
                st.session_state["pos_selected_product_id"] = None
            else:
                st.session_state["pos_selected_product_id"] = int(match.iloc[0]["id"])

        if st.session_state["pos_selected_product_id"] is not None:
            match = products[products["id"] == st.session_state["pos_selected_product_id"]]
            if not match.empty:
                item_row = match.iloc[0]
                st.success(f"Found: {item_row['name']}")
    else:
        selected_item = st.selectbox("Select Product", products["name"].tolist())
        item_row = products[products["name"] == selected_item].iloc[0]
        st.session_state["pos_selected_product_id"] = int(item_row["id"])

    if item_row is None:
        st.info("Scan a barcode or select a product to begin.")
        return

    col1, col2 = st.columns([2, 1])
    with col1:
        st.info(
            f"**{item_row['name']}**  |  Available Stock: {int(item_row['stock_qty'])}  |  "
            f"Retail Price: PKR {item_row['retail_price']:,.2f}"
        )
        qty = st.number_input(
            "Quantity", min_value=1, max_value=int(item_row["stock_qty"]), value=1, key="pos_qty"
        )

        pay_type = st.radio(
            "Payment Type",
            [
                "Cash / Direct Payment",
                "Employee Tab (Monthly Account)",
                "Hospital Expense",
                "OT Expense",
            ],
            key="pos_pay",
        )

        selected_emp_id = None
        if pay_type in EMPLOYEE_LINKED_PAYMENT_TYPES:
            if employees.empty:
                st.error("No employees registered. Add employees first under Employee Ledger.")
            else:
                emp_list = [f"{row['emp_code']} - {row['name']}" for _, row in employees.iterrows()]
                selected_emp_str = st.selectbox("Select Employee", emp_list, key="pos_emp")
                selected_emp_id = int(
                    employees[employees["emp_code"] == selected_emp_str.split(" - ")[0]]["id"].values[0]
                )

        total = qty * item_row["retail_price"]
        st.metric("Bill Total", f"PKR {total:,.2f}")

        can_checkout = not (pay_type in EMPLOYEE_LINKED_PAYMENT_TYPES and employees.empty)
        if st.button("✅ Complete Transaction", disabled=not can_checkout, type="primary"):
            c = conn.cursor()
            c.execute(
                "INSERT INTO sales (payment_type, emp_id, total_amount) VALUES (?, ?, ?)",
                (pay_type, selected_emp_id, total),
            )
            sale_id = c.lastrowid
            c.execute(
                """INSERT INTO sale_items (sale_id, product_id, qty, cost_price, retail_price)
                   VALUES (?, ?, ?, ?, ?)""",
                (sale_id, int(item_row["id"]), int(qty), item_row["cost_price"], item_row["retail_price"]),
            )
            c.execute(
                "UPDATE products SET stock_qty = stock_qty - ? WHERE id = ?",
                (int(qty), int(item_row["id"])),
            )
            conn.commit()

            emp_label = None
            if selected_emp_id is not None:
                emp_row = employees[employees["id"] == selected_emp_id].iloc[0]
                emp_label = f"{emp_row['emp_code']} - {emp_row['name']}"

            st.session_state["last_sale"] = {
                "sale_id": sale_id,
                "sale_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "payment_type": pay_type,
                "emp_label": emp_label,
                "total": total,
                "items": [
                    {
                        "name": item_row["name"],
                        "qty": int(qty),
                        "price": float(item_row["retail_price"]),
                        "line_total": float(total),
                    }
                ],
            }
            st.session_state["pos_selected_product_id"] = None
            st.success(f"Sale recorded! Total Bill: PKR {total:,.2f}")
            st.rerun()

    with col2:
        if st.session_state.get("last_sale"):
            st.subheader("🧾 Receipt")
            render_receipt(st.session_state["last_sale"])
            if st.button("Dismiss Receipt"):
                st.session_state["last_sale"] = None
                st.rerun()


# =====================================================================
# MODULE: INVENTORY MANAGEMENT
# =====================================================================

def module_inventory(conn: sqlite3.Connection) -> None:
    st.header("📦 Inventory & Product Operations")

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Current Stock", "Add New Item", "✏️ Edit / Correct Item", "📊 Reports & Excel Export"]
    )

    # ---------------- Current stock ----------------
    with tab1:
        df = pd.read_sql(
            """SELECT id AS 'Item ID', barcode AS 'Barcode', name AS 'Item Name',
                      cost_price AS 'Cost Price (PKR)', retail_price AS 'Retail Price (PKR)',
                      stock_qty AS 'Stock Quantity'
               FROM products ORDER BY name""",
            conn,
        )
        st.dataframe(df, use_container_width=True)

    # ---------------- Add new item ----------------
    with tab2:
        st.subheader("Add New Product")
        st.caption(
            "Use this to add a brand-new item, or to top up stock of an existing item "
            "(the quantity you enter here is ADDED to what's already in stock)."
        )
        with st.form("add_product_form", clear_on_submit=True):
            p_barcode = st.text_input("Barcode (optional - scan or type)")
            p_name = st.text_input("Product Name")
            c1, c2, c3 = st.columns(3)
            with c1:
                c_price = st.number_input("Cost Price (PKR)", min_value=0.0, step=1.0)
            with c2:
                r_price = st.number_input("Retail Price (PKR)", min_value=0.0, step=1.0)
            with c3:
                stock = st.number_input("Stock Quantity to Add", min_value=0, step=1)
            save_submitted = st.form_submit_button("💾 Save Product")

        if save_submitted:
            if not p_name.strip():
                st.error("Please provide a product name.")
            else:
                barcode_clean = p_barcode.strip() or None
                c = conn.cursor()
                try:
                    c.execute(
                        """INSERT INTO products (barcode, name, cost_price, retail_price, stock_qty)
                           VALUES (?, ?, ?, ?, ?)
                           ON CONFLICT(name) DO UPDATE SET
                             barcode = COALESCE(excluded.barcode, products.barcode),
                             cost_price = excluded.cost_price,
                             retail_price = excluded.retail_price,
                             stock_qty = products.stock_qty + excluded.stock_qty""",
                        (barcode_clean, p_name.strip(), c_price, r_price, stock),
                    )
                    conn.commit()
                    st.success(f"Saved product '{p_name.strip()}' successfully!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("That barcode is already assigned to another product.")

    # ---------------- Edit / correct existing item ----------------
    with tab3:
        st.subheader("Edit or Correct an Existing Product")
        st.caption(
            "Use this to fix a mistake — the values you enter here REPLACE the "
            "current price and stock quantity (they are not added on top)."
        )
        p_df = pd.read_sql("SELECT * FROM products ORDER BY name", conn)

        if p_df.empty:
            st.info("No products yet. Add one in the 'Add New Item' tab first.")
        else:
            edit_name = st.selectbox("Select Product to Edit", p_df["name"].tolist(), key="edit_product")
            row = p_df[p_df["name"] == edit_name].iloc[0]

            with st.form("edit_product_form"):
                e_barcode = st.text_input("Barcode", value=row["barcode"] or "")
                e_name = st.text_input("Product Name", value=row["name"])
                ec1, ec2, ec3 = st.columns(3)
                with ec1:
                    e_cost = st.number_input(
                        "Cost Price (PKR)", min_value=0.0, step=1.0, value=float(row["cost_price"])
                    )
                with ec2:
                    e_retail = st.number_input(
                        "Retail Price (PKR)", min_value=0.0, step=1.0, value=float(row["retail_price"])
                    )
                with ec3:
                    e_stock = st.number_input(
                        "Correct Stock Quantity", min_value=0, step=1, value=int(row["stock_qty"])
                    )
                edit_submitted = st.form_submit_button("✅ Save Correction", type="primary")

            if edit_submitted:
                if not e_name.strip():
                    st.error("Product name cannot be empty.")
                else:
                    barcode_clean = e_barcode.strip() or None
                    try:
                        conn.execute(
                            """UPDATE products
                               SET barcode = ?, name = ?, cost_price = ?, retail_price = ?, stock_qty = ?
                               WHERE id = ?""",
                            (barcode_clean, e_name.strip(), e_cost, e_retail, e_stock, int(row["id"])),
                        )
                        conn.commit()
                        st.success(f"'{e_name.strip()}' updated successfully.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("That name or barcode is already used by another product.")

            st.divider()
            st.subheader("Remove Product")
            del_item = st.selectbox("Select Product to Delete", p_df["name"].tolist(), key="del_product")
            if st.button("🗑️ Delete Item", type="primary"):
                conn.execute("DELETE FROM products WHERE name = ?", (del_item,))
                conn.commit()
                st.warning(f"Deleted {del_item}.")
                st.rerun()

    # ---------------- Reports & export ----------------
    with tab4:
        st.subheader("Inventory Movement Report")
        start_dt, end_dt = date_range_picker("inv")
        start_str, end_str = sql_bounds(start_dt, end_dt)

        current_inv_df = pd.read_sql(
            """SELECT name AS 'Item Name', barcode AS 'Barcode',
                      cost_price AS 'Unit Cost (PKR)', retail_price AS 'Unit Retail (PKR)',
                      stock_qty AS 'Current Stock Qty',
                      (cost_price * stock_qty) AS 'Total Stock Cost (PKR)',
                      (retail_price * stock_qty) AS 'Potential Retail Value (PKR)'
               FROM products ORDER BY name""",
            conn,
        )

        movement_df = pd.read_sql(
            """SELECT p.name AS 'Item Name',
                      SUM(si.qty) AS 'Units Sold In Period',
                      SUM(si.qty * si.cost_price) AS 'Total Cost Value (PKR)',
                      SUM(si.qty * si.retail_price) AS 'Total Revenue Value (PKR)',
                      SUM(si.qty * (si.retail_price - si.cost_price)) AS 'Profit (PKR)'
               FROM sale_items si
               JOIN sales s ON si.sale_id = s.id
               JOIN products p ON si.product_id = p.id
               WHERE s.sale_date BETWEEN ? AND ?
               GROUP BY p.id
               ORDER BY p.name""",
            conn,
            params=(start_str, end_str),
        )

        st.write("### Current Inventory Snapshot")
        st.dataframe(current_inv_df, use_container_width=True)

        st.write(f"### Item Movement ({start_dt} to {end_dt})")
        if movement_df.empty:
            st.info("No sales in this period.")
        else:
            st.dataframe(movement_df, use_container_width=True)
            m1, m2, m3 = st.columns(3)
            m1.metric("Units Sold", int(movement_df["Units Sold In Period"].sum()))
            m2.metric("Revenue", f"PKR {movement_df['Total Revenue Value (PKR)'].sum():,.2f}")
            m3.metric("Profit", f"PKR {movement_df['Profit (PKR)'].sum():,.2f}")

        excel_inv = to_excel_bytes(
            {"Current Inventory": current_inv_df, "Item Movement Report": movement_df}
        )
        st.download_button(
            "📥 Download Inventory Report (.xlsx)",
            data=excel_inv,
            file_name=f"Inventory_Report_{start_dt}_to_{end_dt}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# =====================================================================
# MODULE: EMPLOYEE LEDGER
# =====================================================================

def module_employee_ledger(conn: sqlite3.Connection) -> None:
    st.header("👥 Employee Ledger & Consumption Recovery")
    st.caption("Tracks canteen items bought on the monthly employee tab (Employee Tab payment type).")

    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("Register Employee")
        with st.form("add_employee_form", clear_on_submit=True):
            e_code = st.text_input("Employee ID / Code")
            e_name = st.text_input("Full Name")
            e_dept = st.text_input("Department")
            add_emp_submitted = st.form_submit_button("➕ Add Employee")
        if add_emp_submitted:
            if e_code.strip() and e_name.strip():
                try:
                    conn.execute(
                        "INSERT INTO employees (emp_code, name, dept) VALUES (?, ?, ?)",
                        (e_code.strip(), e_name.strip(), e_dept.strip()),
                    )
                    conn.commit()
                    st.success("Employee registered!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("That employee code already exists.")
            else:
                st.error("Employee code and name are required.")

        st.divider()
        st.subheader("Remove Employee")
        emp_df_all = pd.read_sql("SELECT * FROM employees ORDER BY name", conn)
        if not emp_df_all.empty:
            del_emp = st.selectbox(
                "Select Employee to Delete",
                [f"{r['emp_code']} - {r['name']}" for _, r in emp_df_all.iterrows()],
                key="del_emp",
            )
            if st.button("🗑️ Delete Employee", type="primary"):
                del_code = del_emp.split(" - ")[0]
                conn.execute("DELETE FROM employees WHERE emp_code = ?", (del_code,))
                conn.commit()
                st.warning(f"Deleted {del_emp}.")
                st.rerun()
        else:
            st.info("No employees registered yet.")

    with col2:
        _render_payment_type_ledger(
            conn,
            payment_type="Employee Tab (Monthly Account)",
            key_prefix="emp_ledger",
        )


# =====================================================================
# SHARED HELPER: LEDGER FOR A GIVEN PAYMENT TYPE
# (used by Employee Ledger, Hospital Expense, and OT Expense pages)
# =====================================================================

def _render_payment_type_ledger(conn: sqlite3.Connection, payment_type: str, key_prefix: str) -> None:
    st.subheader("Ledger Date Range")
    start_dt, end_dt = date_range_picker(key_prefix)
    start_str, end_str = sql_bounds(start_dt, end_dt)

    employees = pd.read_sql("SELECT * FROM employees ORDER BY name", conn)
    
    # Build selection list with an option for general/unlinked expenses
    emp_options = ["All / General (No Employee)"]
    if not employees.empty:
        emp_options.extend([f"{r['emp_code']} - {r['name']}" for _, r in employees.iterrows()])

    emp_sel = st.selectbox(
        "View Individual Ledger For:",
        emp_options,
        key=f"{key_prefix}_emp_select",
    )

    # Filter query based on employee selection
    if emp_sel == "All / General (No Employee)":
        ledger_df = pd.read_sql(
            """SELECT s.sale_date AS 'Date & Time', 
                      COALESCE(e.name, 'General / Unlinked') AS 'Employee',
                      p.name AS 'Item Name', si.qty AS 'Quantity',
                      si.retail_price AS 'Unit Price (PKR)',
                      (si.qty * si.retail_price) AS 'Total Bill (PKR)'
               FROM sales s
               JOIN sale_items si ON s.id = si.sale_id
               JOIN products p ON si.product_id = p.id
               LEFT JOIN employees e ON s.emp_id = e.id
               WHERE s.payment_type = ?
                 AND s.sale_date BETWEEN ? AND ?
               ORDER BY s.sale_date DESC""",
            conn,
            params=(payment_type, start_str, end_str),
        )
    else:
        selected_code = emp_sel.split(" - ")[0]
        emp_id = int(employees[employees["emp_code"] == selected_code]["id"].values[0])
        ledger_df = pd.read_sql(
            """SELECT s.sale_date AS 'Date & Time', p.name AS 'Item Name', si.qty AS 'Quantity',
                      si.retail_price AS 'Unit Price (PKR)',
                      (si.qty * si.retail_price) AS 'Total Bill (PKR)'
               FROM sales s
               JOIN sale_items si ON s.id = si.sale_id
               JOIN products p ON si.product_id = p.id
               WHERE s.emp_id = ? AND s.payment_type = ?
                 AND s.sale_date BETWEEN ? AND ?
               ORDER BY s.sale_date DESC""",
            conn,
            params=(emp_id, payment_type, start_str, end_str),
        )

    st.dataframe(ledger_df, use_container_width=True)
    total_due = ledger_df["Total Bill (PKR)"].sum() if not ledger_df.empty else 0
    st.metric(f"Total Expense ({start_dt} to {end_dt})", f"PKR {total_due:,.2f}")

    st.divider()
    st.header("📑 Expense Summary Export")
    st.caption(f"Using the date range selected above: {start_dt} to {end_dt}")

    summary_df = pd.read_sql(
        """SELECT COALESCE(e.emp_code, 'N/A') AS 'Employee ID', 
                  COALESCE(e.name, 'General / Unlinked Expense') AS 'Employee Name', 
                  COALESCE(e.dept, 'General') AS 'Department',
                  COUNT(DISTINCT s.id) AS 'Total Transactions',
                  SUM(si.qty * si.retail_price) AS 'Total Amount (PKR)'
           FROM sales s
           JOIN sale_items si ON s.id = si.sale_id
           LEFT JOIN employees e ON s.emp_id = e.id
           WHERE s.payment_type = ?
             AND s.sale_date BETWEEN ? AND ?
           GROUP BY s.emp_id
           ORDER BY e.name""",
        conn,
        params=(payment_type, start_str, end_str),
    )

    detailed_df = pd.read_sql(
        """SELECT s.sale_date AS 'Date & Time', 
                  COALESCE(e.emp_code, 'N/A') AS 'Employee ID', 
                  COALESCE(e.name, 'General') AS 'Employee Name',
                  COALESCE(e.dept, 'General') AS 'Department', 
                  p.name AS 'Item', si.qty AS 'Quantity',
                  si.retail_price AS 'Unit Price (PKR)', (si.qty * si.retail_price) AS 'Total (PKR)'
           FROM sales s
           JOIN sale_items si ON s.id = si.sale_id
           JOIN products p ON si.product_id = p.id
           LEFT JOIN employees e ON s.emp_id = e.id
           WHERE s.payment_type = ?
             AND s.sale_date BETWEEN ? AND ?
           ORDER BY s.sale_date DESC""",
        conn,
        params=(payment_type, start_str, end_str),
    )

    if summary_df.empty:
        st.info(f"No '{payment_type}' records found for this date range.")
    else:
        st.write("### Summary Totals")
        st.dataframe(summary_df, use_container_width=True)

        excel_out = to_excel_bytes({"Summary": summary_df, "Detailed Log": detailed_df})
        st.download_button(
            f"📥 Download {payment_type} Report (.xlsx)",
            data=excel_out,
            file_name=f"{payment_type.replace(' ', '_')}_Report_{start_dt}_to_{end_dt}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}_download",
        )

# =====================================================================
# MODULE: HOSPITAL EXPENSE (same functionality as Employee Ledger,
# but filtered to the 'Hospital Expense' payment type and kept separate)
# =====================================================================

def module_hospital_expense(conn: sqlite3.Connection) -> None:
    st.header("🏥 Hospital Expense Ledger")
    st.caption("Tracks purchases charged as Hospital Expense from Point of Sale.")
    _render_payment_type_ledger(conn, payment_type="Hospital Expense", key_prefix="hospital")


# =====================================================================
# MODULE: OT EXPENSE (same functionality as Employee Ledger,
# but filtered to the 'OT Expense' payment type and kept separate)
# =====================================================================

def module_ot_expense(conn: sqlite3.Connection) -> None:
    st.header("🕒 OT Expense Ledger")
    st.caption("Tracks purchases charged as OT Expense from Point of Sale.")
    _render_payment_type_ledger(conn, payment_type="OT Expense", key_prefix="ot")


# =====================================================================
# MODULE: PROFIT & SALES REPORTS
# =====================================================================

def module_reports(conn: sqlite3.Connection) -> None:
    st.header("📊 Sales & Profitability Analysis")

    start_dt, end_dt = date_range_picker("sales_report")
    start_str, end_str = sql_bounds(start_dt, end_dt)

    report_df = pd.read_sql(
        """SELECT p.name AS 'Product Name',
                  SUM(si.qty) AS 'Units Sold',
                  SUM(si.qty * si.cost_price) AS 'Total Cost (PKR)',
                  SUM(si.qty * si.retail_price) AS 'Total Revenue (PKR)',
                  SUM(si.qty * (si.retail_price - si.cost_price)) AS 'Profit (PKR)'
           FROM sale_items si
           JOIN sales s ON si.sale_id = s.id
           JOIN products p ON si.product_id = p.id
           WHERE s.sale_date BETWEEN ? AND ?
           GROUP BY p.id
           ORDER BY p.name""",
        conn,
        params=(start_str, end_str),
    )

    if report_df.empty:
        st.info("No sales records found for the selected period.")
        return

    st.dataframe(report_df, use_container_width=True)

    tot_revenue = report_df["Total Revenue (PKR)"].sum()
    tot_cost = report_df["Total Cost (PKR)"].sum()
    tot_profit = report_df["Profit (PKR)"].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Gross Sales", f"PKR {tot_revenue:,.2f}")
    c2.metric("Cost of Goods Sold", f"PKR {tot_cost:,.2f}")
    c3.metric("Net Profit", f"PKR {tot_profit:,.2f}")

    excel_report = to_excel_bytes({"Sales & Profit Report": report_df})
    st.download_button(
        "📥 Download Report (.xlsx)",
        data=excel_report,
        file_name=f"Sales_Report_{start_dt}_to_{end_dt}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =====================================================================
# MODULE: SETTINGS (shop info + receipt / printer config)
# =====================================================================

def module_settings() -> None:
    st.header("⚙️ Settings")

    st.subheader("Shop Details (shown on printed receipts)")
    current = get_all_settings()

    with st.form("settings_form"):
        shop_name = st.text_input("Shop / Canteen Name", value=current["shop_name"])
        shop_address = st.text_input("Address (optional)", value=current["shop_address"])
        shop_phone = st.text_input("Phone Number (optional)", value=current["shop_phone"])
        receipt_footer = st.text_input("Receipt Footer Message", value=current["receipt_footer"])
        paper_width = st.selectbox(
            "Thermal Paper Width",
            ["58", "80"],
            index=["58", "80"].index(current["paper_width_mm"]) if current["paper_width_mm"] in ("58", "80") else 1,
            help="Choose whichever matches your thermal printer's paper roll.",
        )
        saved = st.form_submit_button("💾 Save Settings")

    if saved:
        set_setting("shop_name", shop_name.strip() or DEFAULT_SETTINGS["shop_name"])
        set_setting("shop_address", shop_address.strip())
        set_setting("shop_phone", shop_phone.strip())
        set_setting("receipt_footer", receipt_footer.strip())
        set_setting("paper_width_mm", paper_width)
        st.success("Settings saved.")
        st.rerun()

    st.divider()
    st.subheader("👤 Shop Keeper / Shop Person Access")

    st.info(
        "Admin has full access to every option. Create each Shop Keeper account below "
        "and choose exactly which modules they may use. Paid-bill editing, Excel-data "
        "editing, Settings, and Reset All Data are never granted to a Shop Keeper."
    )

    shop_conn = get_conn()
    shop_user_options = pd.read_sql(
        "SELECT username FROM users WHERE role = 'shop' ORDER BY username", shop_conn
    )
    shop_conn.close()
    if not shop_user_options.empty:
        selected_shop = st.selectbox(
            "Manage Existing Shop Keeper",
            shop_user_options["username"].tolist(),
            key="manage_shop_keeper",
        )
        current_perms = get_user_permissions(selected_shop)
        with st.form("edit_shop_keeper_permissions"):
            st.write("**Access points**")
            selected_perms = []
            for key, label in ACCESS_OPTIONS.items():
                if st.checkbox(
                    label,
                    value=key in current_perms,
                    key=f"perm_{selected_shop}_{key}",
                ):
                    selected_perms.append(key)

            c1, c2 = st.columns(2)
            with c1:
                save_perms = st.form_submit_button("💾 Save Access")
            with c2:
                remove_access = st.form_submit_button("🚫 Remove All Access")

        if save_perms:
            save_user_permissions(selected_shop, selected_perms)
            st.success(f"Access updated for '{selected_shop}'.")
            st.rerun()

        if remove_access:
            save_user_permissions(selected_shop, [])
            st.success(f"All optional access removed from '{selected_shop}'.")
            st.rerun()

    st.markdown("### Create New Shop Keeper")
    with st.form("add_shop_person_form", clear_on_submit=True):
        sp_user = st.text_input("Shop Keeper Username")
        sp_pass = st.text_input("Shop Keeper Password", type="password")
        st.write("**Choose Access Points**")
        new_permissions = []
        for key, label in ACCESS_OPTIONS.items():
            if st.checkbox(label, value=True, key=f"new_perm_{key}"):
                new_permissions.append(key)

        sp_submit = st.form_submit_button("➕ Create Shop Keeper Account")

    if sp_submit:
        if not sp_user.strip() or not sp_pass:
            st.error("Username and password cannot be empty.")
        elif len(sp_pass) < 4:
            st.error("Password should be at least 4 characters.")
        else:
            ok, msg = create_user(
                sp_user,
                sp_pass,
                role="shop",
                permissions=new_permissions,
            )
            st.success(msg) if ok else st.error(msg)

    st.divider()
    st.subheader("🧨 Reset All Data")
    st.warning(
        "This permanently deletes ALL canteen data from the database: "
        "products/inventory, employees, sales, sale items, admin accounts, "
        "and saved shop/receipt settings. This cannot be undone."
    )

    confirm_reset = st.checkbox(
        "I understand that ALL data will be permanently deleted.",
        key="confirm_reset_all_data",
    )

    if st.button(
        "🗑️ RESET ALL DATA",
        type="primary",
        disabled=not confirm_reset,
        key="reset_all_data_button",
    ):
        reset_all_data()

        # Clear the current login/session so the app returns to first-time setup.
        st.session_state.clear()
        st.success("All data has been permanently deleted.")
        st.rerun()

    st.divider()
    st.subheader("🖨️ About Thermal Printing")
    st.markdown(
        "- Install your thermal printer as a normal printer in Windows first "
        "(most 58mm/80mm printers ship with a driver, or work with Windows' "
        "generic/text driver).\n"
        "- After any Point of Sale checkout, a **Print Receipt** button appears — "
        "clicking it opens your browser's print dialog scoped to just the receipt.\n"
        "- In the print dialog, select your thermal printer and, if available, "
        "set paper size/margins to match your roll width for the cleanest cut."
    )


# =====================================================================
# MAIN APP
# =====================================================================

def main() -> None:
    st.set_page_config(page_title="Canteen Management System", layout="wide")
    init_db()

    if not login_gate():
        return  # login_gate already rendered the login/setup screen

    st.sidebar.title("🍽️ CMASH Canteen")
    current_role = st.session_state.get("role", "admin")
    role_label = "Admin" if current_role == "admin" else "Shop Keeper"
    st.sidebar.caption(f"Logged in as **{st.session_state.get('username', 'user')}**")
    st.sidebar.caption(f"Role: **{role_label}**")

    if st.sidebar.button("🚪 Log Out"):
        st.session_state.clear()
        st.rerun()

    st.sidebar.divider()

    # Admin sees every module, including Settings.
    # Shop Keeper sees only the access points selected by Admin.
    menu_options = []
    permission_to_menu = [
        ("pos", "Point of Sale (POS)"),
        ("returns", "Returns"),
        ("inventory", "Inventory Management"),
        ("employees", "Employee Ledger"),
        ("hospital", "Hospital Expense"),
        ("ot", "OT Expense"),
        ("reports", "Profit & Sales Reports"),
    
    ]

    for permission, label in permission_to_menu:
        if user_has_permission(permission):
            menu_options.append(label)

    if current_role == "admin":
        menu_options.append("Settings")

    if not menu_options:
        st.warning("No access points have been assigned to this Shop Keeper. Ask the Admin.")
        return

    menu = st.sidebar.radio("Navigation", menu_options)

    conn = get_conn()
    try:
        if menu == "Point of Sale (POS)":
            module_pos(conn)
        elif menu == "Returns":
            st.header("↩️ Returns")
            st.info(
                "Returns must be recorded against the original paid bill. "
                "The original paid bill is never edited."
            )
            st.warning(
                "Return processing is reserved for the Returns module. "
                "Your existing sales database is preserved."
            )
        elif menu == "Inventory Management":
            module_inventory(conn)
        elif menu == "Product Management":
            module_inventory(conn)
        elif menu == "Price Management":
            module_inventory(conn)
        elif menu == "Stock Management":
            module_inventory(conn)
        elif menu == "Employee Ledger":
            module_employee_ledger(conn)
        elif menu == "Hospital Expense":
            module_hospital_expense(conn)
        elif menu == "OT Expense":
            module_ot_expense(conn)
        elif menu == "Profit & Sales Reports":
            module_reports(conn)
        elif menu == "Excel Reports":
            module_reports(conn)
        elif menu == "Settings":
            if current_role != "admin":
                st.error("Access denied. Settings are Admin-only.")
                return
            module_settings()
    finally:
        conn.close()


if __name__ == "__main__":
    main()
