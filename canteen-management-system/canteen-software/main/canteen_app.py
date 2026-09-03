
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
import json
import secrets
import re
import sqlite3
from datetime import date, datetime, timedelta
import libsql_client
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Compatibility fix for older Streamlit versions
if not hasattr(st, "rerun"):
    st.rerun = st.experimental_rerun

# Payment types that must be tied to a specific employee (deducted from their account)
EMPLOYEE_LINKED_PAYMENT_TYPES = [
    "Employee Tab (Monthly Account)",
]

# =====================================================================
# DATABASE LAYER
# =====================================================================


_TABLES_WITH_ID_PK = {
    "users",
    "products",
    "employees",
    "sales",
    "sale_items",
    "employee_payments",
    "returns",
}


class _TursoCursor:
    """Mimics the small slice of sqlite3.Cursor's API this app uses
    (execute / fetchone / fetchall / lastrowid), backed by a libsql_client
    sync Client. libsql_client's ResultSet doesn't have a Python
    sqlite3-style .lastrowid, so INSERTs into tables that have an 'id'
    primary key are transparently rewritten to add 'RETURNING id' and we
    read the id back from the normal row data. Tables without an 'id'
    column (e.g. 'settings', keyed by 'key') are left alone - blindly
    adding RETURNING id there would reference a column that doesn't exist
    and break the query."""

    def __init__(self, client):
        self._client = client
        self.lastrowid = None
        self._result = None

    def execute(self, sql, params=None):
        params = list(params) if params else []
        stripped = sql.strip()
        upper = stripped.upper()

        is_insert = upper.startswith("INSERT")
        wants_returning_id = False
        if is_insert and "RETURNING" not in upper:
            m = re.match(r"INSERT\s+INTO\s+(\w+)", stripped, re.IGNORECASE)
            table_name = m.group(1).lower() if m else None
            if table_name in _TABLES_WITH_ID_PK:
                stripped = stripped.rstrip().rstrip(";") + " RETURNING id"
                wants_returning_id = True

        self._result = self._client.execute(stripped, params)

        if wants_returning_id:
            self.lastrowid = self._result.rows[0][0] if self._result.rows else None

        return self

    def fetchone(self):
        return self._result.rows[0] if self._result and self._result.rows else None

    def fetchall(self):
        return list(self._result.rows) if self._result else []

    @property
    def columns(self):
        return self._result.columns if self._result else []

    @property
    def rows(self):
        return self._result.rows if self._result else []

    @property
    def description(self):
        # Minimal DBAPI2-style description shape: [(name, ...), ...]
        return [(c,) for c in self.columns]


class _TursoConn:
    """Wraps a libsql_client sync Client to look like sqlite3.Connection
    (execute / cursor / commit / close) so the rest of this app - written
    against that API - works unchanged against Turso.

    libsql_client auto-commits every execute() call over HTTP (there is no
    open multi-statement transaction to commit unless you explicitly use
    client.transaction()), so commit() here is a safe no-op."""

    def __init__(self, client):
        self._client = client

    def execute(self, sql, params=None) -> _TursoCursor:
        cur = self.cursor()
        cur.execute(sql, params)
        return cur

    def cursor(self) -> _TursoCursor:
        return _TursoCursor(self._client)

    def commit(self) -> None:
        pass  # each execute() already commits; nothing to do

    def close(self) -> None:
        self._client.close()


def get_conn() -> _TursoConn:
    """Opens a connection to the persistent Turso (libSQL) database
    configured in .streamlit/secrets.toml. This is what makes data survive
    app restarts/redeploys - unlike a local SQLite file, which Streamlit
    Community Cloud wipes on every reboot.

    Uses the https:// scheme (not libsql://) so libsql_client talks over
    plain HTTP instead of a WebSocket. Streamlit Community Cloud's network
    frequently fails the WebSocket handshake for wss://, causing
    aiohttp.client_exceptions.WSServerHandshakeError - the HTTP transport
    avoids that entirely and works reliably in this environment."""
    url = st.secrets["TURSO_URL"]
    if url.startswith("libsql://"):
        url = "https://" + url[len("libsql://"):]
    client = libsql_client.create_client_sync(
        url=url, auth_token=st.secrets["TURSO_AUTH_TOKEN"]
    )
    return _TursoConn(client)


def _is_unique_violation(exc: Exception) -> bool:
    """libsql_client doesn't raise sqlite3.IntegrityError, so we detect a
    UNIQUE-constraint violation by message text instead of exception type."""
    msg = str(exc).lower()
    return "unique" in msg and ("constraint" in msg or "violat" in msg)


def _extract_select_column_names(sql: str) -> list:
    """Best-effort fallback for figuring out a SELECT's output column names
    purely from the query text. Used only when the DB driver doesn't return
    column metadata for a zero-row result (a libsql_client/HTTP quirk seen
    with GROUP BY / JOIN queries that match no rows) - without this, a
    brand-new report or employee with no transactions yet would come back
    as a DataFrame with NO columns at all, breaking any code that expects
    a named column (e.g. df["Type"], a .merge(on=...))."""
    m = re.search(r"SELECT\s+(.*?)\s+FROM\s", sql, re.IGNORECASE | re.DOTALL)
    if not m:
        return []
    col_clause = m.group(1).strip()
    if col_clause == "*":
        return []  # can't know real column names without schema introspection

    # Split on top-level commas only (ignore commas inside parentheses).
    parts, depth, current = [], 0, ""
    for ch in col_clause:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append(current)
            current = ""
        else:
            current += ch
    if current.strip():
        parts.append(current)

    names = []
    for part in parts:
        part = part.strip()
        alias_match = (
            re.search(r"AS\s+'([^']+)'\s*$", part, re.IGNORECASE)
            or re.search(r'AS\s+"([^"]+)"\s*$', part, re.IGNORECASE)
            or re.search(r"AS\s+(\w+)\s*$", part, re.IGNORECASE)
        )
        if alias_match:
            names.append(alias_match.group(1))
        else:
            # No alias - use the bare column name (last dotted segment).
            bare = part.split(".")[-1].strip()
            names.append(bare if bare else part)
    return names


def query_to_dataframe(
    query: str, conn=None, params: list = None
) -> pd.DataFrame:
    """Executes a SQL query using libsql_client and returns a pandas DataFrame."""
    if isinstance(conn, (list, tuple)):
        params = conn
        conn = None

    if params is None:
        params = []
    elif not isinstance(params, (list, tuple)):
        params = [params]

    should_close = False
    if conn is None:
        conn = get_conn()
        should_close = True

    try:
        res = conn.execute(query, params)
        columns = res.columns
        rows = [list(row) for row in res.rows]
        if not columns and not rows:
            # Zero rows with no column metadata - recover expected column
            # names from the query text so callers still get a properly
            # shaped (just empty) DataFrame instead of one with no columns.
            columns = _extract_select_column_names(query)
        return pd.DataFrame(rows, columns=columns if columns else None)
    except KeyError:
        st.error(
            f"Database Execution Error: Turso rejected query: `{query}`. Verify table and column names exist."
        )
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Database Error: {e}")
        return pd.DataFrame()
    finally:
        if should_close:
            conn.close()

def init_db() -> None:
    conn = get_conn()

    conn.execute("""CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    salt TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT 'admin',
                    permissions TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    barcode TEXT UNIQUE,
                    name TEXT UNIQUE NOT NULL,
                    cost_price REAL NOT NULL,
                    retail_price REAL NOT NULL,
                    stock_qty INTEGER NOT NULL DEFAULT 0
                )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    emp_code TEXT UNIQUE NOT NULL,
                    name TEXT NOT NULL,
                    dept TEXT
                )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS sales (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    payment_type TEXT NOT NULL,
                    emp_id INTEGER,
                    total_amount REAL NOT NULL,
                    sale_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(emp_id) REFERENCES employees(id)
                )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS sale_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sale_id INTEGER NOT NULL,
                    product_id INTEGER NOT NULL,
                    qty INTEGER NOT NULL,
                    cost_price REAL NOT NULL,
                    retail_price REAL NOT NULL,
                    FOREIGN KEY(sale_id) REFERENCES sales(id),
                    FOREIGN KEY(product_id) REFERENCES products(id)
                )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS employee_payments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    emp_id INTEGER NOT NULL,
                    amount REAL NOT NULL,
                    payment_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    note TEXT,
                    FOREIGN KEY(emp_id) REFERENCES employees(id)
                )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS returns (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sale_id INTEGER NOT NULL,
                    sale_item_id INTEGER NOT NULL,
                    product_id INTEGER NOT NULL,
                    qty INTEGER NOT NULL,
                    refund_amount REAL NOT NULL,
                    return_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    note TEXT,
                    FOREIGN KEY(sale_id) REFERENCES sales(id),
                    FOREIGN KEY(sale_item_id) REFERENCES sale_items(id),
                    FOREIGN KEY(product_id) REFERENCES products(id)
                )""")

    conn.close()
    migrate_db()


def reset_all_data() -> None:
    """Permanently delete all application data from the database."""
    conn = get_conn()
    try:
        # Delete child records first because foreign-key enforcement is enabled.
        conn.execute("DELETE FROM returns")
        conn.execute("DELETE FROM sale_items")
        conn.execute("DELETE FROM sales")
        conn.execute("DELETE FROM employee_payments")
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

def get_setting(key: str, default: str = "") -> str:
    conn = get_conn()
    try:
        res = conn.execute("SELECT value FROM settings WHERE key = ?", [key])
        return res.rows[0][0] if res.rows else default
    finally:
        conn.close()


def set_setting(key: str, value: str) -> None:
    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            [key, value],
        )
    finally:
        conn.close()


def get_all_settings() -> dict:
    return {k: get_setting(k) for k in DEFAULT_SETTINGS}



def migrate_db() -> None:
    conn = get_conn()
    try:
        # Check existing columns in products table
        res = conn.execute("PRAGMA table_info(products)")
        # libsql_client returns ResultSet objects containing Row objects
        columns = [row[1] for row in res.rows]

        if "barcode" not in columns:
            conn.execute("ALTER TABLE products ADD COLUMN barcode TEXT")

        if "stock_qty" not in columns:
            conn.execute(
                "ALTER TABLE products ADD COLUMN stock_qty INTEGER NOT NULL DEFAULT 0"
            )

        # Ensure index on barcode
        conn.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_products_barcode ON products(barcode) WHERE barcode IS NOT NULL"
        )

        # Check existing columns in users table (older databases created
        # before role-based access was added won't have these yet).
        res_users = conn.execute("PRAGMA table_info(users)")
        user_columns = [row[1] for row in res_users.rows]

        if "role" not in user_columns:
            # Any account that already exists was created via the original
            # "admin" setup flow, so it's correct to default it to admin.
            conn.execute("ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT 'admin'")

        if "permissions" not in user_columns:
            conn.execute("ALTER TABLE users ADD COLUMN permissions TEXT")
    finally:
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
    "reports_center": "Reports Center (All Reports)",
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


def get_user_permissions(username: str) -> list:
    conn = get_conn()
    try:
        res = conn.execute(
            "SELECT permissions FROM users WHERE username = ?", [username]
        )
        if res.rows and res.rows[0][0]:
            # Convert JSON string or comma-separated list into a python list
            val = res.rows[0][0]
            if isinstance(val, str):
                try:
                    return json.loads(val)
                except Exception:
                    return [p.strip() for p in val.split(",") if p.strip()]
            return list(val)
        return []
    except Exception:
        # Fallback if 'permissions' column doesn't exist in the users table
        return []
    finally:
        conn.close()

def set_user_permissions(username: str, permissions: list) -> bool:
    conn = get_conn()
    try:
        perms_json = json.dumps(permissions)
        conn.execute(
            "UPDATE users SET permissions = ? WHERE username = ?",
            [perms_json, username],
        )
        return True
    except Exception:
        return False
    finally:
        conn.close()
        
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

# =====================================================================
# AUTH / USER HELPERS
# =====================================================================


def any_user_exists() -> bool:
    """Check if any admin/user account exists in the system."""
    conn = get_conn()
    try:
        res = conn.execute("SELECT COUNT(*) FROM users")
        n = res.rows[0][0] if res.rows else 0
        return n > 0
    finally:
        conn.close()


def get_user_by_username(username: str):
    """Retrieve user credentials (including role/permissions) safely."""
    conn = get_conn()
    try:
        res = conn.execute(
            "SELECT id, username, password_hash, salt, role, permissions FROM users WHERE username = ?",
            [username],
        )
        return res.rows[0] if res.rows else None
    finally:
        conn.close()


def create_user(
    username: str,
    password_raw: str,
    role: str = "admin",
    permissions: list | None = None,
) -> tuple[bool, str]:
    """Create a new user with a salted PBKDF2 hash. Returns (success, message)."""
    username = (username or "").strip()
    if not username or not password_raw:
        return False, "Username and password cannot be empty."

    salt = secrets.token_hex(16)
    password_hash = hashlib.pbkdf2_hmac(
        "sha256", password_raw.encode(), salt.encode(), 100000
    ).hex()
    perms_json = json.dumps(normalize_permissions(permissions)) if permissions is not None else None

    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO users (username, password_hash, salt, role, permissions) VALUES (?, ?, ?, ?, ?)",
            [username, password_hash, salt, role, perms_json],
        )
        return True, "Account created."
    except Exception as e:
        if _is_unique_violation(e):
            return False, "That username already exists."
        return False, f"Could not create account: {e}"
    finally:
        conn.close()


def verify_password(stored_hash: str, salt: str, password_raw: str) -> bool:
    """Verify password match."""
    computed_hash = hashlib.pbkdf2_hmac(
        "sha256", password_raw.encode(), salt.encode(), 100000
    ).hex()
    return secrets.compare_digest(stored_hash, computed_hash)


def verify_login(username: str, password_raw: str):
    user = get_user_by_username(username)
    if not user:
        return False, None, []

    # Column order matches the SELECT in get_user_by_username:
    # id, username, password_hash, salt, role, permissions
    stored_hash = user[2]
    salt = user[3]
    role = user[4] if len(user) > 4 and user[4] else "admin"
    perms_raw = user[5] if len(user) > 5 else None

    if verify_password(stored_hash, salt, password_raw):
        return True, role, normalize_permissions(perms_raw)

    return False, None, []


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

def to_excel_bytes(
    sheets: dict[str, pd.DataFrame],
    start_dt: date | None = None,
    end_dt: date | None = None,
) -> io.BytesIO:
    """Writes each dataframe to its own sheet, with a small header showing the
    exact date/time the report was generated (and the date range, if given)
    at the top of every sheet."""
    output = io.BytesIO()
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            meta_rows = [["Generated On:", generated_at]]
            if start_dt and end_dt:
                meta_rows.append(["Report Date Range:", f"{start_dt} to {end_dt}"])
            meta_df = pd.DataFrame(meta_rows)
            meta_df.to_excel(writer, index=False, header=False, sheet_name=safe_name, startrow=0)
            df.to_excel(writer, index=False, sheet_name=safe_name, startrow=len(meta_rows) + 1)
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
# EMPLOYEE PAYMENTS (CREDITS AGAINST THE MONTHLY TAB)
# =====================================================================

def add_employee_payment(
    conn, emp_id: int, amount: float, payment_date: date, note: str
) -> None:
    conn.execute(
        "INSERT INTO employee_payments (emp_id, amount, payment_date, note) VALUES (?, ?, ?, ?)",
        (emp_id, amount, f"{payment_date} 00:00:00", note.strip()),
    )
    conn.commit()


def delete_employee_payment(conn: sqlite3.Connection, payment_id: int) -> None:
    conn.execute("DELETE FROM employee_payments WHERE id = ?", (payment_id,))
    conn.commit()


# =====================================================================
# THERMAL RECEIPT PRINTING
# =====================================================================

def render_receipt(sale: dict) -> None:
    """Renders a printable thermal-receipt widget (58mm or 80mm) with a
    Print button instantly from session state."""
    settings = get_all_settings()
    paper_mm = settings.get("paper_width_mm", "80") or "80"

    items_rows = "".join([
        f"<tr>"
        f"<td style='text-align:left'>{it['name']}</td>"
        f"<td style='text-align:center'>{it['qty']}</td>"
        f"<td style='text-align:right'>{it['price']:,.2f}</td>"
        f"<td style='text-align:right'>{it['line_total']:,.2f}</td>"
        f"</tr>"
        for it in sale["items"]
    ])

    emp_line = f"<p style='margin:2px 0;'>Employee: {sale['emp_label']}</p>" if sale.get("emp_label") else ""
    address_line = f"<p style='margin:2px 0;'>{settings['shop_address']}</p>" if settings.get("shop_address") else ""
    phone_line = f"<p style='margin:2px 0;'>{settings['shop_phone']}</p>" if settings.get("shop_phone") else ""

    receipt_body = f"""
    <div id="receipt-wrapper" style="display:flex; flex-direction:column; align-items:center; font-family:sans-serif;">
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
        <p style="text-align:center; font-size:11px; margin:2px 0;">{settings['receipt_footer']}</p>
      </div>
      <button onclick="window.print()"
              style="margin-top:10px; padding:8px 20px; font-size:14px; cursor:pointer;
                     background:#0f9d58; color:#fff; border:none; border-radius:6px; font-weight:bold;">
        🖨️ Print Receipt
      </button>
    </div>
    <style>
      @media print {{
        button {{ display: none !important; }}
        body {{ margin: 0; padding: 0; }}
      }}
    </style>
    """

    components.html(receipt_body, height=450, scrolling=False)


# =====================================================================
# MODULE: POINT OF SALE
# =====================================================================

def module_pos(conn: sqlite3.Connection) -> None:
    st.header("🛒 Point of Sale & Checkout")
    st.info("Paid bills cannot be edited. Use **Return Paid Sale** to process a return without changing the original bill.")

    products = query_to_dataframe(
    "SELECT * FROM products WHERE stock_qty > 0", conn=conn)
    employees = query_to_dataframe("SELECT * FROM employees", conn=conn)

    if products.empty:
        st.warning("No items available in inventory. Please add products first.")
        return

    if "pos_selected_product_id" not in st.session_state:
        st.session_state["pos_selected_product_id"] = None
    if "pos_cart" not in st.session_state:
        st.session_state["pos_cart"] = []  # list of dicts: product_id, name, qty, price, cost_price

    cart = st.session_state["pos_cart"]
    cart_qty_by_product = {c["product_id"]: c["qty"] for c in cart}

    st.subheader("➕ Add Items to Bill")
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

    if item_row is not None:
        already_in_cart = cart_qty_by_product.get(int(item_row["id"]), 0)
        remaining_stock = int(item_row["stock_qty"]) - already_in_cart

        if remaining_stock <= 0:
            st.warning(
                f"All {int(item_row['stock_qty'])} units of '{item_row['name']}' are already in the cart."
            )
        else:
            ac1, ac2 = st.columns([2, 1])
            with ac1:
                st.info(
                    f"**{item_row['name']}**  |  Available: {remaining_stock}  |  "
                    f"Retail Price: PKR {item_row['retail_price']:,.2f}"
                )
                add_qty = st.number_input(
                    "Quantity to Add", min_value=1, max_value=remaining_stock, value=1, key="pos_add_qty"
                )
            with ac2:
                st.write("")
                st.write("")
                if st.button("➕ Add to Cart", type="primary", use_container_width=True):
                    for c in cart:
                        if c["product_id"] == int(item_row["id"]):
                            c["qty"] += int(add_qty)
                            break
                    else:
                        cart.append(
                            {
                                "product_id": int(item_row["id"]),
                                "name": item_row["name"],
                                "qty": int(add_qty),
                                "price": float(item_row["retail_price"]),
                                "cost_price": float(item_row["cost_price"]),
                            }
                        )
                    st.session_state["pos_cart"] = cart
                    st.session_state["pos_selected_product_id"] = None
                    st.rerun()

    st.divider()
    st.subheader(f"🧺 Current Bill ({len(cart)} item{'s' if len(cart) != 1 else ''})")

    if not cart:
        st.info("Cart is empty. Find a product above and click 'Add to Cart' to start a bill.")
        return

    cart_total = 0.0
    for idx, c in enumerate(cart):
        # Re-check current stock in case it changed since the item was added.
        live_stock_row = products[products["id"] == c["product_id"]]
        live_stock = int(live_stock_row.iloc[0]["stock_qty"]) if not live_stock_row.empty else c["qty"]

        row_c1, row_c2, row_c3, row_c4 = st.columns([3, 2, 2, 1])
        with row_c1:
            st.write(f"**{c['name']}**")
        with row_c2:
            new_qty = st.number_input(
                "Qty",
                min_value=1,
                max_value=max(live_stock, c["qty"]),
                value=c["qty"],
                key=f"cart_qty_{c['product_id']}",
                label_visibility="collapsed",
            )
            if new_qty != c["qty"]:
                cart[idx]["qty"] = int(new_qty)
                st.session_state["pos_cart"] = cart
                st.rerun()
        with row_c3:
            line_total = c["qty"] * c["price"]
            cart_total += line_total
            st.write(f"PKR {c['price']:,.2f} × {c['qty']} = **PKR {line_total:,.2f}**")
        with row_c4:
            if st.button("🗑️", key=f"remove_cart_{c['product_id']}"):
                st.session_state["pos_cart"] = [x for x in cart if x["product_id"] != c["product_id"]]
                st.rerun()

    st.markdown(f"### 🧾 Bill Total: PKR {cart_total:,.2f}")

    col1, col2 = st.columns([2, 1])
    with col1:
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

        can_checkout = not (pay_type in EMPLOYEE_LINKED_PAYMENT_TYPES and employees.empty)
        checkout_clicked = st.button(
            "✅ Complete Transaction", disabled=not can_checkout, type="primary"
        )

        if st.button("🧹 Clear Cart"):
            st.session_state["pos_cart"] = []
            st.rerun()

        if checkout_clicked:
            # Re-validate stock for every cart line right before committing.
            insufficient = []
            for c in cart:
                live_row = products[products["id"] == c["product_id"]]
                live_stock = int(live_row.iloc[0]["stock_qty"]) if not live_row.empty else 0
                if c["qty"] > live_stock:
                    insufficient.append(f"{c['name']} (have {live_stock}, need {c['qty']})")

            if insufficient:
                st.error("Not enough stock for: " + "; ".join(insufficient) + ". Please adjust the cart.")
            else:
                cc = conn.cursor()
                cc.execute(
                    "INSERT INTO sales (payment_type, emp_id, total_amount) VALUES (?, ?, ?)",
                    (pay_type, selected_emp_id, cart_total),
                )
                sale_id = cc.lastrowid
                for c in cart:
                    cc.execute(
                        """INSERT INTO sale_items (sale_id, product_id, qty, cost_price, retail_price)
                           VALUES (?, ?, ?, ?, ?)""",
                        (sale_id, c["product_id"], c["qty"], c["cost_price"], c["price"]),
                    )
                    cc.execute(
                        "UPDATE products SET stock_qty = stock_qty - ? WHERE id = ?",
                        (c["qty"], c["product_id"]),
                    )
                conn.commit()

                emp_label = None
                if selected_emp_id is not None:
                    emp_row = employees[employees["id"] == selected_emp_id].iloc[0]
                    emp_label = f"{emp_row['emp_code']} - {emp_row['name']}"

                # Store sale details to render receipt immediately without requiring a forced rerun
                st.session_state["last_sale"] = {
                    "sale_id": sale_id,
                    "sale_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "payment_type": pay_type,
                    "emp_label": emp_label,
                    "total": cart_total,
                    "items": [
                        {
                            "name": c["name"],
                            "qty": c["qty"],
                            "price": c["price"],
                            "line_total": c["qty"] * c["price"],
                        }
                        for c in cart
                    ],
                }
                # Reset cart state in place
                st.session_state["pos_cart"] = []
                st.session_state["pos_selected_product_id"] = None
                st.success(f"Sale recorded! Total Bill: PKR {cart_total:,.2f}")

    with col2:
        if st.session_state.get("last_sale"):
            st.subheader("🧾 Receipt")
            render_receipt(st.session_state["last_sale"])
            if st.button("Dismiss Receipt"):
                st.session_state["last_sale"] = None
                st.rerun()


# =====================================================================
# MODULE: RETURNS
# =====================================================================

def add_return(
    conn: sqlite3.Connection,
    sale_id: int,
    sale_item_id: int,
    product_id: int,
    qty: int,
    refund_amount: float,
    note: str,
) -> None:
    conn.execute(
        """INSERT INTO returns (sale_id, sale_item_id, product_id, qty, refund_amount, note)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (sale_id, sale_item_id, product_id, qty, refund_amount, note.strip()),
    )
    conn.execute("UPDATE products SET stock_qty = stock_qty + ? WHERE id = ?", (qty, product_id))
    conn.commit()


def module_returns(conn: sqlite3.Connection) -> None:
    st.header("↩️ Returns")
    st.info(
        "Find the original sale, choose which items/quantities are being returned, "
        "and process it. This adds the quantity back to stock and adjusts the relevant "
        "ledger/report totals — the original paid bill itself is never edited or deleted."
    )

    search_mode = st.radio("Find Sale By", ["Sale ID", "Recent Sales List"], horizontal=True)

    if search_mode == "Sale ID":
        sid_input = st.number_input("Sale ID", min_value=1, step=1, key="return_sale_id_input")
        if st.button("🔍 Find Sale"):
            st.session_state["return_sale_id"] = int(sid_input)
    else:
        recent_sales = query_to_dataframe(
            "SELECT id, sale_date, payment_type, total_amount FROM sales ORDER BY sale_date DESC LIMIT 50",
            conn=conn
        )
        if recent_sales.empty:
            st.info("No sales recorded yet.")
            return
        options = {
            f"#{int(r['id'])} | {r['sale_date']} | {r['payment_type']} | PKR {r['total_amount']:,.2f}": int(r["id"])
            for _, r in recent_sales.iterrows()
        }
        sel = st.selectbox("Select a recent sale", list(options.keys()), key="return_recent_select")
        # Automatically persist the selected sale ID to session state
        st.session_state["return_sale_id"] = options[sel]

    sale_id = st.session_state.get("return_sale_id")
    if not sale_id:
        st.info("Please enter or select a Sale ID above to begin.")
        return  # Replaced st.stop() with return to prevent session state resets

    sale_row_df = query_to_dataframe("SELECT * FROM sales WHERE id = ?", conn=conn, params=(sale_id,))
    if sale_row_df.empty:
        st.error(f"No sale found with ID {sale_id}.")
        return
    sale_row = sale_row_df.iloc[0]

    emp_label = None
    if sale_row["emp_id"] is not None and not pd.isna(sale_row["emp_id"]):
        emp_df = query_to_dataframe(
            "SELECT emp_code, name FROM employees WHERE id = ?", conn=conn, params=(int(sale_row["emp_id"]),)
        )
        if not emp_df.empty:
            emp_label = f"{emp_df.iloc[0]['emp_code']} - {emp_df.iloc[0]['name']}"

    st.divider()
    st.subheader(f"Sale #{sale_id}")
    c1, c2, c3 = st.columns(3)
    c1.metric("Date", sale_row["sale_date"])
    c2.metric("Payment Type", sale_row["payment_type"])
    c3.metric("Original Total", f"PKR {sale_row['total_amount']:,.2f}")
    if emp_label:
        st.caption(f"Employee: {emp_label}")

    items_df = query_to_dataframe(
        """SELECT si.id AS sale_item_id, si.product_id, p.name, si.qty, si.retail_price,
                  COALESCE((SELECT SUM(qty) FROM returns WHERE sale_item_id = si.id), 0) AS already_returned
           FROM sale_items si
           JOIN products p ON si.product_id = p.id
           WHERE si.sale_id = ?""",
        conn=conn,
        params=(sale_id,),
    )

    if items_df.empty:
        st.warning("No items found for this sale.")
        return

    st.write("### Select quantities to return")
    return_selections = {}
    any_returnable = False

    for _, row in items_df.iterrows():
        remaining = int(row["qty"] - row["already_returned"])
        if remaining <= 0:
            st.write(f"✅ **{row['name']}** — fully returned already.")
            continue
        any_returnable = True
        rc1, rc2 = st.columns([3, 1])
        with rc1:
            st.write(
                f"**{row['name']}** — bought {int(row['qty'])}, already returned "
                f"{int(row['already_returned'])}, returnable {remaining}, "
                f"unit price PKR {row['retail_price']:,.2f}"
            )
        with rc2:
            qty = st.number_input(
                "Return Qty",
                min_value=0,
                max_value=remaining,
                value=0,
                key=f"return_qty_{row['sale_item_id']}",
                label_visibility="collapsed",
            )
        if qty > 0:
            # Store extracted primitives directly in dictionary to prevent Pandas Series errors
            return_selections[int(row["sale_item_id"])] = {
                "qty": int(qty),
                "product_id": int(row["product_id"]),
                "price": float(row["retail_price"]),
            }

    if not any_returnable:
        st.info("Every item on this sale has already been fully returned.")
        return

    return_note = st.text_input("Reason / Note (optional)", key="return_note")

    if return_selections:
        total_refund = sum(item["qty"] * item["price"] for item in return_selections.values())
        st.metric("Total Refund / Credit Amount", f"PKR {total_refund:,.2f}")

        if st.button("✅ Process Return", type="primary"):
            for sale_item_id, item_data in return_selections.items():
                refund_amt = item_data["qty"] * item_data["price"]
                add_return(
                    conn,
                    sale_id=int(sale_id),
                    sale_item_id=sale_item_id,
                    product_id=item_data["product_id"],
                    qty=item_data["qty"],
                    refund_amount=refund_amt,
                    note=return_note,
                )
            st.success(f"Return processed. PKR {total_refund:,.2f} refunded/credited and stock restored.")
            st.rerun()

    st.divider()
    st.subheader("🧾 Return History for This Sale")
    history_df = query_to_dataframe(
        """SELECT r.return_date AS 'Date', p.name AS 'Item', r.qty AS 'Qty',
                  r.refund_amount AS 'Refund (PKR)', r.note AS 'Note'
           FROM returns r
           JOIN products p ON r.product_id = p.id
           WHERE r.sale_id = ?
           ORDER BY r.return_date DESC""",
        conn=conn,
        params=(sale_id,),
    )
    if history_df.empty:
        st.caption("No returns recorded yet for this sale.")
    else:
        st.dataframe(history_df, use_container_width=True)


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
        df = query_to_dataframe(
            """SELECT id AS 'Item ID', barcode AS 'Barcode', name AS 'Item Name',
                      cost_price AS 'Cost Price (PKR)', retail_price AS 'Retail Price (PKR)',
                      stock_qty AS 'Stock Quantity'
               FROM products ORDER BY name""",
            conn=conn
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
                except Exception as e:
                    if _is_unique_violation(e):
                        st.error("That barcode is already assigned to another product.")
                    else:
                        st.error(f"Could not save product: {e}")

    # ---------------- Edit / correct existing item ----------------
    with tab3:
        st.subheader("Edit or Correct an Existing Product")
        st.caption(
            "Use this to fix a mistake — the values you enter here REPLACE the "
            "current price and stock quantity (they are not added on top)."
        )
        p_df = query_to_dataframe("SELECT * FROM products ORDER BY name", conn=conn)

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
                    except Exception as e:
                        if _is_unique_violation(e):
                            st.error("That name or barcode is already used by another product.")
                        else:
                            st.error(f"Could not update product: {e}")

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

        current_inv_df = query_to_dataframe(
            """SELECT name AS 'Item Name', barcode AS 'Barcode',
                      cost_price AS 'Unit Cost (PKR)', retail_price AS 'Unit Retail (PKR)',
                      stock_qty AS 'Current Stock Qty',
                      (cost_price * stock_qty) AS 'Total Stock Cost (PKR)',
                      (retail_price * stock_qty) AS 'Potential Retail Value (PKR)'
               FROM products ORDER BY name""",
            conn,
        )

        movement_df = query_to_dataframe(
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
            conn=conn,
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
            {"Current Inventory": current_inv_df, "Item Movement Report": movement_df},
            start_dt,
            end_dt,
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
    st.caption(
        "Tracks canteen items bought on the monthly employee tab (debit) and "
        "payments received from employees against that tab (credit)."
    )

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
                except Exception as e:
                    if _is_unique_violation(e):
                        st.error("That employee code already exists.")
                    else:
                        st.error(f"Could not register employee: {e}")
            else:
                st.error("Employee code and name are required.")

        st.divider()
        st.subheader("Remove Employee")
        emp_df_all = query_to_dataframe("SELECT * FROM employees ORDER BY name", conn)
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
        st.subheader("Ledger Date Range")
        start_dt, end_dt = date_range_picker("emp_ledger")
        _render_employee_ledger_content(conn, start_dt, end_dt, key_prefix="emp_ledger", show_actions=True)


# =====================================================================
# EMPLOYEE LEDGER: DEBIT (PURCHASES) / CREDIT (PAYMENTS) DATA HELPERS
# =====================================================================

def _get_employee_debit_credit_summary(conn: sqlite3.Connection, start_str: str, end_str: str) -> pd.DataFrame:
    employees_df = query_to_dataframe("SELECT id, emp_code, name, dept FROM employees ORDER BY name", conn=conn)

    debit_df = query_to_dataframe(
        """SELECT s.emp_id AS id, SUM(si.qty * si.retail_price) AS debit_total
           FROM sales s
           JOIN sale_items si ON s.id = si.sale_id
           WHERE s.payment_type = 'Employee Tab (Monthly Account)'
             AND s.sale_date BETWEEN ? AND ?
           GROUP BY s.emp_id""",
        conn,
        params=(start_str, end_str),
    )
    if debit_df.empty or "id" not in debit_df.columns:
        debit_df = pd.DataFrame(columns=["id", "debit_total"])

    credit_df = query_to_dataframe(
        """SELECT emp_id AS id, SUM(amount) AS credit_total
           FROM employee_payments
           WHERE payment_date BETWEEN ? AND ?
           GROUP BY emp_id""",
        conn,
        params=(start_str, end_str),
    )
    if credit_df.empty or "id" not in credit_df.columns:
        credit_df = pd.DataFrame(columns=["id", "credit_total"])

    summary_df = employees_df.merge(debit_df, on="id", how="left").merge(credit_df, on="id", how="left")
    summary_df["debit_total"] = summary_df["debit_total"].fillna(0)
    summary_df["credit_total"] = summary_df["credit_total"].fillna(0)
    summary_df["balance"] = summary_df["debit_total"] - summary_df["credit_total"]
    return summary_df


def _get_employee_debit_credit_detail(conn: sqlite3.Connection, emp_id: int, start_str: str, end_str: str) -> pd.DataFrame:
    expected_cols = ["date_sort", "Date", "Type", "Description", "Amount (PKR)"]

    debit_rows = query_to_dataframe(
        """SELECT s.sale_date AS date_sort, s.sale_date AS 'Date',
                  'Debit (Purchase)' AS 'Type', p.name AS 'Description',
                  (si.qty * si.retail_price) AS 'Amount (PKR)'
           FROM sales s
           JOIN sale_items si ON s.id = si.sale_id
           JOIN products p ON si.product_id = p.id
           WHERE s.emp_id = ? AND s.payment_type = 'Employee Tab (Monthly Account)'
             AND s.sale_date BETWEEN ? AND ?""",
        conn,
        params=(emp_id, start_str, end_str),
    )
    if debit_rows.empty or list(debit_rows.columns) != expected_cols:
        debit_rows = pd.DataFrame(columns=expected_cols)

    credit_rows = query_to_dataframe(
        """SELECT payment_date AS date_sort, payment_date AS 'Date',
                  'Credit (Payment)' AS 'Type',
                  COALESCE(NULLIF(note, ''), 'Payment received') AS 'Description',
                  amount AS 'Amount (PKR)'
           FROM employee_payments
           WHERE emp_id = ? AND payment_date BETWEEN ? AND ?""",
        conn,
        params=(emp_id, start_str, end_str),
    )
    if credit_rows.empty or list(credit_rows.columns) != expected_cols:
        credit_rows = pd.DataFrame(columns=expected_cols)

    combined = pd.concat([debit_rows, credit_rows], ignore_index=True)
    if not combined.empty:
        combined = combined.sort_values("date_sort", ascending=False).drop(columns=["date_sort"])
    else:
        combined = combined.drop(columns=["date_sort"], errors="ignore")
    return combined


def _render_employee_ledger_content(
    conn: sqlite3.Connection,
    start_dt: date,
    end_dt: date,
    key_prefix: str = "emp_ledger",
    show_actions: bool = True,
) -> None:
    start_str, end_str = sql_bounds(start_dt, end_dt)

    employees = query_to_dataframe("SELECT * FROM employees ORDER BY name", conn=conn)
    if employees.empty:
        st.info("No employees registered yet.")
        return

    if show_actions:
        st.divider()
        st.subheader("💳 Record Payment / Credit")
        st.caption("Use this when an employee pays off some or all of their outstanding tab.")
        with st.form("record_payment_form", clear_on_submit=True):
            pay_emp_str = st.selectbox(
                "Employee",
                [f"{r['emp_code']} - {r['name']}" for _, r in employees.iterrows()],
                key=f"{key_prefix}_payment_emp",
            )
            pay_amount = st.number_input(
                "Amount Received (PKR)", min_value=0.0, step=1.0, key=f"{key_prefix}_payment_amount"
            )
            pay_date = st.date_input("Payment Date", value=date.today(), key=f"{key_prefix}_payment_date")
            pay_note = st.text_input("Note (optional)", key=f"{key_prefix}_payment_note")
            pay_submitted = st.form_submit_button("💾 Record Payment")
        if pay_submitted:
            if pay_amount <= 0:
                st.error("Amount must be greater than 0.")
            else:
                pay_emp_id = int(
                    employees[employees["emp_code"] == pay_emp_str.split(" - ")[0]]["id"].values[0]
                )
                add_employee_payment(conn, pay_emp_id, pay_amount, pay_date, pay_note)
                st.success(f"Recorded PKR {pay_amount:,.2f} payment for {pay_emp_str}.")
                st.rerun()

        st.divider()

    emp_sel = st.selectbox(
        "View Individual Ledger For:",
        [f"{r['emp_code']} - {r['name']}" for _, r in employees.iterrows()],
        key=f"{key_prefix}_view_emp_select",
    )
    emp_id = int(employees[employees["emp_code"] == emp_sel.split(" - ")[0]]["id"].values[0])

    combined_df = _get_employee_debit_credit_detail(conn, emp_id, start_str, end_str)
    st.dataframe(combined_df, use_container_width=True)

    debit_total = (
        combined_df.loc[combined_df["Type"] == "Debit (Purchase)", "Amount (PKR)"].sum()
        if not combined_df.empty
        else 0
    )
    credit_total = (
        combined_df.loc[combined_df["Type"] == "Credit (Payment)", "Amount (PKR)"].sum()
        if not combined_df.empty
        else 0
    )
    balance = debit_total - credit_total

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Purchases (Debit)", f"PKR {debit_total:,.2f}")
    m2.metric("Total Paid (Credit)", f"PKR {credit_total:,.2f}")
    m3.metric("Outstanding Balance", f"PKR {balance:,.2f}")

    if show_actions:
        raw_payments = query_to_dataframe(
            """SELECT id, payment_date, amount, note FROM employee_payments
               WHERE emp_id = ? AND payment_date BETWEEN ? AND ?
               ORDER BY payment_date DESC""",
            conn,
            params=(emp_id, start_str, end_str),
        )
        if not raw_payments.empty:
            st.caption("Remove a wrongly entered payment:")
            options = {
                f"{r['payment_date']} | PKR {r['amount']:,.2f}" + (f" | {r['note']}" if r["note"] else ""): int(
                    r["id"]
                )
                for _, r in raw_payments.iterrows()
            }
            del_label = st.selectbox(
                "Select payment to delete", list(options.keys()), key=f"{key_prefix}_del_payment"
            )
            if st.button("🗑️ Delete Payment Entry", key=f"{key_prefix}_del_payment_btn"):
                delete_employee_payment(conn, options[del_label])
                st.warning("Payment entry deleted.")
                st.rerun()

    st.divider()
    st.header("📑 Balance Summary Export (All Employees)")
    st.caption(f"Using the date range: {start_dt} to {end_dt}")

    summary_df = _get_employee_debit_credit_summary(conn, start_str, end_str)
    display_df = summary_df[(summary_df["debit_total"] > 0) | (summary_df["credit_total"] > 0)].copy()

    if display_df.empty:
        st.info("No employee tab activity (purchases or payments) in this date range.")
    else:
        display_df = display_df.rename(
            columns={
                "emp_code": "Employee ID",
                "name": "Employee Name",
                "dept": "Department",
                "debit_total": "Total Purchases (PKR)",
                "credit_total": "Total Paid (PKR)",
                "balance": "Outstanding Balance (PKR)",
            }
        )[
            [
                "Employee ID",
                "Employee Name",
                "Department",
                "Total Purchases (PKR)",
                "Total Paid (PKR)",
                "Outstanding Balance (PKR)",
            ]
        ]
        st.dataframe(display_df, use_container_width=True)

        detailed_debits = query_to_dataframe(
            """SELECT s.sale_date AS 'Date', e.emp_code AS 'Employee ID', e.name AS 'Employee Name',
                      'Debit (Purchase)' AS 'Type', p.name AS 'Description',
                      (si.qty * si.retail_price) AS 'Amount (PKR)'
               FROM sales s
               JOIN sale_items si ON s.id = si.sale_id
               JOIN products p ON si.product_id = p.id
               JOIN employees e ON s.emp_id = e.id
               WHERE s.payment_type = 'Employee Tab (Monthly Account)'
                 AND s.sale_date BETWEEN ? AND ?""",
            conn,
            params=(start_str, end_str),
        )
        detailed_credits = query_to_dataframe(
            """SELECT ep.payment_date AS 'Date', e.emp_code AS 'Employee ID', e.name AS 'Employee Name',
                      'Credit (Payment)' AS 'Type',
                      COALESCE(NULLIF(ep.note, ''), 'Payment received') AS 'Description',
                      ep.amount AS 'Amount (PKR)'
               FROM employee_payments ep
               JOIN employees e ON ep.emp_id = e.id
               WHERE ep.payment_date BETWEEN ? AND ?""",
            conn,
            params=(start_str, end_str),
        )
        detailed_all_df = pd.concat([detailed_debits, detailed_credits], ignore_index=True)
        if not detailed_all_df.empty:
            detailed_all_df = detailed_all_df.sort_values("Date", ascending=False)

        excel_out = to_excel_bytes(
            {"Balance Summary": display_df, "Detailed Debit-Credit Log": detailed_all_df},
            start_dt,
            end_dt,
        )
        st.download_button(
            "📥 Download Employee Ledger Report (.xlsx)",
            data=excel_out,
            file_name=f"Employee_Ledger_{start_dt}_to_{end_dt}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}_download",
        )


# =====================================================================
# SHARED HELPER: LEDGER FOR A GIVEN PAYMENT TYPE
# (used by Employee Ledger, Hospital Expense, and OT Expense pages)
# =====================================================================

def _render_payment_type_ledger(conn: sqlite3.Connection, payment_type: str, key_prefix: str) -> None:
    st.subheader("Ledger Date Range")
    start_dt, end_dt = date_range_picker(key_prefix)
    start_str, end_str = sql_bounds(start_dt, end_dt)

    employees = query_to_dataframe("SELECT * FROM employees ORDER BY name", conn=conn)

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
        ledger_df = query_to_dataframe(
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
        ledger_df = query_to_dataframe(
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

    summary_df = query_to_dataframe(
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

    detailed_df = query_to_dataframe(
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

        excel_out = to_excel_bytes({"Summary": summary_df, "Detailed Log": detailed_df}, start_dt, end_dt)
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

    report_df = query_to_dataframe(
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

    excel_report = to_excel_bytes({"Sales & Profit Report": report_df}, start_dt, end_dt)
    st.download_button(
        "📥 Download Report (.xlsx)",
        data=excel_report,
        file_name=f"Sales_Report_{start_dt}_to_{end_dt}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =====================================================================
# MODULE: REPORTS CENTER (all range-based reports in one place)
# =====================================================================

def module_full_reports(conn: sqlite3.Connection) -> None:
    st.header("📈 Reports Center")
    st.caption(
        "One place to pull range-based reports: overall Sales, Employee Ledger "
        "(debit/credit), Hospital Expense, and OT Expense — all for the same date range."
    )

    start_dt, end_dt = date_range_picker("reports_center")
    start_str, end_str = sql_bounds(start_dt, end_dt)

    tab1, tab2, tab3, tab4 = st.tabs(
        ["💰 Sales Report", "👥 Employee Ledger", "🏥 Hospital Expense", "🕒 OT Expense"]
    )

    # ---------------- Sales Report ----------------
    with tab1:
        report_df = query_to_dataframe(
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
            st.info("No sales records found for this period.")
        else:
            st.dataframe(report_df, use_container_width=True)
            c1, c2, c3 = st.columns(3)
            c1.metric("Gross Sales", f"PKR {report_df['Total Revenue (PKR)'].sum():,.2f}")
            c2.metric("Cost of Goods Sold", f"PKR {report_df['Total Cost (PKR)'].sum():,.2f}")
            c3.metric("Net Profit", f"PKR {report_df['Profit (PKR)'].sum():,.2f}")

            excel_out = to_excel_bytes({"Sales & Profit Report": report_df}, start_dt, end_dt)
            st.download_button(
                "📥 Download Sales Report (.xlsx)",
                data=excel_out,
                file_name=f"Sales_Report_{start_dt}_to_{end_dt}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="reports_center_sales_download",
            )

    # ---------------- Employee Ledger (debit/credit) ----------------
    with tab2:
        employees =query_to_dataframe("SELECT * FROM employees ORDER BY name", conn)
        if employees.empty:
            st.info("No employees registered yet.")
        else:
            _render_employee_ledger_content(
                conn, start_dt, end_dt, key_prefix="reports_center_emp", show_actions=False
            )

    # ---------------- Hospital Expense ----------------
    with tab3:
        _render_expense_summary_only(
            conn, payment_type="Hospital Expense", start_dt=start_dt, end_dt=end_dt, key_prefix="reports_center_hospital"
        )

    # ---------------- OT Expense ----------------
    with tab4:
        _render_expense_summary_only(
            conn, payment_type="OT Expense", start_dt=start_dt, end_dt=end_dt, key_prefix="reports_center_ot"
        )


def _render_expense_summary_only(
    conn: sqlite3.Connection, payment_type: str, start_dt: date, end_dt: date, key_prefix: str
) -> None:
    """Read-only summary + export for Hospital/OT expenses, for use inside Reports Center
    (no employee-registration UI, since that's not the point of this page)."""
    start_str, end_str = sql_bounds(start_dt, end_dt)

    summary_df = query_to_dataframe(
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

    if summary_df.empty:
        st.info(f"No '{payment_type}' records found for this date range.")
        return

    st.dataframe(summary_df, use_container_width=True)
    st.metric("Total", f"PKR {summary_df['Total Amount (PKR)'].sum():,.2f}")

    detailed_df = query_to_dataframe(
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

    excel_out = to_excel_bytes({"Summary": summary_df, "Detailed Log": detailed_df}, start_dt, end_dt)
    st.download_button(
        f"📥 Download {payment_type} Report (.xlsx)",
        data=excel_out,
        file_name=f"{payment_type.replace(' ', '_')}_Report_{start_dt}_to_{end_dt}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{key_prefix}_download",
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

    shop_user_options = query_to_dataframe(
    "SELECT username FROM users ORDER BY username"
    )

    if not shop_user_options.empty and "username" in shop_user_options.columns:
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
        ("reports_center", "Reports Center"),
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
            module_returns(conn)
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
        elif menu == "Reports Center":
            module_full_reports(conn)
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
