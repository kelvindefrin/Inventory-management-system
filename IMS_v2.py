import csv
import logging
import os
import shutil
import sqlite3
import sys
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font as tkfont

APP_NAME = "Tory"
VERSION = "2.0.0"
DB_NAME = "inventory.db"
TAGLINE = "Track   \u2022   Manage   \u2022   Grow"

FONT = "Segoe UI"
SYM = "Segoe UI Symbol"
ALL_CATS = "All categories"
CHK_ON, CHK_OFF = "\u2611", "\u2610"

SCALE = 1.0  # set at start-up from the screen DPI


def px(n):
    """Scale a pixel size for high-DPI screens."""
    return int(round(n * SCALE))


# --------------------------------------------------------------------------
# Themes (light + dark)
# --------------------------------------------------------------------------
LIGHT = {
    "name": "light",
    "bg": "#f3f6fb", "header": "#ffffff", "surface": "#ffffff", "border": "#dfe5ef",
    "text": "#111827", "muted": "#6b7280", "disabled": "#a3adbd",
    "accent": "#2563eb", "accent_hover": "#1d4ed8", "accent_fg": "#ffffff",
    "btn_bg": "#ffffff", "btn_hover": "#eef2f8", "btn_fg": "#374151", "btn_border": "#d3dae6",
    "danger_bg": "#fff1f1", "danger_hover": "#fddede", "danger_fg": "#dc2626", "danger_border": "#f4b4b4",
    "field": "#ffffff", "field_border": "#d3dae6",
    "sidebar": "#0f1f3d", "side_text": "#c9d4ea", "side_hover": "#1b2f58",
    "head_bg": "#f1f4f9", "head_hover": "#e3e9f3", "row_alt": "#f8fafc", "sel_bg": "#dbe8ff",
    "low_bg": "#fff4d6", "low_fg": "#7a5200", "out_bg": "#fde2e2", "out_fg": "#9b1c1c",
    "good": "#15803d", "bad": "#dc2626", "scroll": "#c5cddb",
    "avatar_bg": "#dbe3f0", "avatar_fg": "#4b5d7d",
    "cards": {
        "items": {"bg": "#eef5ff", "border": "#cddffb", "icon": "#2563eb", "title": "#2563eb"},
        "units": {"bg": "#ecfbf4", "border": "#bdeedb", "icon": "#10b981", "title": "#059669"},
        "value": {"bg": "#fff6ec", "border": "#fbd9b5", "icon": "#f97316", "title": "#ea580c"},
        "low":   {"bg": "#f5f0ff", "border": "#dccbfa", "icon": "#8b5cf6", "title": "#7c3aed"},
    },
}

DARK = {
    "name": "dark",
    "bg": "#0a1426", "header": "#0c1a33", "surface": "#0d1b33", "border": "#1f3560",
    "text": "#eaf0fb", "muted": "#8ea4c8", "disabled": "#5a6f96",
    "accent": "#2563eb", "accent_hover": "#3b78f5", "accent_fg": "#ffffff",
    "btn_bg": "#0d1b33", "btn_hover": "#162c52", "btn_fg": "#dbe6f7", "btn_border": "#2a4373",
    "danger_bg": "#2a1220", "danger_hover": "#3b1828", "danger_fg": "#ff6b7a", "danger_border": "#7a2a3a",
    "field": "#0b172c", "field_border": "#2a4373",
    "sidebar": "#0a1730", "side_text": "#c3d0ea", "side_hover": "#14284c",
    "head_bg": "#102346", "head_hover": "#17305c", "row_alt": "#0f1f3a", "sel_bg": "#1b4a9c",
    "low_bg": "#3a2f10", "low_fg": "#fbd36b", "out_bg": "#3d1520", "out_fg": "#ff8a96",
    "good": "#34d399", "bad": "#ff7b8a", "scroll": "#2a4373",
    "avatar_bg": "#2a4170", "avatar_fg": "#c9d8f3",
    "cards": {
        "items": {"bg": "#0b2a5b", "border": "#1d63d6", "icon": "#1d6fe0", "title": "#7ab6ff"},
        "units": {"bg": "#0a3a37", "border": "#12806f", "icon": "#10b981", "title": "#4ee0b5"},
        "value": {"bg": "#3a2410", "border": "#d9731a", "icon": "#f97316", "title": "#ffab5e"},
        "low":   {"bg": "#251a52", "border": "#6d4bd1", "icon": "#8b5cf6", "title": "#b79cff"},
    },
}

THEMES = {"light": LIGHT, "dark": DARK}
THEME = dict(LIGHT)


def set_theme(name):
    THEME.clear()
    THEME.update(THEMES.get(name, LIGHT))


def mix(c1, c2, t):
    """Blend two #rrggbb colours (t=0 -> c1, t=1 -> c2)."""
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(int(x * (1 - t) + y * t) for x, y in zip(a, b))


# --------------------------------------------------------------------------
# Paths & helpers
# --------------------------------------------------------------------------
def app_dir():
    """Folder of the script, or of the .exe when frozen by PyInstaller."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def data_dir():
    """Per-user data folder (writable even under Program Files)."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "InventoryManager")
    os.makedirs(path, exist_ok=True)
    return path


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_int(text, label, minimum=None):
    s = str(text).strip().replace(",", "")
    try:
        value = int(s)
    except ValueError:
        try:
            f = float(s)
            if f != int(f):
                raise ValueError
            value = int(f)
        except (ValueError, OverflowError):
            raise ValueError(f"{label} must be a whole number.") from None
    if minimum is not None and value < minimum:
        raise ValueError(f"{label} must be at least {minimum}.")
    return value


def parse_float(text, label, minimum=None):
    try:
        value = float(str(text).strip().replace(",", ""))
    except ValueError:
        raise ValueError(f"{label} must be a number.") from None
    if minimum is not None and value < minimum:
        raise ValueError(f"{label} must be at least {minimum}.")
    return value


# --------------------------------------------------------------------------
# Database layer
# --------------------------------------------------------------------------
class Database:
    def __init__(self, path):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self):
        c = self.conn
        c.execute(
            """CREATE TABLE IF NOT EXISTS items (
                   id INTEGER PRIMARY KEY AUTOINCREMENT,
                   name TEXT NOT NULL UNIQUE,
                   quantity INTEGER NOT NULL DEFAULT 0,
                   price REAL NOT NULL DEFAULT 0,
                   min_stock INTEGER NOT NULL DEFAULT 5)"""
        )
        # Upgrade databases created by older versions.
        have = {r["name"] for r in c.execute("PRAGMA table_info(items)")}
        for col in ("sku", "category", "created_at", "updated_at"):
            if col not in have:
                c.execute(f"ALTER TABLE items ADD COLUMN {col} TEXT NOT NULL DEFAULT ''")
        c.execute(
            """CREATE TABLE IF NOT EXISTS movements (
                   id INTEGER PRIMARY KEY AUTOINCREMENT,
                   item_id INTEGER,
                   item_name TEXT NOT NULL,
                   change INTEGER NOT NULL,
                   new_qty INTEGER NOT NULL,
                   reason TEXT NOT NULL DEFAULT '',
                   ts TEXT NOT NULL)"""
        )
        c.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
        c.execute("""CREATE TABLE IF NOT EXISTS orders (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL,
                   total REAL NOT NULL DEFAULT 0, status TEXT NOT NULL DEFAULT 'Completed')""")
        c.execute("""CREATE TABLE IF NOT EXISTS order_lines (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL,
                   item_id INTEGER, item_name TEXT NOT NULL, quantity INTEGER NOT NULL,
                   unit_price REAL NOT NULL, line_total REAL NOT NULL)""")
        c.execute("""CREATE TABLE IF NOT EXISTS expenses (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, description TEXT NOT NULL,
                   amount REAL NOT NULL, ts TEXT NOT NULL)""")
        c.execute("""CREATE TABLE IF NOT EXISTS customers (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
                   phone TEXT NOT NULL, created_at TEXT NOT NULL)""")
        order_cols = {r[1] for r in c.execute("PRAGMA table_info(orders)")}
        if "customer_id" not in order_cols:
            c.execute("ALTER TABLE orders ADD COLUMN customer_id INTEGER")
        c.commit()

    # -- settings ---------------------------------------------------------
    def get_setting(self, key, default=""):
        row = self.conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?,?)", (key, value))
        self.conn.commit()

    # -- queries ----------------------------------------------------------
    def list_items(self, search="", category="", low_only=False):
        sql, args = "SELECT * FROM items WHERE 1=1", []
        if search:
            sql += " AND (name LIKE ? OR sku LIKE ? OR category LIKE ?)"
            args += [f"%{search}%"] * 3
        if category:
            sql += " AND category = ?"
            args.append(category)
        if low_only:
            sql += " AND quantity <= min_stock"
        return self.conn.execute(sql, args).fetchall()

    def get(self, item_id):
        return self.conn.execute("SELECT * FROM items WHERE id=?", (item_id,)).fetchone()

    def categories(self):
        rows = self.conn.execute(
            "SELECT DISTINCT category FROM items WHERE category <> '' ORDER BY category").fetchall()
        return [r[0] for r in rows]

    def stats(self):
        r = self.conn.execute(
            """SELECT COUNT(*), COALESCE(SUM(quantity),0), COALESCE(SUM(quantity*price),0),
                      COALESCE(SUM(CASE WHEN quantity <= min_stock THEN 1 ELSE 0 END),0)
               FROM items""").fetchone()
        return {"items": r[0], "units": r[1], "value": r[2], "low": r[3]}

    def category_summary(self):
        return self.conn.execute(
            """SELECT CASE WHEN category='' THEN 'Uncategorised' ELSE category END AS cat,
                      COUNT(*), COALESCE(SUM(quantity),0), COALESCE(SUM(quantity*price),0),
                      COALESCE(SUM(CASE WHEN quantity <= min_stock THEN 1 ELSE 0 END),0)
               FROM items GROUP BY cat ORDER BY cat COLLATE NOCASE""").fetchall()

    def existing_ids(self):
        return {r[0] for r in self.conn.execute("SELECT id FROM items")}

    def movements(self, item_id=None, limit=1000):
        if item_id is None:
            return self.conn.execute(
                "SELECT * FROM movements ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
        return self.conn.execute(
            "SELECT * FROM movements WHERE item_id=? ORDER BY id DESC LIMIT ?",
            (item_id, limit)).fetchall()

    def checkout(self, cart, customer_id=None):
        """Record a sale and decrement its inventory atomically."""
        if not cart:
            raise ValueError("Add at least one item to the bill.")
        c = self.conn
        try:
            c.execute("BEGIN IMMEDIATE")
            total = 0.0
            lines = []
            for item_id, quantity in cart.items():
                quantity = parse_int(quantity, "Quantity", 1)
                item = c.execute("SELECT * FROM items WHERE id=?", (item_id,)).fetchone()
                if item is None:
                    raise ValueError("An item in this bill no longer exists.")
                if item["quantity"] < quantity:
                    raise ValueError(f"Not enough stock for {item['name']}. Available: {item['quantity']}.")
                line_total = round(quantity * item["price"], 2)
                total += line_total
                lines.append((item, quantity, line_total))
            if customer_id is not None and c.execute("SELECT 1 FROM customers WHERE id=?", (customer_id,)).fetchone() is None:
                raise ValueError("The selected customer no longer exists.")
            cur = c.execute("INSERT INTO orders (ts,total,status,customer_id) VALUES (?,?,?,?)",
                            (now(), round(total, 2), "Completed", customer_id))
            order_id = cur.lastrowid
            for item, quantity, line_total in lines:
                new_qty = item["quantity"] - quantity
                c.execute("UPDATE items SET quantity=?, updated_at=? WHERE id=?", (new_qty, now(), item["id"]))
                c.execute("INSERT INTO order_lines (order_id,item_id,item_name,quantity,unit_price,line_total) VALUES (?,?,?,?,?,?)",
                          (order_id, item["id"], item["name"], quantity, item["price"], line_total))
                self._log(item["id"], item["name"], -quantity, new_qty, f"Sale #{order_id}")
            c.commit()
            return order_id, round(total, 2)
        except Exception:
            c.rollback()
            raise

    def create_customer(self, name, phone):
        cur = self.conn.execute("INSERT INTO customers (name,phone,created_at) VALUES (?,?,?)",
                                (name, phone, now()))
        self.conn.commit()
        return cur.lastrowid

    def customers(self, search=""):
        sql = """SELECT c.id,c.name,c.phone,c.created_at,COUNT(o.id) AS orders,
                 COALESCE(SUM(CASE WHEN o.status='Completed' THEN o.total ELSE 0 END),0) AS spent,
                 MAX(o.ts) AS last_order
                 FROM customers c LEFT JOIN orders o ON o.customer_id=c.id"""
        args = []
        if search:
            sql += " WHERE c.name LIKE ? OR c.phone LIKE ? OR CAST(c.id AS TEXT) LIKE ?"
            args = [f"%{search}%"] * 3
        sql += " GROUP BY c.id ORDER BY c.id DESC"
        return self.conn.execute(sql, args).fetchall()

    def get_customer(self, customer_id):
        return self.conn.execute("SELECT * FROM customers WHERE id=?", (customer_id,)).fetchone()

    def add_expense(self, description, amount):
        self.conn.execute("INSERT INTO expenses (description,amount,ts) VALUES (?,?,?)",
                          (description, amount, now()))
        self.conn.commit()

    def dashboard_data(self):
        c = self.conn
        order_count, income = c.execute("SELECT COUNT(*),COALESCE(SUM(total),0) FROM orders WHERE status='Completed'").fetchone()
        expense = c.execute("SELECT COALESCE(SUM(amount),0) FROM expenses").fetchone()[0]
        sold = c.execute("SELECT COALESCE(SUM(quantity),0) FROM order_lines").fetchone()[0]
        received = c.execute("SELECT COALESCE(SUM(CASE WHEN change>0 THEN change ELSE 0 END),0),COALESCE(SUM(CASE WHEN change<0 THEN -change ELSE 0 END),0) FROM movements").fetchone()
        low = c.execute("SELECT name,quantity,min_stock FROM items WHERE quantity<=min_stock ORDER BY quantity,min_stock DESC LIMIT 8").fetchall()
        category_sales = c.execute("SELECT COALESCE(NULLIF(i.category,''),'Uncategorised'),SUM(ol.line_total) FROM order_lines ol LEFT JOIN items i ON i.id=ol.item_id GROUP BY 1 ORDER BY 2 DESC LIMIT 6").fetchall()
        return {"orders": order_count, "income": income, "expense": expense, "sold": sold,
                "stock_in": received[0], "stock_out": received[1], "low": low,
                "category_sales": category_sales}

    # -- changes ----------------------------------------------------------
    def _log(self, item_id, name, change, new_qty, reason):
        self.conn.execute(
            "INSERT INTO movements (item_id, item_name, change, new_qty, reason, ts) VALUES (?,?,?,?,?,?)",
            (item_id, name, change, new_qty, reason, now()))

    def _check_name(self, name, item_id=0):
        dup = self.conn.execute(
            "SELECT id FROM items WHERE lower(name)=lower(?) AND id<>?", (name, item_id)).fetchone()
        if dup:
            raise ValueError("An item with that name already exists.")

    def add_item(self, d):
        self._check_name(d["name"])
        ts = now()
        cur = self.conn.execute(
            """INSERT INTO items (sku, name, category, quantity, price, min_stock, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?)""",
            (d["sku"], d["name"], d["category"], d["quantity"], d["price"], d["min_stock"], ts, ts))
        if d["quantity"]:
            self._log(cur.lastrowid, d["name"], d["quantity"], d["quantity"], "Opening stock")
        self.conn.commit()
        return cur.lastrowid

    def update_item(self, item_id, d):
        old = self.get(item_id)
        if old is None:
            raise ValueError("Item no longer exists.")
        self._check_name(d["name"], item_id)
        self.conn.execute(
            """UPDATE items SET sku=?, name=?, category=?, quantity=?, price=?, min_stock=?, updated_at=?
               WHERE id=?""",
            (d["sku"], d["name"], d["category"], d["quantity"], d["price"], d["min_stock"], now(), item_id))
        if d["quantity"] != old["quantity"]:
            self._log(item_id, d["name"], d["quantity"] - old["quantity"], d["quantity"], "Manual edit")
        self.conn.commit()

    def adjust_stock(self, item_id, change, reason):
        row = self.get(item_id)
        if row is None:
            raise ValueError("Item no longer exists.")
        new_qty = row["quantity"] + change
        if new_qty < 0:
            raise ValueError(f"Not enough stock. Only {row['quantity']} available.")
        self.conn.execute("UPDATE items SET quantity=?, updated_at=? WHERE id=?",
                          (new_qty, now(), item_id))
        self._log(item_id, row["name"], change, new_qty, reason)
        self.conn.commit()
        return new_qty

    def delete_item(self, item_id):
        row = self.get(item_id)
        if row is None:
            return
        self._log(item_id, row["name"], -row["quantity"], 0, "Item deleted")
        self.conn.execute("DELETE FROM items WHERE id=?", (item_id,))
        self.conn.commit()

    def backup(self, dest):
        target = sqlite3.connect(dest)
        try:
            self.conn.backup(target)
        finally:
            target.close()

    def restore(self, src_path):
        src = sqlite3.connect(src_path)
        try:
            src.execute("SELECT name, quantity, price FROM items LIMIT 1")  # validates the file
            src.backup(self.conn)
        finally:
            src.close()
        self._init_schema()


# --------------------------------------------------------------------------
# Custom widgets
# --------------------------------------------------------------------------
def round_rect(cv, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2,
           x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


class FlatButton(tk.Label):
    """A flat, theme-aware button with hover and disabled states."""

    def __init__(self, parent, text="", command=None, kind="secondary", icon="", padx=14, pady=7,
                 font=None):
        self.command = command
        self.enabled = True
        self.kind = kind
        self.pal = self._palette(kind)
        label = f"{icon}  {text}" if icon and text else (icon or text)
        super().__init__(parent, text=label, font=font or (FONT, 10), padx=px(padx), pady=px(pady),
                         cursor="hand2", bd=0, highlightthickness=1)
        self._apply()
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.bind("<Button-1>", self._click)

    @staticmethod
    def _palette(kind):
        T = THEME
        if kind == "primary":
            return dict(bg=T["accent"], fg=T["accent_fg"], border=T["accent"], hover=T["accent_hover"])
        if kind == "danger":
            return dict(bg=T["danger_bg"], fg=T["danger_fg"], border=T["danger_border"], hover=T["danger_hover"])
        return dict(bg=T["btn_bg"], fg=T["btn_fg"], border=T["btn_border"], hover=T["btn_hover"])

    def _apply(self):
        p = self.pal
        fg = p["fg"] if self.enabled else THEME["disabled"]
        self.configure(bg=p["bg"], fg=fg, highlightbackground=p["border"], highlightcolor=p["border"],
                       cursor="hand2" if self.enabled else "arrow")

    def set_kind(self, kind):
        self.kind = kind
        self.pal = self._palette(kind)
        self._apply()

    def set_enabled(self, flag):
        self.enabled = bool(flag)
        self._apply()

    def _enter(self, _e):
        try:
            if self.enabled:
                self.configure(bg=self.pal["hover"])
        except tk.TclError:
            pass

    def _leave(self, _e):
        try:
            self.configure(bg=self.pal["bg"])
        except tk.TclError:
            pass

    def _click(self, _e):
        if self.enabled and self.command:
            self.command()


class IconBadge(tk.Canvas):
    """Round coloured icon, as used on the dashboard cards."""

    def __init__(self, parent, size, color, glyph, bg, fg="#ffffff", font_size=16):
        s = px(size)
        super().__init__(parent, width=s, height=s, bg=bg, highlightthickness=0, bd=0)
        self.create_oval(1, 1, s - 1, s - 1, fill=mix(color, bg, 0.72), outline="")
        pad = int(s * 0.13)
        self.create_oval(pad, pad, s - pad, s - pad, fill=color, outline="")
        self.create_text(s / 2, s / 2, text=glyph, fill=fg, font=(SYM, font_size, "bold"))


class SearchBox(tk.Frame):
    """Search field with an icon and placeholder text."""

    def __init__(self, parent, placeholder, on_change, text=""):
        T = THEME
        super().__init__(parent, bg=T["field"], highlightthickness=1,
                         highlightbackground=T["field_border"], highlightcolor=T["accent"])
        self.placeholder, self.on_change, self._ph, self._last = placeholder, on_change, False, None
        tk.Label(self, text="\u2315", font=(SYM, 14), bg=T["field"], fg=T["muted"]).pack(
            side="left", padx=(px(10), px(4)))
        self.entry = tk.Entry(self, relief="flat", bd=0, bg=T["field"], fg=T["text"],
                              insertbackground=T["text"], font=(FONT, 10), width=26, highlightthickness=0)
        self.entry.pack(side="left", fill="both", expand=True, ipady=px(7), padx=(0, px(10)))
        self.entry.bind("<FocusIn>", self._focus_in)
        self.entry.bind("<FocusOut>", self._focus_out)
        self.entry.bind("<KeyRelease>", self._changed)
        if text:
            self.entry.insert(0, text)
        else:
            self._show_placeholder()

    def _show_placeholder(self):
        self.entry.delete(0, "end")
        self.entry.insert(0, self.placeholder)
        self.entry.configure(fg=THEME["muted"])
        self._ph = True

    def _focus_in(self, _e):
        if self._ph:
            self.entry.delete(0, "end")
            self.entry.configure(fg=THEME["text"])
            self._ph = False

    def _focus_out(self, _e):
        if not self.entry.get().strip():
            self._show_placeholder()

    def _changed(self, _e):
        value = self.get()
        if value != self._last:
            self._last = value
            self.on_change()

    def get(self):
        return "" if self._ph else self.entry.get().strip()

    def set(self, text):
        self._focus_in(None)
        self.entry.delete(0, "end")
        self.entry.insert(0, text)
        if not text:
            self._show_placeholder()
        self._last = self.get()

    def focus_entry(self):
        self.entry.focus_set()
        self._focus_in(None)
        self.entry.select_range(0, "end")


# --------------------------------------------------------------------------
# Dialogs
# --------------------------------------------------------------------------
class Dialog(tk.Toplevel):
    """Modal dialog base. Subclasses build into self.body, then call self.show()."""

    def __init__(self, parent, title):
        super().__init__(parent, bg=THEME["surface"])
        self.withdraw()
        self.parent = parent
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.result = None
        self.body = ttk.Frame(self, padding=22, style="Card.TFrame")
        self.body.pack(fill="both", expand=True)
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Escape>", lambda e: self.cancel())

    def buttons(self, ok_text="Save", row=0):
        bar = ttk.Frame(self.body, style="Card.TFrame")
        bar.grid(row=row, column=0, columnspan=2, sticky="e", pady=(18, 0))
        FlatButton(bar, "Cancel", self.cancel).pack(side="right")
        FlatButton(bar, ok_text, self.ok, "primary", padx=18).pack(side="right", padx=(0, 8))
        self.bind("<Return>", lambda e: self.ok())

    def show(self):
        self.update_idletasks()
        p = self.parent
        x = p.winfo_rootx() + (p.winfo_width() - self.winfo_reqwidth()) // 2
        y = p.winfo_rooty() + (p.winfo_height() - self.winfo_reqheight()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.deiconify()
        self.grab_set()
        self.focus_force()
        self.wait_window()
        return self.result

    def cancel(self):
        self.result = None
        self.destroy()

    def ok(self):
        raise NotImplementedError


class ItemDialog(Dialog):
    def __init__(self, parent, categories, default_min, item=None):
        super().__init__(parent, "Edit Item" if item else "Add Item")
        b = self.body
        self.vars = {k: tk.StringVar() for k in ("name", "sku", "category", "quantity", "price", "min_stock")}
        if item:
            for k in self.vars:
                self.vars[k].set(item[k])
        else:
            self.vars["quantity"].set("0")
            self.vars["price"].set("0")
            self.vars["min_stock"].set(str(default_min))

        ttk.Label(b, text="Edit item" if item else "Add a new item", style="Card.Heading.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        fields = [("Item name *", "name"), ("SKU / Code", "sku"), ("Category", "category"),
                  ("Quantity", "quantity"), ("Price per unit", "price"), ("Low-stock alert level", "min_stock")]
        first = None
        for i, (label, key) in enumerate(fields, start=1):
            ttk.Label(b, text=label, style="Card.TLabel").grid(row=i, column=0, sticky="w", pady=5, padx=(0, 16))
            if key == "category":
                w = ttk.Combobox(b, textvariable=self.vars[key], values=categories, width=28)
            else:
                w = ttk.Entry(b, textvariable=self.vars[key], width=30)
            w.grid(row=i, column=1, sticky="ew", pady=5)
            first = first or w
        ttk.Label(b, text="You are warned when quantity is at or below the alert level.",
                  style="Card.Muted.TLabel").grid(row=len(fields) + 1, column=0, columnspan=2, sticky="w",
                                                  pady=(6, 0))
        self.buttons("Save", row=len(fields) + 2)
        first.focus_set()

    def ok(self):
        v = {k: s.get().strip() for k, s in self.vars.items()}
        try:
            if not v["name"]:
                raise ValueError("Item name cannot be empty.")
            data = {
                "name": v["name"], "sku": v["sku"], "category": v["category"],
                "quantity": parse_int(v["quantity"], "Quantity", 0),
                "price": parse_float(v["price"], "Price", 0),
                "min_stock": parse_int(v["min_stock"], "Low-stock alert level", 0),
            }
        except ValueError as e:
            messagebox.showwarning("Check your input", str(e), parent=self)
            return
        self.result = data
        self.destroy()


class StockDialog(Dialog):
    REASONS_IN = ["Purchase", "Return from customer", "Stock count correction", "Other"]
    REASONS_OUT = ["Sale", "Damaged", "Expired", "Lost / stolen", "Stock count correction", "Other"]

    def __init__(self, parent, item):
        super().__init__(parent, "Adjust Stock")
        self.item = item
        b = self.body
        ttk.Label(b, text=item["name"], style="Card.Heading.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(b, text=f"Current quantity: {item['quantity']}", style="Card.Muted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self.mode = tk.StringVar(value="in")
        modes = ttk.Frame(b, style="Card.TFrame")
        modes.grid(row=2, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Radiobutton(modes, text="Stock in (+)", variable=self.mode, value="in", style="Card.TRadiobutton",
                        command=self._mode_changed).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(modes, text="Stock out (-)", variable=self.mode, value="out", style="Card.TRadiobutton",
                        command=self._mode_changed).pack(side="left")

        ttk.Label(b, text="Quantity", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=6, padx=(0, 16))
        self.qty = tk.StringVar(value="1")
        self.qty.trace_add("write", lambda *a: self._preview())
        qty_entry = ttk.Entry(b, textvariable=self.qty, width=12)
        qty_entry.grid(row=3, column=1, sticky="w", pady=6)

        ttk.Label(b, text="Reason", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=6)
        self.reason = tk.StringVar(value=self.REASONS_IN[0])
        self.reason_box = ttk.Combobox(b, textvariable=self.reason, values=self.REASONS_IN, width=26)
        self.reason_box.grid(row=4, column=1, sticky="w", pady=6)

        self.preview = ttk.Label(b, text="", style="Card.Muted.TLabel")
        self.preview.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self._preview()
        self.buttons("Apply", row=6)
        qty_entry.focus_set()
        qty_entry.select_range(0, "end")

    def _mode_changed(self):
        reasons = self.REASONS_IN if self.mode.get() == "in" else self.REASONS_OUT
        self.reason_box.configure(values=reasons)
        self.reason.set(reasons[0])
        self._preview()

    def _delta(self):
        n = parse_int(self.qty.get(), "Quantity", 1)
        return n if self.mode.get() == "in" else -n

    def _preview(self):
        try:
            new = self.item["quantity"] + self._delta()
            self.preview.configure(text=f"New quantity will be: {new}" if new >= 0 else "Not enough stock.")
        except ValueError:
            self.preview.configure(text="")

    def ok(self):
        try:
            delta = self._delta()
            if self.item["quantity"] + delta < 0:
                raise ValueError(f"Not enough stock. Only {self.item['quantity']} available.")
        except ValueError as e:
            messagebox.showwarning("Check your input", str(e), parent=self)
            return
        self.result = (delta, self.reason.get().strip() or "Other")
        self.destroy()


class HistoryDialog(Dialog):
    def __init__(self, parent, rows, title):
        super().__init__(parent, title)
        self.resizable(True, True)
        b = self.body
        cols = ("ts", "item", "change", "qty", "reason")
        tree = ttk.Treeview(b, columns=cols, show="headings", height=16)
        for col, text, width, anchor in (("ts", "Date / time", 150, "w"), ("item", "Item", 220, "w"),
                                         ("change", "Change", 70, "e"), ("qty", "New qty", 80, "e"),
                                         ("reason", "Reason", 180, "w")):
            tree.heading(col, text=text)
            tree.column(col, width=px(width), anchor=anchor)
        tree.tag_configure("in", foreground=THEME["good"])
        tree.tag_configure("out", foreground=THEME["bad"])
        for r in rows:
            tree.insert("", "end", values=(r["ts"], r["item_name"], f"{r['change']:+d}", r["new_qty"], r["reason"]),
                        tags=("in" if r["change"] >= 0 else "out",))
        sb = ttk.Scrollbar(b, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        b.rowconfigure(0, weight=1)
        b.columnconfigure(0, weight=1)
        if not rows:
            ttk.Label(b, text="No movements recorded yet.", style="Card.Muted.TLabel").grid(row=1, column=0, pady=8)
        FlatButton(b, "Close", self.cancel).grid(row=2, column=0, columnspan=2, sticky="e", pady=(12, 0))


# --------------------------------------------------------------------------
# Main window
# --------------------------------------------------------------------------
class App(tk.Tk):
    NAV = (("dashboard", "Dashboard", "\u2302"), ("billing", "Billing", "\u25A4"),
           ("customers", "Customers", "\u263A"), ("items", "Items", "\u25A3"),
           ("reports", "Reports", "\u25A5"), ("tools", "Tools", "\u2692"),
           ("settings", "Settings", "\u2699"), ("help", "Help", "?"))

    COLUMNS = (  # id, heading, width, anchor
        ("check", CHK_OFF, 44, "center"), ("sku", "SKU", 110, "w"), ("name", "Item name", 250, "w"),
        ("category", "Category", 150, "center"), ("quantity", "Qty", 80, "center"),
        ("price", "Price", 110, "center"), ("value", "Stock value", 130, "center"),
        ("min_stock", "Alert level", 100, "center"), ("status", "Status", 90, "center"),
    )
    NUMERIC = {"quantity", "price", "value", "min_stock"}
    REPORTS = (("low", "Low-stock report"), ("moves", "Stock movements"), ("category", "Stock by category"))

    def __init__(self, db):
        super().__init__()
        global SCALE
        try:
            SCALE = max(1.0, float(self.tk.call("tk", "scaling")) / (96 / 72))
        except (tk.TclError, ValueError):
            SCALE = 1.0
        self.db = db
        self.currency = db.get_setting("currency", "Rs. ")
        self.customer_required = db.get_setting("customer_required", "0") == "1"
        try:
            self.default_min = int(db.get_setting("default_min", "5"))
        except ValueError:
            self.default_min = 5
        self.user_name = db.get_setting("user_name", "Admin") or "Admin"
        self.theme_name = db.get_setting("theme", "light")
        if self.theme_name not in THEMES:
            self.theme_name = "light"
        set_theme(self.theme_name)

        self.sort_col, self.sort_rev = "name", False
        self.search_text = ""
        self.cat_var = tk.StringVar(value=ALL_CATS)
        self.low_only = tk.BooleanVar(value=False)
        self.checked = set()
        self.page = "dashboard"
        self.bill_cart = {}
        self.bill_selected_customer = None
        self.report_kind = "low"
        self.tree = self.rep_tree = self.search_box = None
        self.bill_customer_box = self.customer_tree = None
        self.card_vars = None
        self._page_refresh = None
        self._row_tags, self._sel_prev = {}, None

        self.title(APP_NAME)
        w, h = px(1280), px(780)
        w, h = min(w, self.winfo_screenwidth() - 40), min(h, self.winfo_screenheight() - 80)
        self.geometry(f"{w}x{h}")
        self.minsize(min(px(1020), w), min(px(600), h))
        self._init_fonts()
        self._style()
        self._build_shell()
        self._bind_keys()
        self.report_callback_exception = self._on_error
        self.refresh()

    # ----------------------------------------------------------------------
    # Look & feel
    # ----------------------------------------------------------------------
    def _init_fonts(self):
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(name).configure(family=FONT, size=10)
            except tk.TclError:
                pass

    def _style(self):
        T = THEME
        s = ttk.Style(self)
        s.theme_use("clam")  # fully skinnable look on every Windows version
        self.configure(bg=T["bg"])
        s.configure(".", background=T["bg"], foreground=T["text"], font=(FONT, 10))

        for p, b in (("", T["bg"]), ("Card.", T["surface"])):
            s.configure(f"{p}TFrame", background=b)
            s.configure(f"{p}TLabel", background=b, foreground=T["text"])
            s.configure(f"{p}Muted.TLabel", background=b, foreground=T["muted"])
            s.configure(f"{p}Heading.TLabel", background=b, foreground=T["text"], font=(FONT, 12, "bold"))
            for kind in ("TCheckbutton", "TRadiobutton"):
                s.configure(f"{p}{kind}", background=b, foreground=T["text"], focuscolor=b,
                            indicatorbackground=T["field"], indicatorforeground="#ffffff",
                            upperbordercolor=T["field_border"], lowerbordercolor=T["field_border"],
                            indicatormargin=(0, 0, 6, 0))
                s.map(f"{p}{kind}", background=[("active", b)], foreground=[("active", T["text"])],
                      indicatorbackground=[("selected", T["accent"]), ("pressed", T["accent"])],
                      upperbordercolor=[("selected", T["accent"])], lowerbordercolor=[("selected", T["accent"])])

        for w in ("TEntry", "TCombobox"):
            s.configure(w, padding=6, fieldbackground=T["field"], foreground=T["text"],
                        bordercolor=T["field_border"], lightcolor=T["field_border"], darkcolor=T["field_border"],
                        insertcolor=T["text"], selectbackground=T["accent"], selectforeground="#ffffff")
            s.map(w, bordercolor=[("focus", T["accent"])], lightcolor=[("focus", T["accent"])],
                  darkcolor=[("focus", T["accent"])])
        s.configure("TCombobox", background=T["btn_bg"], arrowcolor=T["muted"], arrowsize=px(14))
        s.map("TCombobox", fieldbackground=[("readonly", T["field"]), ("disabled", T["field"])],
              foreground=[("readonly", T["text"])], selectbackground=[("readonly", T["field"])],
              selectforeground=[("readonly", T["text"])], background=[("active", T["btn_hover"])])
        self.option_add("*TCombobox*Listbox.background", T["field"])
        self.option_add("*TCombobox*Listbox.foreground", T["text"])
        self.option_add("*TCombobox*Listbox.selectBackground", T["accent"])
        self.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        self.option_add("*TCombobox*Listbox.font", (FONT, 10))

        for orient in ("Vertical", "Horizontal"):
            s.configure(f"{orient}.TScrollbar", background=T["scroll"], troughcolor=T["surface"],
                        bordercolor=T["surface"], arrowcolor=T["muted"], lightcolor=T["scroll"],
                        darkcolor=T["scroll"], relief="flat")
            s.map(f"{orient}.TScrollbar", background=[("active", T["muted"])])

        s.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        s.configure("Treeview", rowheight=px(34), background=T["surface"], fieldbackground=T["surface"],
                    foreground=T["text"], borderwidth=0, relief="flat", font=(FONT, 10))
        s.configure("Treeview.Heading", font=(FONT, 10, "bold"), background=T["head_bg"], foreground=T["text"],
                    padding=(8, 9), relief="flat", borderwidth=0, bordercolor=T["head_bg"],
                    lightcolor=T["head_bg"], darkcolor=T["head_bg"])
        s.map("Treeview.Heading", background=[("active", T["head_hover"])])
        s.map("Treeview", background=[("selected", T["sel_bg"])], foreground=[("selected", T["text"])])

        # Work around the Tk 8.6.9 bug (bundled with Python 3.8) where Treeview
        # row tag colours are ignored.
        def fixed_map(option):
            return [e for e in s.map("Treeview", query_opt=option) if e[:2] != ("!disabled", "!selected")]
        s.map("Treeview", foreground=fixed_map("foreground"), background=fixed_map("background"))

    def _mk_menu(self):
        T = THEME
        return tk.Menu(self, tearoff=0, bg=T["surface"], fg=T["text"], activebackground=T["accent"],
                       activeforeground="#ffffff", bd=1, relief="flat", font=(FONT, 10))

    # ----------------------------------------------------------------------
    # Shell: header + sidebar + content area
    # ----------------------------------------------------------------------
    def _build_shell(self):
        T = THEME
        for w in self.winfo_children():
            w.destroy()
        self.tree = self.rep_tree = self.search_box = None
        self.configure(bg=T["bg"])
        self._build_menus()
        self._build_header()
        tk.Frame(self, height=1, bg=T["border"]).pack(fill="x")
        body = tk.Frame(self, bg=T["bg"])
        body.pack(fill="both", expand=True)
        self._build_sidebar(body)
        self.content = tk.Frame(body, bg=T["bg"])
        self.content.pack(side="left", fill="both", expand=True, padx=px(18), pady=(px(16), px(10)))
        self.show_page(self.page)

    def _build_menus(self):
        self.admin_menu = self._mk_menu()
        other = "dark" if self.theme_name == "light" else "light"
        self.admin_menu.add_command(label=f"Switch to {other} theme", accelerator="Ctrl+T", command=self.toggle_theme)
        self.admin_menu.add_command(label="Settings", command=lambda: self.show_page("settings"))
        self.admin_menu.add_command(label="About", command=self.about)
        self.admin_menu.add_separator()
        self.admin_menu.add_command(label="Exit", command=self.destroy)

        self.context = self._mk_menu()
        self.context.add_command(label="Edit item...", command=self.edit_item)
        self.context.add_command(label="Adjust stock...", command=self.adjust_stock)
        self.context.add_command(label="View history", command=self.show_history)
        self.context.add_separator()
        self.context.add_command(label="Delete item", command=self.delete_item)

    def _build_header(self):
        T = THEME
        h = tk.Frame(self, bg=T["header"], height=px(60))
        h.pack(fill="x")
        h.pack_propagate(False)

        logo = tk.Canvas(h, width=px(38), height=px(38), bg=T["header"], highlightthickness=0)
        round_rect(logo, 1, 1, px(38) - 1, px(38) - 1, px(10), fill=T["accent"], outline="")
        logo.create_text(px(19), px(19), text="\u25A3", fill="#ffffff", font=(SYM, 17, "bold"))
        logo.pack(side="left", padx=(px(16), px(12)))
        tk.Label(h, text=APP_NAME, font=(FONT, 15, "bold"), bg=T["header"], fg=T["text"]).pack(side="left")
        tk.Label(h, text=TAGLINE, font=(FONT, 9, "italic"), bg=T["header"], fg=T["muted"]).pack(
            side="left", padx=(px(18), 0), pady=(px(4), 0))

        # Right side: user menu, notifications, theme switch
        user = tk.Frame(h, bg=T["header"], cursor="hand2")
        user.pack(side="right", padx=(px(6), px(18)))
        s = px(34)
        av = tk.Canvas(user, width=s, height=s, bg=T["header"], highlightthickness=0, cursor="hand2")
        av.create_oval(1, 1, s - 1, s - 1, fill=T["avatar_bg"], outline="")
        av.create_oval(s * .36, s * .2, s * .64, s * .48, fill=T["avatar_fg"], outline="")
        av.create_arc(s * .24, s * .52, s * .76, s * 1.0, start=0, extent=180, fill=T["avatar_fg"], outline="")
        av.pack(side="left")
        lab = tk.Label(user, text=f"{self.user_name}  \u25BE", font=(FONT, 10), bg=T["header"], fg=T["text"],
                       cursor="hand2")
        lab.pack(side="left", padx=(px(8), 0))
        for w in (user, av, lab):
            w.bind("<Button-1>", lambda e, u=user: self.admin_menu.tk_popup(u.winfo_rootx(),
                                                                             u.winfo_rooty() + u.winfo_height()))

        bw, bh = px(42), px(36)
        self.bell = tk.Canvas(h, width=bw, height=bh, bg=T["header"], highlightthickness=0, cursor="hand2")
        self.bell.create_text(bw * .45, bh * .55, text="\U0001F514", fill=T["muted"], font=(SYM, 15))
        self.bell_dot = self.bell.create_oval(bw - px(17), px(2), bw - px(3), px(16), fill="#ef4444",
                                              outline=T["header"], width=1, state="hidden")
        self.bell_num = self.bell.create_text(bw - px(10), px(9), text="", fill="#ffffff",
                                              font=(FONT, 7, "bold"), state="hidden")
        self.bell.bind("<Button-1>", lambda e: self.bell_click())
        self.bell.pack(side="right", padx=(0, px(6)))

        icon = "\u2600" if self.theme_name == "dark" else "\u263E"
        FlatButton(h, icon=icon, command=self.toggle_theme, padx=10, pady=4, font=(SYM, 12)).pack(
            side="right", padx=(0, px(10)))

    def _build_sidebar(self, parent):
        T = THEME
        sb = tk.Frame(parent, bg=T["sidebar"], width=px(178))
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        tk.Frame(sb, height=px(12), bg=T["sidebar"]).pack()
        self.nav = {}
        for key, text, glyph in self.NAV:
            row = tk.Frame(sb, bg=T["sidebar"], cursor="hand2")
            row.pack(fill="x", padx=px(10), pady=px(3))
            ic = tk.Label(row, text=glyph, font=(SYM, 14), width=2, bg=T["sidebar"], fg=T["side_text"],
                          cursor="hand2")
            ic.pack(side="left", padx=(px(10), px(4)), pady=px(9))
            tx = tk.Label(row, text=text, font=(FONT, 10), bg=T["sidebar"], fg=T["side_text"],
                          cursor="hand2", anchor="w")
            tx.pack(side="left", fill="x", expand=True)
            self.nav[key] = (row, ic, tx)
            for w in (row, ic, tx):
                w.bind("<Button-1>", lambda e, k=key: self.show_page(k))
                w.bind("<Enter>", lambda e, k=key: self._paint_nav(k, True))
                w.bind("<Leave>", lambda e, k=key: self._paint_nav(k, False))
        tk.Label(sb, text=f"Version {VERSION}", font=(FONT, 8), bg=T["sidebar"], fg=T["muted"]).pack(
            side="bottom", pady=px(12))

    def _paint_nav(self, key, hover=False):
        T = THEME
        row, ic, tx = self.nav[key]
        active = key == self.page
        bg = T["accent"] if active else (T["side_hover"] if hover else T["sidebar"])
        fg = "#ffffff" if active else T["side_text"]
        try:
            row.configure(bg=bg)
            ic.configure(bg=bg, fg=fg)
            tx.configure(bg=bg, fg=fg, font=(FONT, 10, "bold") if active else (FONT, 10))
        except tk.TclError:
            pass

    def show_page(self, key):
        if self.page == "billing" and getattr(self, "bill_customer_box", None):
            self.bill_selected_customer = self.bill_customer_ids.get(self.bill_customer.get())
        self.page = key
        for w in self.content.winfo_children():
            w.destroy()
        self.tree = self.rep_tree = self.search_box = None
        self.card_vars = None
        self._page_refresh = None
        self._sel_prev = None
        for k in self.nav:
            self._paint_nav(k)
        getattr(self, "_page_" + key)(self.content)
        self.refresh()

    def _title(self, parent, title, subtitle=""):
        T = THEME
        f = tk.Frame(parent, bg=T["bg"])
        f.pack(fill="x", pady=(0, px(14)))
        tk.Label(f, text=title, font=(FONT, 17, "bold"), bg=T["bg"], fg=T["text"]).pack(anchor="w")
        if subtitle:
            tk.Label(f, text=subtitle, font=(FONT, 10), bg=T["bg"], fg=T["muted"]).pack(anchor="w")

    def _card(self, parent):
        T = THEME
        return tk.Frame(parent, bg=T["surface"], highlightthickness=1, highlightbackground=T["border"])

    # ----------------------------------------------------------------------
    # Inventory pages (Dashboard + Items)
    # ----------------------------------------------------------------------
    def _page_dashboard(self, p):
        self._title(p, "Dashboard", "Sales, stock flow, income and low-stock overview.")
        self._build_cards(p)
        holder = tk.Frame(p, bg=THEME["bg"])
        holder.pack(fill="both", expand=True, pady=(px(14), 0))
        self.dash_canvases = []
        specs = (("Order statistics", "pie"), ("Stock in & out", "stock"),
                 ("Income & expenses", "finance"), ("Stock balance & low stock", "balance"))
        for idx, (title, kind) in enumerate(specs):
            card = self._card(holder)
            card.grid(row=idx // 2, column=idx % 2, sticky="nsew",
                      padx=(0 if idx % 2 == 0 else px(8), px(8) if idx % 2 == 0 else 0),
                      pady=(0, px(8)))
            holder.rowconfigure(idx // 2, weight=1, uniform="dashrow")
            holder.columnconfigure(idx % 2, weight=1, uniform="dashcol")
            tk.Label(card, text=title, font=(FONT, 11, "bold"), bg=THEME["surface"], fg=THEME["text"]).pack(anchor="w", padx=px(14), pady=(px(10), 0))
            canvas = tk.Canvas(card, height=px(145), bg=THEME["surface"], highlightthickness=0)
            canvas.pack(fill="both", expand=True, padx=px(8), pady=px(4))
            self.dash_canvases.append((canvas, kind))
        self._page_refresh = self._refresh_dashboard

    def _refresh_dashboard(self):
        if not self.card_vars:
            return
        st = self.db.stats()
        data = self.db.dashboard_data()
        self.card_vars["items"].set(f"{st['items']:,}")
        self.card_vars["units"].set(f"{st['units']:,}")
        self.card_vars["value"].set(self.money(st["value"]))
        self.card_vars["low"].set(f"{st['low']:,}")
        for canvas, kind in getattr(self, "dash_canvases", []):
            canvas.delete("all")
            canvas.update_idletasks()
            w, h = max(canvas.winfo_width(), 220), max(canvas.winfo_height(), 110)
            muted, text = THEME["muted"], THEME["text"]
            if kind == "pie":
                rows = data["category_sales"]
                if not rows:
                    canvas.create_text(w/2, h/2, text="No completed orders yet", fill=muted, font=(FONT, 10))
                    continue
                total = sum(float(r[1]) for r in rows) or 1
                colors = [THEME["accent"], "#10b981", "#f97316", "#8b5cf6", "#e11d48", "#06b6d4"]
                start = 90
                for i, row in enumerate(rows):
                    extent = 360 * float(row[1]) / total
                    canvas.create_arc(12, 12, min(h-6, 125), h-6, start=start, extent=-extent,
                                      fill=colors[i % len(colors)], outline=THEME["surface"], width=2, style="pieslice")
                    start -= extent
                    canvas.create_rectangle(145, 14+i*20, 157, 26+i*20, fill=colors[i % len(colors)], outline="")
                    canvas.create_text(164, 20+i*20, anchor="w", text=f"{row[0]}  {self.money(row[1])}", fill=text, font=(FONT, 9))
                canvas.create_text(w-8, h-10, anchor="se", text=f"{data['orders']} completed orders", fill=muted, font=(FONT, 8))
            elif kind == "stock":
                vals = (data["stock_in"], data["stock_out"])
                colors = (THEME["good"], THEME["bad"])
                labels = ("Stock in", "Stock out")
                maxv = max(vals) or 1
                for i, (v, color, label) in enumerate(zip(vals, colors, labels)):
                    y = 28 + i*48
                    canvas.create_text(8, y+10, anchor="w", text=label, fill=text, font=(FONT, 9))
                    x0 = 82; x1 = x0 + max(2, (w-135)*v/maxv)
                    canvas.create_rectangle(x0, y, x1, y+20, fill=color, outline="")
                    canvas.create_text(min(w-25, x1+22), y+10, text=str(v), fill=text, font=(FONT, 9, "bold"))
                canvas.create_text(w-8, h-8, anchor="se", text="Units across recorded movements", fill=muted, font=(FONT, 8))
            elif kind == "finance":
                vals = (data["income"], data["expense"])
                colors = (THEME["good"], THEME["bad"])
                labels = ("Income", "Expenses")
                maxv = max(vals) or 1
                for i, (v, color, label) in enumerate(zip(vals, colors, labels)):
                    y = 24 + i*48
                    canvas.create_text(8, y+10, anchor="w", text=label, fill=text, font=(FONT, 9))
                    x0 = 82; x1 = x0 + max(2, (w-150)*v/maxv)
                    canvas.create_rectangle(x0, y, x1, y+20, fill=color, outline="")
                    canvas.create_text(w-8, y+10, anchor="e", text=self.money(v), fill=text, font=(FONT, 9, "bold"))
                canvas.create_text(8, h-8, anchor="sw", text=f"Net: {self.money(data['income']-data['expense'])}", fill=muted, font=(FONT, 8))
            else:
                canvas.create_text(8, 8, anchor="nw", text=f"Units on hand: {st['units']:,}    Low / out: {st['low']}", fill=text, font=(FONT, 10, "bold"))
                rows = data["low"]
                if not rows:
                    canvas.create_text(10, 38, anchor="nw", text="No low-stock items", fill=THEME["good"], font=(FONT, 9))
                else:
                    for i, r in enumerate(rows[:5]):
                        canvas.create_text(10, 34+i*20, anchor="nw", text=f"{r['name'][:24]}  {r['quantity']} / min {r['min_stock']}", fill=THEME["bad"], font=(FONT, 9))

    def _page_items(self, p):
        self._title(p, "Items", "Browse, search and manage everything in stock.")
        self._inventory_page(p, with_cards=False)

    def _page_billing(self, p):
        self._title(p, "Billing", "Create a bill from current inventory. Checkout immediately updates stock.")
        body = tk.Frame(p, bg=THEME["bg"])
        body.pack(fill="both", expand=True)
        left = tk.Frame(body, bg=THEME["bg"])
        left.pack(side="left", fill="both", expand=True, padx=(0, px(12)))
        right = self._card(body)
        right.pack(side="right", fill="both", expand=True)
        bar = tk.Frame(left, bg=THEME["bg"])
        bar.pack(fill="x", pady=(0, px(10)))
        customer_bar = tk.Frame(left, bg=THEME["bg"])
        customer_bar.pack(fill="x", pady=(0, px(8)))
        customer_label = "Customer (required)" if self.customer_required else "Customer (optional)"
        tk.Label(customer_bar, text=customer_label, font=(FONT, 9, "bold"), bg=THEME["bg"], fg=THEME["text"]).pack(side="left", padx=(0, px(8)))
        self.bill_customer = tk.StringVar()
        self.bill_customer_ids = {}
        self.bill_customer_box = ttk.Combobox(customer_bar, textvariable=self.bill_customer, state="readonly")
        self.bill_customer_box.pack(side="left", fill="x", expand=True)
        self.bill_customer_box.bind("<<ComboboxSelected>>", lambda e: setattr(
            self, "bill_selected_customer", self.bill_customer_ids.get(self.bill_customer.get())))
        self._fill_bill_customers()
        FlatButton(customer_bar, "Create customer", lambda: self.show_page("customers")).pack(side="left", padx=(px(8), 0))
        self.bill_search = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.bill_search)
        ent.pack(side="left", fill="x", expand=True)
        ent.bind("<KeyRelease>", lambda e: self._fill_bill_items())
        FlatButton(bar, "Add selected", self._bill_add, "primary").pack(side="left", padx=(px(8), 0))
        self.bill_items = ttk.Treeview(left, columns=("sku", "name", "stock", "price"), show="headings", selectmode="browse")
        for col, label, width in (("sku", "SKU", 90), ("name", "Item", 180), ("stock", "Available", 75), ("price", "Price", 95)):
            self.bill_items.heading(col, text=label); self.bill_items.column(col, width=px(width), anchor="w" if col in ("sku", "name") else "e")
        self.bill_items.pack(fill="both", expand=True)
        self.bill_items.bind("<Double-1>", lambda e: self._bill_add())
        tk.Label(right, text="Current bill", font=(FONT, 12, "bold"), bg=THEME["surface"], fg=THEME["text"]).pack(anchor="w", padx=px(14), pady=px(12))
        self.bill_tree = ttk.Treeview(right, columns=("qty", "price", "total"), show="tree headings", height=12)
        self.bill_tree.heading("#0", text="Item"); self.bill_tree.column("#0", width=px(145))
        for col, label in (("qty", "Qty"), ("price", "Price"), ("total", "Total")):
            self.bill_tree.heading(col, text=label); self.bill_tree.column(col, width=px(75), anchor="e")
        self.bill_tree.pack(fill="both", expand=True, padx=px(10))
        self.bill_total = tk.StringVar(value=self.money(0))
        tk.Label(right, textvariable=self.bill_total, font=(FONT, 18, "bold"), bg=THEME["surface"], fg=THEME["text"]).pack(anchor="e", padx=px(14), pady=px(12))
        actions = tk.Frame(right, bg=THEME["surface"]); actions.pack(fill="x", padx=px(12), pady=(0, px(12)))
        FlatButton(actions, "Remove line", self._bill_remove).pack(side="left")
        FlatButton(actions, "Record expense", self._expense_dialog).pack(side="left", padx=px(6))
        FlatButton(actions, "Checkout", self._bill_checkout, "primary").pack(side="right")
        self._fill_bill_items()
        self._draw_bill()

    def _fill_bill_items(self):
        if not getattr(self, "bill_items", None): return
        self.bill_items.delete(*self.bill_items.get_children())
        q = self.bill_search.get().strip().lower()
        for r in self.db.list_items():
            if q and q not in r["name"].lower() and q not in r["sku"].lower(): continue
            self.bill_items.insert("", "end", iid=str(r["id"]), values=(r["sku"], r["name"], r["quantity"], self.money(r["price"])))

    def _bill_add(self):
        sel = self.bill_items.selection()
        if not sel: return
        item_id = int(sel[0]); item = self.db.get(item_id)
        if not item or item["quantity"] <= self.bill_cart.get(item_id, 0):
            messagebox.showwarning("Stock limit", "The bill quantity cannot exceed available stock.", parent=self); return
        self.bill_cart[item_id] = self.bill_cart.get(item_id, 0) + 1
        self._draw_bill()

    def _bill_remove(self):
        sel = self.bill_tree.selection()
        if sel:
            self.bill_cart.pop(int(sel[0]), None); self._draw_bill()

    def _draw_bill(self):
        if not getattr(self, "bill_tree", None): return
        self.bill_tree.delete(*self.bill_tree.get_children()); total = 0
        for iid, qty in self.bill_cart.items():
            item = self.db.get(iid)
            if not item: continue
            line = qty * item["price"]; total += line
            self.bill_tree.insert("", "end", iid=str(iid), text=item["name"], values=(qty, self.money(item["price"]), self.money(line)))
        self.bill_total.set(self.money(total))

    def _bill_checkout(self):
        if not self.bill_cart:
            messagebox.showinfo("Checkout", "Add items to the bill first.", parent=self); return
        customer_id = self.bill_customer_ids.get(self.bill_customer.get()) if getattr(self, "bill_customer", None) else None
        if self.customer_required and customer_id is None:
            messagebox.showwarning("Customer required", "Select a customer before completing this sale.", parent=self); return
        if not messagebox.askyesno("Checkout", f"Complete this sale for {self.bill_total.get()}?", parent=self): return
        try:
            order_id, total = self.db.checkout(self.bill_cart, customer_id)
        except (ValueError, sqlite3.Error) as exc:
            messagebox.showerror("Checkout failed", str(exc), parent=self); return
        self.bill_cart.clear(); self._draw_bill(); self._fill_bill_items(); self.refresh()
        messagebox.showinfo("Sale complete", f"Order #{order_id} recorded. Total: {self.money(total)}", parent=self)

    def _fill_bill_customers(self):
        if not getattr(self, "bill_customer_box", None):
            return
        selected = self.bill_customer.get()
        rows = self.db.customers()
        labels = [f"CUS-{r['id']:06d} · {r['name']} · {r['phone']}" for r in rows]
        self.bill_customer_ids = {label: row["id"] for label, row in zip(labels, rows)}
        self.bill_customer_box.configure(values=labels)
        if selected in self.bill_customer_ids:
            self.bill_customer.set(selected)
        elif self.bill_selected_customer in self.bill_customer_ids.values():
            self.bill_customer.set(next(label for label, cid in self.bill_customer_ids.items()
                                        if cid == self.bill_selected_customer))
        else:
            self.bill_customer.set("")

    def _page_customers(self, p):
        self._title(p, "Customers", "Create customer profiles and track their orders and lifetime spend.")
        setting = tk.Frame(p, bg=THEME["bg"])
        setting.pack(fill="x", pady=(0, px(12)))
        self.v_customer_required = tk.BooleanVar(value=self.customer_required)
        ttk.Checkbutton(setting, text="Customer details required at billing checkout",
                        variable=self.v_customer_required, command=self._save_customer_required).pack(side="left")
        card = self._card(p)
        card.pack(fill="x", pady=(0, px(12)))
        form = ttk.Frame(card, style="Card.TFrame", padding=14)
        form.pack(fill="x")
        self.customer_name_var = tk.StringVar()
        self.customer_phone_var = tk.StringVar()
        ttk.Label(form, text="Customer name", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(form, textvariable=self.customer_name_var, width=28).grid(row=0, column=1, padx=(0, 16))
        ttk.Label(form, text="Phone number", style="Card.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 8))
        ttk.Entry(form, textvariable=self.customer_phone_var, width=22).grid(row=0, column=3, padx=(0, 14))
        FlatButton(form, "Create customer", self._create_customer, "primary").grid(row=0, column=4)

        tools = tk.Frame(p, bg=THEME["bg"])
        tools.pack(fill="x", pady=(0, px(8)))
        self.customer_search = tk.StringVar()
        customer_search_entry = ttk.Entry(tools, textvariable=self.customer_search)
        customer_search_entry.pack(side="left", fill="x", expand=True)
        self.customer_search.set("")
        customer_search_entry.bind("<KeyRelease>", lambda e: self._refresh_customers())
        tk.Label(tools, text="Customer ID is assigned automatically and stays linked to their orders.",
                 font=(FONT, 9), bg=THEME["bg"], fg=THEME["muted"]).pack(side="right", padx=(px(12), 0))
        table_card = self._card(p)
        table_card.pack(fill="both", expand=True)
        columns = (("id", "Customer ID", 115), ("name", "Name", 190), ("phone", "Phone", 140),
                   ("orders", "Orders", 85), ("spent", "Lifetime spend", 135), ("last", "Last order", 155))
        self.customer_tree = ttk.Treeview(table_card, columns=[c[0] for c in columns], show="headings")
        for col, label, width in columns:
            self.customer_tree.heading(col, text=label)
            self.customer_tree.column(col, width=px(width), anchor="e" if col in ("orders", "spent") else "w")
        self.customer_tree.pack(fill="both", expand=True, padx=px(8), pady=px(8))
        self._refresh_customers()

    def _save_customer_required(self):
        self.customer_required = self.v_customer_required.get()
        self.db.set_setting("customer_required", "1" if self.customer_required else "0")

    def _create_customer(self):
        name = self.customer_name_var.get().strip()
        phone = self.customer_phone_var.get().strip()
        if not name:
            messagebox.showwarning("Check customer", "Enter the customer's name.", parent=self); return
        digits = sum(ch.isdigit() for ch in phone)
        if digits < 7:
            messagebox.showwarning("Check customer", "Enter a valid phone number with at least 7 digits.", parent=self); return
        try:
            customer_id = self.db.create_customer(name, phone)
        except sqlite3.Error as exc:
            messagebox.showerror("Could not create customer", str(exc), parent=self); return
        self.customer_name_var.set(""); self.customer_phone_var.set("")
        self._refresh_customers()
        if getattr(self, "bill_customer_box", None):
            self._fill_bill_customers()
            for label, cid in self.bill_customer_ids.items():
                if cid == customer_id:
                    self.bill_customer.set(label); break
        self.bill_selected_customer = customer_id
        messagebox.showinfo("Customer created", f"Customer ID: CUS-{customer_id:06d}", parent=self)

    def _refresh_customers(self):
        if not getattr(self, "customer_tree", None):
            return
        self.customer_tree.delete(*self.customer_tree.get_children())
        for row in self.db.customers(self.customer_search.get().strip()):
            self.customer_tree.insert("", "end", iid=str(row["id"]), values=(
                f"CUS-{row['id']:06d}", row["name"], row["phone"], row["orders"],
                self.money(row["spent"]), row["last_order"] or "—"))

    def _expense_dialog(self):
        win = tk.Toplevel(self); win.title("Record expense"); win.transient(self); win.grab_set()
        f = ttk.Frame(win, padding=16); f.pack(fill="both", expand=True)
        desc, amount = tk.StringVar(), tk.StringVar()
        ttk.Label(f, text="Description").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=desc, width=30).grid(row=0, column=1, pady=4)
        ttk.Label(f, text="Amount").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=amount, width=30).grid(row=1, column=1, pady=4)
        def save():
            try:
                value = parse_float(amount.get(), "Amount", 0.01)
                if not desc.get().strip(): raise ValueError("Enter a description.")
                self.db.add_expense(desc.get().strip(), value)
            except ValueError as exc:
                messagebox.showwarning("Check expense", str(exc), parent=win); return
            win.destroy(); self.refresh()
        FlatButton(f, "Save expense", save, "primary").grid(row=2, column=1, sticky="e", pady=(10, 0))

    def _currency_glyph(self):
        cur = self.currency.strip().rstrip(".")
        if cur.lower() in ("rs", "inr", "\u20b9"):
            return "\u20B9"
        return cur if len(cur) == 1 else "\u00A4"

    def _inventory_page(self, p, with_cards):
        T = THEME
        if with_cards:
            self._build_cards(p)
        self._build_toolbar(p, 16 if with_cards else 0)

        # status bar (packed first so it is never squeezed out)
        bar = tk.Frame(p, bg=T["bg"])
        bar.pack(side="bottom", fill="x", pady=(px(10), 0))
        self.status_count, self.status_db = tk.StringVar(), tk.StringVar()
        tk.Label(bar, text="\u25A4", font=(SYM, 11), bg=T["bg"], fg=T["muted"]).pack(side="left")
        tk.Label(bar, textvariable=self.status_count, font=(FONT, 9), bg=T["bg"], fg=T["muted"]).pack(
            side="left", padx=(px(6), 0))
        tk.Frame(bar, width=1, height=px(14), bg=T["border"]).pack(side="left", padx=px(14))
        tk.Label(bar, text="\u25A3", font=(SYM, 11), bg=T["bg"], fg=T["muted"]).pack(side="left")
        tk.Label(bar, textvariable=self.status_db, font=(FONT, 9), bg=T["bg"], fg=T["muted"]).pack(
            side="left", padx=(px(6), 0))

        card = self._card(p)
        card.pack(fill="both", expand=True)
        holder = tk.Frame(card, bg=T["surface"])
        holder.pack(fill="both", expand=True, padx=1, pady=1)
        ids = [c[0] for c in self.COLUMNS]
        self.tree = ttk.Treeview(holder, columns=ids, show="headings", selectmode="browse")
        for cid, text, width, anchor in self.COLUMNS:
            if cid == "check":
                self.tree.heading(cid, text=text, command=self._toggle_all)
                self.tree.column(cid, width=px(width), minwidth=px(width), anchor=anchor, stretch=False)
            else:
                self.tree.heading(cid, text=text, command=lambda c=cid: self.sort_by(c))
                self.tree.column(cid, width=px(width), anchor=anchor, minwidth=px(60))
        self.tree.tag_configure("odd", background=T["row_alt"])
        self.tree.tag_configure("low", background=T["low_bg"], foreground=T["low_fg"])
        self.tree.tag_configure("out", background=T["out_bg"], foreground=T["out_fg"])
        self.tree.tag_configure("sel", background=T["sel_bg"], foreground=T["text"])
        vs = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(holder, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)

        # Empty-state overlay
        self.empty = tk.Frame(holder, bg=T["surface"])
        IconBadge(self.empty, 64, T["accent"], "\u25A3", T["surface"], font_size=20).pack()
        self.empty_title = tk.Label(self.empty, font=(FONT, 10, "bold"), bg=T["surface"], fg=T["text"])
        self.empty_title.pack(pady=(px(8), 0))
        self.empty_hint = tk.Label(self.empty, font=(FONT, 9), bg=T["surface"], fg=T["muted"])
        self.empty_hint.pack(pady=(px(2), 0))

        self.tree.bind("<Button-1>", self._tree_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda e: self.edit_item() if self.tree.identify_row(e.y) else None)
        self.tree.bind("<Button-3>", self._popup)
        self.tree.bind("<Return>", lambda e: self.edit_item())
        self.tree.bind("<Delete>", lambda e: self.delete_item())
        self.tree.bind("<space>", self._space_toggle)
        self._page_refresh = self._refresh_inventory

    def _build_cards(self, p):
        T = THEME
        wrap = tk.Frame(p, bg=T["bg"])
        wrap.pack(fill="x")
        self.card_vars = {}
        spec = (("items", "TOTAL ITEMS", "\u25A3"), ("units", "UNITS IN STOCK", "\u25A4"),
                ("value", "STOCK VALUE", self._currency_glyph()), ("low", "LOW / OUT OF STOCK", "\u26A0"))
        for i, (key, title, glyph) in enumerate(spec):
            pal = T["cards"][key]
            card = tk.Frame(wrap, bg=pal["bg"], highlightthickness=1, highlightbackground=pal["border"])
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else px(14), 0))
            wrap.columnconfigure(i, weight=1, uniform="cards")
            inner = tk.Frame(card, bg=pal["bg"])
            inner.pack(fill="x", padx=px(16), pady=px(14))
            IconBadge(inner, 54, pal["icon"], glyph, pal["bg"]).pack(side="left")
            txt = tk.Frame(inner, bg=pal["bg"])
            txt.pack(side="left", padx=(px(14), 0))
            tk.Label(txt, text=title, font=(FONT, 9, "bold"), bg=pal["bg"], fg=pal["title"]).pack(anchor="w")
            var = tk.StringVar(value="-")
            tk.Label(txt, textvariable=var, font=(FONT, 19, "bold"), bg=pal["bg"], fg=T["text"]).pack(anchor="w")
            self.card_vars[key] = var

    def _build_toolbar(self, p, top):
        T = THEME
        tool = tk.Frame(p, bg=T["bg"])
        tool.pack(fill="x", pady=(px(top), px(12)))
        FlatButton(tool, "Add Item", self.add_item, "primary", icon="+", padx=18).pack(side="left")
        self.btn_edit = FlatButton(tool, "Edit", self.edit_item, icon="\u270E")
        self.btn_stock = FlatButton(tool, "Adjust Stock", self.adjust_stock, icon="\u21C4")
        self.btn_hist = FlatButton(tool, "History", self.show_history, icon="\u25F7")
        self.btn_del = FlatButton(tool, "Delete", self.delete_item, "danger", icon="\u2716")
        for b in (self.btn_edit, self.btn_stock, self.btn_hist, self.btn_del):
            b.pack(side="left", padx=(px(8), 0))

        FlatButton(tool, "Clear", self.clear_filters, icon="\u21BB").pack(side="right")
        ttk.Checkbutton(tool, text="Low stock only", variable=self.low_only,
                        command=self._refresh_inventory).pack(side="right", padx=px(14))
        self.cat_box = ttk.Combobox(tool, textvariable=self.cat_var, state="readonly", width=18)
        self.cat_box.pack(side="right")
        self.cat_box.bind("<<ComboboxSelected>>", lambda e: (self.cat_box.selection_clear(),
                                                             self._refresh_inventory()))
        self.search_box = SearchBox(tool, "Search by item name, SKU...", self._search_changed, self.search_text)
        self.search_box.pack(side="right", padx=(0, px(12)))

    def _search_changed(self):
        self.search_text = self.search_box.get()
        self._refresh_inventory()

    # -- data display -----------------------------------------------------
    def money(self, v):
        return f"{self.currency}{v:,.2f}"

    def _selected_id(self):
        if not self.tree:
            return None
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def _target_id(self):
        sel = self._selected_id()
        if sel is not None:
            return sel
        if len(self.checked) == 1:
            return next(iter(self.checked))
        return None

    def _update_buttons(self):
        if not self.tree:
            return
        tid = self._target_id()
        for b in (self.btn_edit, self.btn_stock, self.btn_hist):
            b.set_enabled(tid is not None)
        n = len(self.checked)
        self.btn_del.set_enabled(bool(n) or self._selected_id() is not None)
        self.btn_del.configure(text=f"\u2716  Delete ({n})" if n else "\u2716  Delete")

    def _sort_key(self, row):
        c = self.sort_col
        if c == "value":
            return row["quantity"] * row["price"]
        if c == "status":
            return row["quantity"] - row["min_stock"]
        if c in self.NUMERIC:
            return row[c]
        if c == "check":
            return row["id"] in self.checked
        return str(row[c]).lower()

    def sort_by(self, col):
        self.sort_rev = (not self.sort_rev) if col == self.sort_col else False
        self.sort_col = col
        self._refresh_inventory()

    def _on_select(self, _e=None):
        if not self.tree:
            return
        sel = self.tree.selection()
        new = sel[0] if sel else None
        prev = self._sel_prev
        if prev and prev != new and self.tree.exists(prev):
            self.tree.item(prev, tags=self._row_tags.get(prev, ()))
        if new:
            self.tree.item(new, tags=("sel",))
        self._sel_prev = new
        self._update_buttons()

    def _tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) == "cell" and self.tree.identify_column(event.x) == "#1":
            row = self.tree.identify_row(event.y)
            if row:
                self._toggle_check(int(row))
                return "break"

    def _space_toggle(self, _e):
        sel = self._selected_id()
        if sel is not None:
            self._toggle_check(sel)
        return "break"

    def _toggle_check(self, item_id):
        self.checked.symmetric_difference_update({item_id})
        self.tree.set(str(item_id), "check", CHK_ON if item_id in self.checked else CHK_OFF)
        self._update_check_header()
        self._update_buttons()

    def _toggle_all(self):
        visible = [int(i) for i in self.tree.get_children()]
        if not visible:
            return
        if all(i in self.checked for i in visible):
            self.checked.difference_update(visible)
        else:
            self.checked.update(visible)
        for i in visible:
            self.tree.set(str(i), "check", CHK_ON if i in self.checked else CHK_OFF)
        self._update_check_header()
        self._update_buttons()

    def _update_check_header(self):
        visible = [int(i) for i in self.tree.get_children()]
        on = bool(visible) and all(i in self.checked for i in visible)
        self.tree.heading("check", text=CHK_ON if on else CHK_OFF)

    def _refresh_inventory(self):
        if not self.tree:
            return
        cats = self.db.categories()
        self.cat_box.configure(values=[ALL_CATS] + cats)
        if self.cat_var.get() not in cats:
            self.cat_var.set(ALL_CATS)
        category = "" if self.cat_var.get() == ALL_CATS else self.cat_var.get()

        keep = self._selected_id()
        rows = self.db.list_items(self.search_text, category, self.low_only.get())
        rows.sort(key=self._sort_key, reverse=self.sort_rev)

        self.tree.delete(*self.tree.get_children())
        self._row_tags, self._sel_prev = {}, None
        for n, r in enumerate(rows):
            if r["quantity"] == 0:
                status, tags = "OUT", ("out",)
            elif r["quantity"] <= r["min_stock"]:
                status, tags = "LOW", ("low",)
            else:
                status, tags = "OK", (("odd",) if n % 2 else ())
            iid = str(r["id"])
            self._row_tags[iid] = tags
            self.tree.insert("", "end", iid=iid, tags=tags, values=(
                CHK_ON if r["id"] in self.checked else CHK_OFF, r["sku"], r["name"], r["category"],
                r["quantity"], self.money(r["price"]), self.money(r["quantity"] * r["price"]),
                r["min_stock"], status))
        if keep and self.tree.exists(str(keep)):
            self.tree.selection_set(str(keep))
            self.tree.see(str(keep))
            self._on_select()

        for cid, text, *_ in self.COLUMNS:  # sort arrow on the active heading
            if cid != "check":
                arrow = (" \u25BC" if self.sort_rev else " \u25B2") if cid == self.sort_col else ""
                self.tree.heading(cid, text=text + arrow)
        self._update_check_header()

        st = self.db.stats()
        if self.card_vars:
            self.card_vars["items"].set(f"{st['items']:,}")
            self.card_vars["units"].set(f"{st['units']:,}")
            self.card_vars["value"].set(self.money(st["value"]))
            self.card_vars["low"].set(f"{st['low']:,}")
        self.status_count.set(f"Showing {len(rows)} of {st['items']} items")
        self.status_db.set(f"Database: {self.db.path}")

        if rows:
            self.empty.place_forget()
        else:
            if st["items"] == 0:
                self.empty_title.configure(text="No items to display yet.")
                self.empty_hint.configure(text='Click on "Add Item" to create your first item.')
            else:
                self.empty_title.configure(text="No items match your filters.")
                self.empty_hint.configure(text="Try a different search, or click Clear.")
            self.empty.place(relx=0.5, rely=0.5, anchor="center")
        self._update_buttons()

    def refresh(self):
        st = self.db.stats()
        self._update_bell(st["low"])
        if self.checked:
            self.checked &= self.db.existing_ids()
        if self._page_refresh:
            self._page_refresh()

    def _update_bell(self, low):
        try:
            state = "normal" if low else "hidden"
            self.bell.itemconfigure(self.bell_dot, state=state)
            self.bell.itemconfigure(self.bell_num, state=state, text=str(low) if low < 10 else "9+")
        except (tk.TclError, AttributeError):
            pass

    def bell_click(self):
        if self.db.stats()["low"] == 0:
            messagebox.showinfo("Notifications", "Everything is well stocked. No low-stock items.", parent=self)
            return
        self.show_low_stock()

    def clear_filters(self):
        self.search_text = ""
        if self.search_box:
            self.search_box.set("")
        self.cat_var.set(ALL_CATS)
        self.low_only.set(False)
        self._refresh_inventory()

    def show_low_stock(self):
        self.low_only.set(True)
        if self.page != "dashboard":
            self.show_page("dashboard")
        else:
            self._refresh_inventory()

    def _popup(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.context.tk_popup(event.x_root, event.y_root)

    # ----------------------------------------------------------------------
    # Reports page
    # ----------------------------------------------------------------------
    def _page_reports(self, p):
        T = THEME
        self._title(p, "Reports", "Review stock levels, movements and value.")
        bar = tk.Frame(p, bg=T["bg"])
        bar.pack(fill="x", pady=(0, px(12)))
        self.rep_btns = {}
        for key, label in self.REPORTS:
            b = FlatButton(bar, label, lambda k=key: self._set_report(k), padx=16)
            b.pack(side="left", padx=(0, px(8)))
            self.rep_btns[key] = b
        FlatButton(bar, "Export report", self.export_report, icon="\u21E9").pack(side="right")

        status = tk.Frame(p, bg=T["bg"])
        status.pack(side="bottom", fill="x", pady=(px(10), 0))
        self.rep_info = tk.StringVar()
        tk.Label(status, textvariable=self.rep_info, font=(FONT, 9), bg=T["bg"], fg=T["muted"]).pack(side="left")

        card = self._card(p)
        card.pack(fill="both", expand=True)
        holder = tk.Frame(card, bg=T["surface"])
        holder.pack(fill="both", expand=True, padx=1, pady=1)
        self.rep_tree = ttk.Treeview(holder, show="headings", selectmode="browse")
        for tag, bg, fg in (("odd", T["row_alt"], T["text"]), ("low", T["low_bg"], T["low_fg"]),
                            ("out", T["out_bg"], T["out_fg"])):
            self.rep_tree.tag_configure(tag, background=bg, foreground=fg)
        self.rep_tree.tag_configure("good", foreground=T["good"])
        self.rep_tree.tag_configure("bad", foreground=T["bad"])
        vs = ttk.Scrollbar(holder, orient="vertical", command=self.rep_tree.yview)
        hs = ttk.Scrollbar(holder, orient="horizontal", command=self.rep_tree.xview)
        self.rep_tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.rep_tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        self.rep_empty = tk.Label(holder, text="Nothing to show for this report.", font=(FONT, 10),
                                  bg=T["surface"], fg=T["muted"])
        self._page_refresh = self._fill_report

    def _set_report(self, key):
        self.report_kind = key
        self._fill_report()

    def _report_data(self):
        """Return (columns, rows) where rows are (values, tag)."""
        kind = self.report_kind
        if kind == "low":
            cols = [("sku", "SKU", 110, "w"), ("name", "Item name", 260, "w"), ("category", "Category", 150, "w"),
                    ("qty", "Qty", 80, "e"), ("alert", "Alert level", 100, "e"), ("short", "Shortfall", 100, "e"),
                    ("status", "Status", 90, "center")]
            items = sorted(self.db.list_items(low_only=True), key=lambda r: (r["quantity"] - r["min_stock"]))
            rows = [((r["sku"], r["name"], r["category"], r["quantity"], r["min_stock"],
                      max(r["min_stock"] - r["quantity"], 0), "OUT" if r["quantity"] == 0 else "LOW"),
                     "out" if r["quantity"] == 0 else "low") for r in items]
        elif kind == "moves":
            cols = [("ts", "Date / time", 160, "w"), ("item", "Item", 260, "w"), ("change", "Change", 90, "e"),
                    ("qty", "New qty", 90, "e"), ("reason", "Reason", 220, "w")]
            rows = [((r["ts"], r["item_name"], f"{r['change']:+d}", r["new_qty"], r["reason"]),
                     "good" if r["change"] >= 0 else "bad") for r in self.db.movements()]
        else:
            cols = [("cat", "Category", 260, "w"), ("items", "Items", 100, "e"), ("units", "Units", 110, "e"),
                    ("value", "Stock value", 150, "e"), ("low", "Low / out", 110, "e")]
            rows = [((r[0], r[1], f"{r[2]:,}", self.money(r[3]), r[4]), "") for r in self.db.category_summary()]
        return cols, rows

    def _fill_report(self):
        if not self.rep_tree:
            return
        for key, b in self.rep_btns.items():
            b.set_kind("primary" if key == self.report_kind else "secondary")
        cols, rows = self._report_data()
        self._rep_cols, self._rep_rows = cols, rows
        t = self.rep_tree
        t.delete(*t.get_children())
        t.configure(columns=[c[0] for c in cols])
        for cid, text, width, anchor in cols:
            t.heading(cid, text=text)
            t.column(cid, width=px(width), anchor=anchor, minwidth=px(60))
        for n, (vals, tag) in enumerate(rows):
            t.insert("", "end", values=vals, tags=(tag,) if tag else (("odd",) if n % 2 else ()))
        if rows:
            self.rep_empty.place_forget()
        else:
            self.rep_empty.place(relx=0.5, rely=0.45, anchor="center")
        note = "  (latest 1,000)" if self.report_kind == "moves" and len(rows) >= 1000 else ""
        self.rep_info.set(f"{len(rows)} row(s){note}")

    def export_report(self):
        if not getattr(self, "_rep_rows", None):
            messagebox.showinfo("Export report", "There is nothing to export for this report.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="Export report", defaultextension=".csv",
            initialfile=f"report_{self.report_kind}_{datetime.now():%Y%m%d}.csv",
            filetypes=[("CSV files (Excel)", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow([c[1] for c in self._rep_cols])
                for vals, _tag in self._rep_rows:
                    w.writerow(vals)
            messagebox.showinfo("Export complete", f"Report saved to:\n{path}", parent=self)
        except OSError as e:
            messagebox.showerror("Export failed", str(e), parent=self)

    # ----------------------------------------------------------------------
    # Tools page
    # ----------------------------------------------------------------------
    def _page_tools(self, p):
        T = THEME
        self._title(p, "Tools", "Import, export and protect your data.")
        grid = tk.Frame(p, bg=T["bg"])
        grid.pack(fill="x")
        tiles = (
            ("\u21E7", "Import from CSV", "Add or update many items at once from a spreadsheet file.",
             "Import...", self.import_csv, "items"),
            ("\u21E9", "Export to CSV", "Save your full inventory as a file you can open in Excel.",
             "Export...", self.export_csv, "units"),
            ("\u2261", "CSV template", "Download a sample file showing the columns the importer expects.",
             "Save template...", self.save_template, "value"),
            ("\u25C9", "Back up database", "Make a safe copy of all items, history and settings.",
             "Back up...", self.backup, "items"),
            ("\u21BA", "Restore backup", "Replace the current data with a previously saved backup.",
             "Restore...", self.restore, "low"),
            ("\u25A4", "Open data folder", "Show the folder where your database and error log are stored.",
             "Open folder", self.open_data_folder, "units"),
        )
        for i, (glyph, title, desc, btn, cmd, key) in enumerate(tiles):
            pal = T["cards"][key]
            card = self._card(grid)
            card.grid(row=i // 3, column=i % 3, sticky="nsew", padx=px(7), pady=px(7))
            grid.columnconfigure(i % 3, weight=1, uniform="tiles")
            inner = tk.Frame(card, bg=T["surface"])
            inner.pack(fill="both", expand=True, padx=px(18), pady=px(16))
            top = tk.Frame(inner, bg=T["surface"])
            top.pack(fill="x")
            IconBadge(top, 46, pal["icon"], glyph, T["surface"], font_size=15).pack(side="left")
            tk.Label(top, text=title, font=(FONT, 11, "bold"), bg=T["surface"], fg=T["text"]).pack(
                side="left", padx=(px(12), 0))
            tk.Label(inner, text=desc, font=(FONT, 9), bg=T["surface"], fg=T["muted"], justify="left",
                     anchor="w", wraplength=px(260)).pack(fill="x", pady=(px(12), px(14)))
            FlatButton(inner, btn, cmd).pack(anchor="w")

    # ----------------------------------------------------------------------
    # Settings page
    # ----------------------------------------------------------------------
    def _page_settings(self, p):
        self._title(p, "Settings", "Customise how Tory looks and behaves.")
        card = self._card(p)
        card.pack(fill="x")
        f = ttk.Frame(card, style="Card.TFrame", padding=24)
        f.pack(fill="x")
        self.v_name = tk.StringVar(value=self.user_name)
        self.v_cur = tk.StringVar(value=self.currency)
        self.v_min = tk.StringVar(value=str(self.default_min))
        self.v_theme = tk.StringVar(value=self.theme_name)

        rows = (("Display name", self.v_name, "Shown in the top-right corner."),
                ("Currency symbol", self.v_cur, 'For example "Rs. ", "$" or "\u20AC". Spaces are kept.'),
                ("Default low-stock level", self.v_min, "Used for new items and CSV rows without an alert level."))
        for i, (label, var, hint) in enumerate(rows):
            ttk.Label(f, text=label, style="Card.TLabel").grid(row=i * 2, column=0, sticky="w", padx=(0, 24), pady=(8, 0))
            ttk.Entry(f, textvariable=var, width=24).grid(row=i * 2, column=1, sticky="w", pady=(8, 0))
            ttk.Label(f, text=hint, style="Card.Muted.TLabel").grid(row=i * 2 + 1, column=1, sticky="w")
        r = len(rows) * 2
        ttk.Label(f, text="Appearance", style="Card.TLabel").grid(row=r, column=0, sticky="w", pady=(16, 0))
        modes = ttk.Frame(f, style="Card.TFrame")
        modes.grid(row=r, column=1, sticky="w", pady=(16, 0))
        ttk.Radiobutton(modes, text="Light", variable=self.v_theme, value="light",
                        style="Card.TRadiobutton").pack(side="left", padx=(0, 20))
        ttk.Radiobutton(modes, text="Dark", variable=self.v_theme, value="dark",
                        style="Card.TRadiobutton").pack(side="left")
        FlatButton(f, "Save settings", self.save_settings, "primary", padx=20).grid(
            row=r + 1, column=1, sticky="w", pady=(22, 0))

    def save_settings(self):
        try:
            low = parse_int(self.v_min.get(), "Default low-stock level", 0)
        except ValueError as e:
            messagebox.showwarning("Check your input", str(e), parent=self)
            return
        self.user_name = self.v_name.get().strip() or "Admin"
        self.currency = self.v_cur.get()
        self.default_min = low
        self.db.set_setting("user_name", self.user_name)
        self.db.set_setting("currency", self.currency)
        self.db.set_setting("default_min", str(low))
        new_theme = self.v_theme.get()
        if new_theme != self.theme_name:
            self.apply_theme(new_theme)
        else:
            self._build_shell()
            self.refresh()
        messagebox.showinfo("Settings", "Your settings have been saved.", parent=self)

    def apply_theme(self, name):
        self.theme_name = name if name in THEMES else "light"
        set_theme(self.theme_name)
        self.db.set_setting("theme", self.theme_name)
        self._style()
        self._build_shell()
        self.refresh()

    def toggle_theme(self):
        self.apply_theme("dark" if self.theme_name == "light" else "light")

    # ----------------------------------------------------------------------
    # Help page
    # ----------------------------------------------------------------------
    def _page_help(self, p):
        self._title(p, "Help", "Shortcuts, import format and application details.")
        card = self._card(p)
        card.pack(fill="both", expand=True)
        f = ttk.Frame(card, style="Card.TFrame", padding=24)
        f.pack(fill="both", expand=True)
        f.columnconfigure(0, weight=1, uniform="h")
        f.columnconfigure(1, weight=1, uniform="h")

        left = ttk.Frame(f, style="Card.TFrame")
        left.grid(row=0, column=0, sticky="nw", padx=(0, 20))
        ttk.Label(left, text="Keyboard shortcuts", style="Card.Heading.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        keys = (("Ctrl + N", "Add a new item"), ("Enter", "Edit the selected item"),
                ("Ctrl + S", "Adjust stock of the selected item"), ("Space", "Tick / untick the selected row"),
                ("Delete", "Delete the ticked or selected items"), ("Ctrl + F", "Jump to search"),
                ("Ctrl + E", "Export inventory to CSV"), ("Ctrl + T", "Switch light / dark theme"),
                ("F5", "Refresh"))
        for i, (k, d) in enumerate(keys, start=1):
            ttk.Label(left, text=k, style="Card.TLabel", font=(FONT, 10, "bold")).grid(
                row=i, column=0, sticky="w", pady=3, padx=(0, 18))
            ttk.Label(left, text=d, style="Card.Muted.TLabel").grid(row=i, column=1, sticky="w", pady=3)

        right = ttk.Frame(f, style="Card.TFrame")
        right.grid(row=0, column=1, sticky="nw")
        ttk.Label(right, text="Importing from CSV", style="Card.Heading.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Label(right, style="Card.Muted.TLabel", justify="left", wraplength=px(420), text=(
            "The first row must contain column names. Only 'name' is required.\n"
            "Optional columns: sku, category, quantity, price, min_stock.\n"
            "Rows whose name already exists update that item. Use Tools \u2192 CSV template for an example.")
        ).pack(anchor="w")
        ttk.Label(right, text="About", style="Card.Heading.TLabel").pack(anchor="w", pady=(20, 8))
        ttk.Label(right, style="Card.Muted.TLabel", justify="left", wraplength=px(420), text=(
            f"{APP_NAME} {VERSION}\nSimple, fast stock control for Windows.\n\nData file:\n{self.db.path}")
        ).pack(anchor="w")

    # ----------------------------------------------------------------------
    # Key bindings
    # ----------------------------------------------------------------------
    def _bind_keys(self):
        self.bind("<Control-n>", lambda e: self.add_item())
        self.bind("<Control-f>", lambda e: self.search_box.focus_entry() if self.search_box else None)
        self.bind("<Control-s>", lambda e: self.adjust_stock())
        self.bind("<Control-e>", lambda e: self.export_csv())
        self.bind("<Control-t>", lambda e: self.toggle_theme())
        self.bind("<F5>", lambda e: self.refresh())

    # ----------------------------------------------------------------------
    # Actions
    # ----------------------------------------------------------------------
    def add_item(self):
        data = ItemDialog(self, self.db.categories(), self.default_min).show()
        if data:
            try:
                new_id = self.db.add_item(data)
            except ValueError as e:
                messagebox.showerror("Cannot add item", str(e), parent=self)
                return
            self.refresh()
            if self.tree and self.tree.exists(str(new_id)):
                self.tree.selection_set(str(new_id))
                self.tree.see(str(new_id))

    def edit_item(self):
        item_id = self._target_id()
        if not item_id:
            return
        item = self.db.get(item_id)
        if item is None:
            self.refresh()
            return
        data = ItemDialog(self, self.db.categories(), self.default_min, item).show()
        if data:
            try:
                self.db.update_item(item_id, data)
            except ValueError as e:
                messagebox.showerror("Cannot save item", str(e), parent=self)
                return
            self.refresh()

    def adjust_stock(self):
        item_id = self._target_id()
        if not item_id:
            return
        item = self.db.get(item_id)
        if item is None:
            self.refresh()
            return
        result = StockDialog(self, item).show()
        if result:
            try:
                self.db.adjust_stock(item_id, *result)
            except ValueError as e:
                messagebox.showerror("Cannot adjust stock", str(e), parent=self)
                return
            self.refresh()

    def delete_item(self):
        ids = sorted(self.checked)
        if not ids:
            sel = self._selected_id()
            ids = [sel] if sel else []
        if not ids:
            return
        if len(ids) == 1:
            item = self.db.get(ids[0])
            if item is None:
                self.refresh()
                return
            msg = f"Delete '{item['name']}' permanently?\nThis cannot be undone."
        else:
            msg = f"Delete {len(ids)} selected items permanently?\nThis cannot be undone."
        if messagebox.askyesno("Delete", msg, icon="warning", default="no", parent=self):
            for i in ids:
                self.db.delete_item(i)
            self.checked.difference_update(ids)
            self.refresh()

    def show_history(self, all_items=False):
        item_id = None if all_items else self._target_id()
        if not all_items and not item_id:
            return
        item = self.db.get(item_id) if item_id else None
        title = "Stock movement history" if all_items else f"History - {item['name'] if item else 'deleted item'}"
        HistoryDialog(self, self.db.movements(item_id), title).show()

    # -- import / export / backup ------------------------------------------
    def export_csv(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Export inventory", defaultextension=".csv",
            initialfile=f"inventory_{datetime.now():%Y%m%d}.csv",
            filetypes=[("CSV files (Excel)", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            rows = sorted(self.db.list_items(), key=lambda r: r["name"].lower())
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["sku", "name", "category", "quantity", "price", "min_stock"])
                for r in rows:
                    w.writerow([r["sku"], r["name"], r["category"], r["quantity"], r["price"], r["min_stock"]])
            messagebox.showinfo("Export complete", f"Exported {len(rows)} items to:\n{path}", parent=self)
        except OSError as e:
            messagebox.showerror("Export failed", str(e), parent=self)

    def save_template(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Save CSV template", defaultextension=".csv", initialfile="inventory_template.csv",
            filetypes=[("CSV files (Excel)", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["sku", "name", "category", "quantity", "price", "min_stock"])
                w.writerow(["PEN-001", "Blue Ballpoint Pen", "Stationery", 120, 10.5, 20])
                w.writerow(["NTB-002", "A4 Notebook", "Stationery", 45, 60, 10])
            messagebox.showinfo("Template saved", f"Template saved to:\n{path}", parent=self)
        except OSError as e:
            messagebox.showerror("Could not save template", str(e), parent=self)

    def import_csv(self):
        path = filedialog.askopenfilename(
            parent=self, title="Import inventory",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        added = updated = 0
        errors = []
        try:
            with open(path, newline="", encoding="utf-8-sig") as fh:
                reader = csv.DictReader(fh)
                reader.fieldnames = [(h or "").strip().lower() for h in (reader.fieldnames or [])]
                if "name" not in reader.fieldnames:
                    raise ValueError("The CSV must have a 'name' column.\n"
                                     "Optional: sku, category, quantity, price, min_stock.")
                for line, row in enumerate(reader, start=2):
                    try:
                        name = (row.get("name") or "").strip()
                        if not name:
                            raise ValueError("name is empty")
                        data = {
                            "name": name,
                            "sku": (row.get("sku") or "").strip(),
                            "category": (row.get("category") or "").strip(),
                            "quantity": parse_int(row.get("quantity") or 0, "quantity", 0),
                            "price": parse_float(row.get("price") or 0, "price", 0),
                            "min_stock": parse_int(row.get("min_stock") or self.default_min, "min_stock", 0),
                        }
                        existing = self.db.conn.execute(
                            "SELECT id FROM items WHERE lower(name)=lower(?)", (name,)).fetchone()
                        if existing:
                            self.db.update_item(existing["id"], data)
                            updated += 1
                        else:
                            self.db.add_item(data)
                            added += 1
                    except ValueError as e:
                        errors.append(f"Line {line}: {e}")
        except (OSError, UnicodeDecodeError, ValueError, csv.Error) as e:
            messagebox.showerror("Import failed", str(e), parent=self)
            return
        self.refresh()
        msg = f"Added {added} new item(s), updated {updated} existing item(s)."
        if errors:
            msg += f"\n\n{len(errors)} row(s) skipped:\n" + "\n".join(errors[:10])
            if len(errors) > 10:
                msg += "\n..."
        messagebox.showinfo("Import complete", msg, parent=self)

    def backup(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Back up database", defaultextension=".db",
            initialfile=f"inventory_backup_{datetime.now():%Y%m%d_%H%M}.db",
            filetypes=[("Database files", "*.db"), ("All files", "*.*")])
        if path:
            try:
                self.db.backup(path)
                messagebox.showinfo("Backup complete", f"Backup saved to:\n{path}", parent=self)
            except (OSError, sqlite3.Error) as e:
                messagebox.showerror("Backup failed", str(e), parent=self)

    def restore(self):
        path = filedialog.askopenfilename(
            parent=self, title="Restore from backup",
            filetypes=[("Database files", "*.db"), ("All files", "*.*")])
        if not path:
            return
        if os.path.abspath(path) == os.path.abspath(self.db.path):
            messagebox.showinfo("Restore", "That is the database currently in use.", parent=self)
            return
        if not messagebox.askyesno(
                "Restore backup",
                "This replaces ALL current items, history and settings with the backup.\n"
                "Consider making a backup of the current data first.\n\nContinue?",
                icon="warning", default="no", parent=self):
            return
        try:
            self.db.restore(path)
        except sqlite3.Error as e:
            messagebox.showerror("Restore failed", f"This does not look like a valid backup.\n\n{e}", parent=self)
            return
        self.currency = self.db.get_setting("currency", self.currency)
        self.user_name = self.db.get_setting("user_name", self.user_name) or "Admin"
        try:
            self.default_min = int(self.db.get_setting("default_min", str(self.default_min)))
        except ValueError:
            pass
        self.checked.clear()
        self._build_shell()
        self.refresh()
        messagebox.showinfo("Restore complete", "The backup was restored.", parent=self)

    def open_data_folder(self):
        folder = os.path.dirname(self.db.path)
        try:
            if os.name == "nt":
                os.startfile(folder)  # noqa: S606 - Windows only
            else:
                raise OSError
        except OSError:
            messagebox.showinfo("Data folder", folder, parent=self)

    # -- misc -------------------------------------------------------------
    def about(self):
        messagebox.showinfo(
            f"About {APP_NAME}",
            f"{APP_NAME} {VERSION}\n\nSimple, fast stock control for Windows.\n\n"
            f"Data file:\n{self.db.path}", parent=self)

    def _on_error(self, exc, val, tb):
        logging.error("Unhandled error", exc_info=(exc, val, tb))
        messagebox.showerror("Unexpected error",
                             f"Something went wrong:\n{val}\n\nDetails were saved to error.log", parent=self)


# --------------------------------------------------------------------------
# Start-up
# --------------------------------------------------------------------------
def enable_high_dpi():
    """Crisp text on high-DPI screens; silently skipped where unsupported."""
    if os.name != "nt":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1+
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()       # Windows 7 / Vista
    except Exception:
        pass


def open_database():
    path = os.path.join(data_dir(), DB_NAME)
    if not os.path.exists(path):
        # Bring over an inventory.db from an older version, if present.
        for old in (os.path.join(app_dir(), DB_NAME), os.path.join(os.getcwd(), DB_NAME)):
            if os.path.isfile(old):
                shutil.copyfile(old, path)
                break
    return Database(path)


def main():
    logging.basicConfig(filename=os.path.join(data_dir(), "error.log"), level=logging.ERROR,
                        format="%(asctime)s %(message)s")
    enable_high_dpi()
    try:
        db = open_database()
        App(db).mainloop()
    except Exception:
        logging.error(traceback.format_exc())
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror(APP_NAME, "The application could not start.\n"
                                           "See error.log in %APPDATA%\\InventoryManager.")
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()

