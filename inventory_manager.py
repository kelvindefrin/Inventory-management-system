
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

APP_NAME = "Inventory Manager"
VERSION = "2.0.0"
DB_NAME = "inventory.db"


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
    try:
        value = int(str(text).strip())
    except ValueError:
        raise ValueError(f"{label} must be a whole number.")
    if minimum is not None and value < minimum:
        raise ValueError(f"{label} must be at least {minimum}.")
    return value


def parse_float(text, label, minimum=None):
    try:
        value = float(str(text).strip().replace(",", ""))
    except ValueError:
        raise ValueError(f"{label} must be a number.")
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
        # Upgrade databases created by the old console version.
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

    def movements(self, item_id=None, limit=1000):
        if item_id is None:
            return self.conn.execute(
                "SELECT * FROM movements ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
        return self.conn.execute(
            "SELECT * FROM movements WHERE item_id=? ORDER BY id DESC LIMIT ?",
            (item_id, limit)).fetchall()

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


# --------------------------------------------------------------------------
# Dialogs
# --------------------------------------------------------------------------
class Dialog(tk.Toplevel):
    """Modal dialog base. Subclasses build into self.body, then call self.show()."""

    def __init__(self, parent, title):
        super().__init__(parent)
        self.withdraw()
        self.parent = parent
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.result = None
        self.body = ttk.Frame(self, padding=18)
        self.body.pack(fill="both", expand=True)
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Escape>", lambda e: self.cancel())

    def buttons(self, ok_text="Save", row=0):
        bar = ttk.Frame(self.body)
        bar.grid(row=row, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(bar, text="Cancel", command=self.cancel).pack(side="right")
        ttk.Button(bar, text=ok_text, style="Accent.TButton", command=self.ok).pack(side="right", padx=(0, 8))
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

        fields = [("Item name *", "name"), ("SKU / Code", "sku"), ("Category", "category"),
                  ("Quantity", "quantity"), ("Price per unit", "price"), ("Low-stock alert level", "min_stock")]
        for i, (label, key) in enumerate(fields):
            ttk.Label(b, text=label).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 14))
            if key == "category":
                w = ttk.Combobox(b, textvariable=self.vars[key], values=categories, width=28)
            else:
                w = ttk.Entry(b, textvariable=self.vars[key], width=30)
            w.grid(row=i, column=1, sticky="ew", pady=5)
            if i == 0:
                first = w
        ttk.Label(b, text="You are warned when quantity is at or below the alert level.",
                  style="Muted.TLabel").grid(row=len(fields), column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.buttons("Save", row=len(fields) + 1)
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
        ttk.Label(b, text=item["name"], style="Heading.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(b, text=f"Current quantity: {item['quantity']}", style="Muted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self.mode = tk.StringVar(value="in")
        modes = ttk.Frame(b)
        modes.grid(row=2, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Radiobutton(modes, text="Stock in (+)", variable=self.mode, value="in",
                        command=self._mode_changed).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(modes, text="Stock out (-)", variable=self.mode, value="out",
                        command=self._mode_changed).pack(side="left")

        ttk.Label(b, text="Quantity").grid(row=3, column=0, sticky="w", pady=6, padx=(0, 14))
        self.qty = tk.StringVar(value="1")
        self.qty.trace_add("write", lambda *a: self._preview())
        qty_entry = ttk.Entry(b, textvariable=self.qty, width=12)
        qty_entry.grid(row=3, column=1, sticky="w", pady=6)

        ttk.Label(b, text="Reason").grid(row=4, column=0, sticky="w", pady=6)
        self.reason = tk.StringVar(value=self.REASONS_IN[0])
        self.reason_box = ttk.Combobox(b, textvariable=self.reason, values=self.REASONS_IN, width=26)
        self.reason_box.grid(row=4, column=1, sticky="w", pady=6)

        self.preview = ttk.Label(b, text="", style="Muted.TLabel")
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


class SettingsDialog(Dialog):
    def __init__(self, parent, currency, default_min):
        super().__init__(parent, "Settings")
        b = self.body
        self.currency = tk.StringVar(value=currency)
        self.default_min = tk.StringVar(value=str(default_min))
        ttk.Label(b, text="Currency symbol").grid(row=0, column=0, sticky="w", pady=6, padx=(0, 14))
        ttk.Entry(b, textvariable=self.currency, width=12).grid(row=0, column=1, sticky="w", pady=6)
        ttk.Label(b, text="Default low-stock level").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(b, textvariable=self.default_min, width=12).grid(row=1, column=1, sticky="w", pady=6)
        self.buttons("Save", row=2)

    def ok(self):
        try:
            low = parse_int(self.default_min.get(), "Default low-stock level", 0)
        except ValueError as e:
            messagebox.showwarning("Check your input", str(e), parent=self)
            return
        self.result = (self.currency.get(), low)
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
            tree.column(col, width=width, anchor=anchor)
        tree.tag_configure("in", foreground="#157347")
        tree.tag_configure("out", foreground="#b02a37")
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
            ttk.Label(b, text="No movements recorded yet.", style="Muted.TLabel").grid(row=1, column=0, pady=8)
        ttk.Button(b, text="Close", command=self.cancel).grid(row=2, column=0, columnspan=2, sticky="e", pady=(10, 0))


# --------------------------------------------------------------------------
# Main window
# --------------------------------------------------------------------------
class App(tk.Tk):
    COLUMNS = (  # id, heading, width, anchor
        ("sku", "SKU", 100, "w"), ("name", "Item name", 240, "w"), ("category", "Category", 140, "w"),
        ("quantity", "Qty", 70, "e"), ("price", "Price", 100, "e"), ("value", "Stock value", 120, "e"),
        ("min_stock", "Alert level", 90, "e"), ("status", "Status", 90, "center"),
    )
    NUMERIC = {"quantity", "price", "value", "min_stock"}

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.currency = db.get_setting("currency", "Rs. ")
        self.default_min = int(db.get_setting("default_min", "5"))
        self.sort_col, self.sort_rev = "name", False

        self.title(f"{APP_NAME}")
        self.geometry("1120x680")
        self.minsize(900, 520)
        self.configure(bg="#f3f5f9")
        self._style()
        self._menu()
        self._build()
        self._bind_keys()
        self.report_callback_exception = self._on_error
        self.refresh()

    # -- look & feel ------------------------------------------------------
    def _style(self):
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(name).configure(family="Segoe UI", size=10)
            except tk.TclError:
                pass
        s = ttk.Style(self)
        s.theme_use("clam")  # consistent, fully skinnable look on Windows 7 - 11
        bg, ink, muted, blue = "#f3f5f9", "#1f2937", "#6b7280", "#2563eb"
        s.configure(".", background=bg, foreground=ink, font=("Segoe UI", 10))
        s.configure("TFrame", background=bg)
        s.configure("TLabel", background=bg, foreground=ink)
        s.configure("Muted.TLabel", foreground=muted)
        s.configure("Heading.TLabel", font=("Segoe UI", 12, "bold"))
        s.configure("TButton", padding=(12, 6), background="#e5e9f0", borderwidth=1, relief="flat")
        s.map("TButton", background=[("active", "#d6dce8"), ("disabled", "#eceff4")],
              foreground=[("disabled", "#a0a7b4")])
        s.configure("Accent.TButton", background=blue, foreground="white")
        s.map("Accent.TButton", background=[("active", "#1d4ed8"), ("disabled", "#9db8f0")],
              foreground=[("disabled", "white")])
        s.configure("Danger.TButton", background="#fde2e2", foreground="#b02a37")
        s.map("Danger.TButton", background=[("active", "#f9c9c9"), ("disabled", "#f5eded")],
              foreground=[("disabled", "#d4a5a8")])
        s.configure("TEntry", padding=5, fieldbackground="white")
        s.configure("TCombobox", padding=5, fieldbackground="white")
        s.configure("Card.TFrame", background="white", relief="solid", borderwidth=1, bordercolor="#e2e6ee")
        s.configure("CardTitle.TLabel", background="white", foreground=muted, font=("Segoe UI", 9))
        s.configure("CardValue.TLabel", background="white", font=("Segoe UI", 18, "bold"))
        s.configure("Treeview", rowheight=28, background="white", fieldbackground="white", borderwidth=0)
        s.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e9edf5", padding=(8, 6),
                    relief="flat")
        s.map("Treeview.Heading", background=[("active", "#dbe2ef")])
        s.map("Treeview", background=[("selected", blue)], foreground=[("selected", "white")])
        s.configure("Status.TLabel", background="#e9edf5", padding=(10, 4), foreground=muted)

        # Work around the Tk 8.6.9 bug (bundled with Python 3.8) where Treeview
        # row tag colours are ignored.
        def fixed_map(option):
            return [e for e in s.map("Treeview", query_opt=option) if e[:2] != ("!disabled", "!selected")]
        s.map("Treeview", foreground=fixed_map("foreground"), background=fixed_map("background"))

    def _menu(self):
        bar = tk.Menu(self)
        f = tk.Menu(bar, tearoff=0)
        f.add_command(label="Import from CSV...", command=self.import_csv)
        f.add_command(label="Export to CSV...", accelerator="Ctrl+E", command=self.export_csv)
        f.add_separator()
        f.add_command(label="Back up database...", command=self.backup)
        f.add_separator()
        f.add_command(label="Exit", command=self.destroy)
        bar.add_cascade(label="File", menu=f)

        i = tk.Menu(bar, tearoff=0)
        i.add_command(label="Add item...", accelerator="Ctrl+N", command=self.add_item)
        i.add_command(label="Edit item...", accelerator="Enter", command=self.edit_item)
        i.add_command(label="Adjust stock...", accelerator="Ctrl+S", command=self.adjust_stock)
        i.add_command(label="Delete item", accelerator="Del", command=self.delete_item)
        bar.add_cascade(label="Items", menu=i)

        r = tk.Menu(bar, tearoff=0)
        r.add_command(label="Stock movement history", command=lambda: self.show_history(all_items=True))
        r.add_command(label="Low-stock report", command=self.show_low_stock)
        bar.add_cascade(label="Reports", menu=r)

        t = tk.Menu(bar, tearoff=0)
        t.add_command(label="Refresh", accelerator="F5", command=self.refresh)
        t.add_command(label="Settings...", command=self.open_settings)
        bar.add_cascade(label="Tools", menu=t)

        h = tk.Menu(bar, tearoff=0)
        h.add_command(label="About", command=self.about)
        bar.add_cascade(label="Help", menu=h)
        self.config(menu=bar)

        self.context = tk.Menu(self, tearoff=0)
        self.context.add_command(label="Edit item...", command=self.edit_item)
        self.context.add_command(label="Adjust stock...", command=self.adjust_stock)
        self.context.add_command(label="View history", command=self.show_history)
        self.context.add_separator()
        self.context.add_command(label="Delete item", command=self.delete_item)

    def _build(self):
        root = ttk.Frame(self, padding=(16, 14, 16, 0))
        root.pack(fill="both", expand=True)

        # Dashboard cards
        cards = ttk.Frame(root)
        cards.pack(fill="x")
        self.card_vars = {}
        for i, (key, title) in enumerate((("items", "TOTAL ITEMS"), ("units", "UNITS IN STOCK"),
                                          ("value", "STOCK VALUE"), ("low", "LOW / OUT OF STOCK"))):
            card = ttk.Frame(cards, style="Card.TFrame", padding=(16, 10))
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
            cards.columnconfigure(i, weight=1, uniform="cards")
            ttk.Label(card, text=title, style="CardTitle.TLabel").pack(anchor="w")
            var = tk.StringVar(value="-")
            ttk.Label(card, textvariable=var, style="CardValue.TLabel").pack(anchor="w")
            self.card_vars[key] = var

        # Toolbar
        tool = ttk.Frame(root)
        tool.pack(fill="x", pady=(14, 8))
        ttk.Button(tool, text="+  Add Item", style="Accent.TButton", command=self.add_item).pack(side="left")
        self.btn_edit = ttk.Button(tool, text="Edit", command=self.edit_item)
        self.btn_stock = ttk.Button(tool, text="Adjust Stock", command=self.adjust_stock)
        self.btn_hist = ttk.Button(tool, text="History", command=self.show_history)
        self.btn_del = ttk.Button(tool, text="Delete", style="Danger.TButton", command=self.delete_item)
        for b in (self.btn_edit, self.btn_stock, self.btn_hist, self.btn_del):
            b.pack(side="left", padx=(8, 0))

        ttk.Button(tool, text="Clear", command=self.clear_filters).pack(side="right")
        self.low_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(tool, text="Low stock only", variable=self.low_only,
                        command=self.refresh).pack(side="right", padx=12)
        self.cat_var = tk.StringVar(value="All categories")
        self.cat_box = ttk.Combobox(tool, textvariable=self.cat_var, state="readonly", width=18)
        self.cat_box.pack(side="right")
        self.cat_box.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self.refresh())
        self.search_entry = ttk.Entry(tool, textvariable=self.search_var, width=28)
        self.search_entry.pack(side="right", padx=(0, 10))
        ttk.Label(tool, text="Search:").pack(side="right", padx=(0, 6))

        # Table
        table = ttk.Frame(root)
        table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table, columns=[c[0] for c in self.COLUMNS], show="headings",
                                 selectmode="browse")
        for cid, text, width, anchor in self.COLUMNS:
            self.tree.heading(cid, text=text, command=lambda c=cid: self.sort_by(c))
            self.tree.column(cid, width=width, anchor=anchor, minwidth=60)
        self.tree.tag_configure("odd", background="#f8fafc")
        self.tree.tag_configure("low", background="#fff4d6", foreground="#7a5200")
        self.tree.tag_configure("out", background="#fde2e2", foreground="#9b1c1c")
        vs = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_buttons())
        self.tree.bind("<Double-1>", lambda e: self.edit_item() if self.tree.identify_row(e.y) else None)
        self.tree.bind("<Button-3>", self._popup)

        # Status bar
        self.status = tk.StringVar()
        ttk.Label(self, textvariable=self.status, style="Status.TLabel").pack(side="bottom", fill="x")

    def _bind_keys(self):
        self.bind("<Control-n>", lambda e: self.add_item())
        self.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), self.search_entry.select_range(0, "end")))
        self.bind("<Control-s>", lambda e: self.adjust_stock())
        self.bind("<Control-e>", lambda e: self.export_csv())
        self.bind("<F5>", lambda e: self.refresh())
        self.tree.bind("<Return>", lambda e: self.edit_item())
        self.tree.bind("<Delete>", lambda e: self.delete_item())

    # -- data display -----------------------------------------------------
    def money(self, v):
        return f"{self.currency}{v:,.2f}"

    def _selected_id(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def _update_buttons(self):
        state = "normal" if self._selected_id() else "disabled"
        for b in (self.btn_edit, self.btn_stock, self.btn_hist, self.btn_del):
            b.configure(state=state)

    def _sort_key(self, row):
        c = self.sort_col
        if c == "value":
            return row["quantity"] * row["price"]
        if c == "status":
            return row["quantity"] - row["min_stock"]
        if c in self.NUMERIC:
            return row[c]
        return str(row[c]).lower()

    def sort_by(self, col):
        self.sort_rev = (not self.sort_rev) if col == self.sort_col else False
        self.sort_col = col
        self.refresh()

    def refresh(self):
        cats = self.db.categories()
        self.cat_box.configure(values=["All categories"] + cats)
        if self.cat_var.get() not in cats:
            self.cat_var.set("All categories")
        category = "" if self.cat_var.get() == "All categories" else self.cat_var.get()

        keep = self._selected_id()
        rows = self.db.list_items(self.search_var.get().strip(), category, self.low_only.get())
        rows.sort(key=self._sort_key, reverse=self.sort_rev)

        self.tree.delete(*self.tree.get_children())
        for n, r in enumerate(rows):
            if r["quantity"] == 0:
                status, tag = "OUT", "out"
            elif r["quantity"] <= r["min_stock"]:
                status, tag = "LOW", "low"
            else:
                status, tag = "OK", ("odd" if n % 2 else "")
            self.tree.insert("", "end", iid=str(r["id"]), tags=(tag,), values=(
                r["sku"], r["name"], r["category"], r["quantity"], self.money(r["price"]),
                self.money(r["quantity"] * r["price"]), r["min_stock"], status))
        if keep and self.tree.exists(str(keep)):
            self.tree.selection_set(str(keep))
            self.tree.see(str(keep))

        for cid, text, *_ in self.COLUMNS:  # sort arrow on the active heading
            arrow = (" \u25BC" if self.sort_rev else " \u25B2") if cid == self.sort_col else ""
            self.tree.heading(cid, text=text + arrow)

        st = self.db.stats()
        self.card_vars["items"].set(f"{st['items']:,}")
        self.card_vars["units"].set(f"{st['units']:,}")
        self.card_vars["value"].set(self.money(st["value"]))
        self.card_vars["low"].set(f"{st['low']:,}")
        self.status.set(f"Showing {len(rows)} of {st['items']} items   |   "
                        f"Database: {self.db.path}")
        self._update_buttons()

    def clear_filters(self):
        self.search_var.set("")
        self.cat_var.set("All categories")
        self.low_only.set(False)
        self.refresh()

    def show_low_stock(self):
        self.low_only.set(True)
        self.refresh()

    def _popup(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.context.tk_popup(event.x_root, event.y_root)

    # -- actions ----------------------------------------------------------
    def add_item(self):
        data = ItemDialog(self, self.db.categories(), self.default_min).show()
        if data:
            try:
                new_id = self.db.add_item(data)
            except ValueError as e:
                messagebox.showerror("Cannot add item", str(e), parent=self)
                return
            self.refresh()
            self.tree.selection_set(str(new_id))
            self.tree.see(str(new_id))

    def edit_item(self):
        item_id = self._selected_id()
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
        item_id = self._selected_id()
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
        item_id = self._selected_id()
        if not item_id:
            return
        item = self.db.get(item_id)
        if item and messagebox.askyesno(
                "Delete item", f"Delete '{item['name']}' permanently?\nThis cannot be undone.",
                icon="warning", default="no", parent=self):
            self.db.delete_item(item_id)
            self.refresh()

    def show_history(self, all_items=False):
        item_id = None if all_items else self._selected_id()
        if not all_items and not item_id:
            return
        title = "Stock movement history" if all_items else f"History - {self.db.get(item_id)['name']}"
        HistoryDialog(self, self.db.movements(item_id), title).show()

    def open_settings(self):
        result = SettingsDialog(self, self.currency, self.default_min).show()
        if result:
            self.currency, self.default_min = result
            self.db.set_setting("currency", self.currency)
            self.db.set_setting("default_min", str(self.default_min))
            self.refresh()

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
        # Bring over an inventory.db from the old console version, if present.
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
