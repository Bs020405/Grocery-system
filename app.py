"""
Grocery Store Management System (Excel version)
- No database. All data is saved automatically in grocery_data.xlsx (openpyxl).
- Pages: Dashboard, Billing, Products
Run:  pip install flask openpyxl
      python app.py
Open: http://127.0.0.1:5000      (login: admin / admin123)
"""
import copy
import hmac
import os
import threading
from collections import Counter
from contextlib import contextmanager
from datetime import datetime, date, timedelta

from flask import Flask, render_template, request, redirect, url_for, flash, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = "change-this-secret"
app.config["SESSION_COOKIE_NAME"] = "grocery_session"   # own cookie name, so other apps on localhost cannot clash

# Data file (created in the same folder as app.py)
EXCEL_FILE = os.environ.get("GROCERY_EXCEL") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "grocery_data.xlsx")

# Login details (change as you like)
ADMIN_USER = os.environ.get("GROCERY_USER", "admin")
ADMIN_PASS = os.environ.get("GROCERY_PASS", "admin123")
STORE_NAME = "FreshMart"   # your store name


# ---------------------------------------------------------------- excel data store
# Har sheet: (data key, [(field, Excel column heading), ...])
SHEETS = {
    "Products": ("products", [("id", "ID"), ("name", "Name"), ("category", "Category"), ("unit", "Unit"),
                              ("cost", "Cost"), ("price", "Price"), ("gst", "GST %"), ("stock", "Stock"),
                              ("min_stock", "Min stock"), ("expiry", "Expiry")]),
    "Sales": ("sales", [("id", "Bill"), ("created", "Date"), ("customer", "Customer"), ("phone", "Phone"),
                        ("subtotal", "Subtotal"), ("gst", "GST"), ("discount", "Discount"),
                        ("total", "Total"), ("payment", "Payment"), ("paid", "Paid")]),
    "Sale Items": ("sale_items", [("sale_id", "Bill"), ("product_id", "Product ID"), ("name", "Product"),
                                  ("unit", "Unit"), ("qty", "Qty"), ("price", "Price"), ("cost", "Cost"),
                                  ("gst_rate", "GST %")]),
    "Payments": ("payments", [("id", "Payment"), ("created", "Date"), ("customer", "Customer"),
                              ("phone", "Phone"), ("amount", "Amount"), ("mode", "Mode"),
                              ("note", "Note")]),
}
NUMERIC = {"id", "cost", "price", "gst", "stock", "min_stock", "subtotal", "discount", "total",
           "qty", "sale_id", "product_id", "gst_rate", "paid", "amount"}
INT_FIELDS = {"id", "sale_id", "product_id"}

_lock = threading.RLock()
_cache = {"mtime": None, "data": None}


def num(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _clean(field, v):
    if field in NUMERIC:
        n = num(v)
        return int(n) if field in INT_FIELDS else n
    if field == "phone" and isinstance(v, float) and v.is_integer():
        v = int(v)
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):          # if Excel turned it into a date, convert it back to text
        return v.strftime("%Y-%m-%d %H:%M" if field == "created" else "%Y-%m-%d")
    return str(v).strip()


def _read_file():
    data = {key: [] for key, _ in SHEETS.values()}
    if not os.path.exists(EXCEL_FILE):
        return data
    wb = load_workbook(EXCEL_FILE, data_only=True)
    for sheet, (key, cols) in SHEETS.items():
        if sheet not in wb.sheetnames:
            continue
        rows = wb[sheet].iter_rows(values_only=True)
        head = next(rows, ()) or ()
        pos = {str(h).strip().lower(): i for i, h in enumerate(head) if h is not None}
        for row in rows:
            if not any(c not in (None, "") for c in row):
                continue
            rec = {}
            for field, label in cols:
                i = pos.get(label.lower())
                rec[field] = _clean(field, row[i] if i is not None and i < len(row) else None)
            if key == "sales":                       # old files have no Paid column: treat as fully paid
                pi = pos.get("paid")
                raw = row[pi] if pi is not None and pi < len(row) else None
                if raw in (None, ""):
                    rec["paid"] = rec["total"]
            data[key].append(rec)
    wb.close()
    nid = next_id(data["products"])             # give an ID to products typed by hand in Excel
    for p in data["products"]:
        if not p["id"]:
            p["id"], nid = nid, nid + 1
    return data


def load():
    """Reads the data from Excel (re-reads only when the file has changed)."""
    with _lock:
        mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else None
        if _cache["data"] is None or _cache["mtime"] != mtime:
            _cache["data"], _cache["mtime"] = _read_file(), mtime
        return copy.deepcopy(_cache["data"])


CUSTOMER_BILL_HEADERS = ["Bill No", "Date", "Customer", "Phone", "Payment", "Item", "Unit", "Qty", "Rate",
                         "GST %", "Item Amount", "Bill Subtotal", "Bill GST", "Discount", "Bill Total",
                         "Bill Paid", "Bill Pending"]


def customer_bill_rows(data):
    """One row per item of every customer bill (bill totals are written on the first row of each bill)."""
    by_bill = {}
    for it in data["sale_items"]:
        by_bill.setdefault(it["sale_id"], []).append(it)
    rows = []
    for s in sorted(data["sales"], key=lambda s: s["id"]):
        for n, it in enumerate(by_bill.get(s["id"]) or [None]):
            first = n == 0
            rows.append([s["id"], s["created"], s["customer"], s["phone"], s["payment"],
                         it["name"] if it else "", it["unit"] if it else "", it["qty"] if it else "",
                         it["price"] if it else "", it["gst_rate"] if it else "",
                         round(it["qty"] * it["price"], 2) if it else "",
                         s["subtotal"] if first else None, s["gst"] if first else None,
                         s["discount"] if first else None, s["total"] if first else None,
                         s.get("paid", s["total"]) if first else None,
                         round(s["total"] - s.get("paid", s["total"]), 2) if first else None])
    return rows


def save(data):
    """Writes all the data to the Excel file."""
    with _lock:
        wb = Workbook()
        wb.remove(wb.active)
        head_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        head_fill = PatternFill("solid", fgColor="6D28D9")
        body_font = Font(name="Arial", size=10)
        # Customer Bills is a readable log (written only); the other sheets are the real data store
        tables = [("Customer Bills", CUSTOMER_BILL_HEADERS, customer_bill_rows(data))]
        for sheet, (key, cols) in SHEETS.items():
            tables.append((sheet, [label for _, label in cols],
                           [[rec.get(field, "") for field, _ in cols] for rec in data[key]]))
        for title, headers, rows in tables:
            ws = wb.create_sheet(title)
            ws.append(headers)
            for r in rows:
                ws.append(r)
            for c in ws[1]:
                c.font, c.fill, c.alignment = head_font, head_fill, Alignment(horizontal="center")
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.font = body_font
            for i, h in enumerate(headers, 1):
                longest = max([len(h)] + [len(str(r[i - 1])) for r in rows[:200] if r[i - 1] is not None])
                ws.column_dimensions[get_column_letter(i)].width = min(max(longest + 3, 9), 38)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"
        tmp = EXCEL_FILE + ".tmp"
        wb.save(tmp)
        try:
            os.replace(tmp, EXCEL_FILE)
        except OSError:
            if os.path.exists(tmp):
                os.remove(tmp)
            raise
        _cache["data"], _cache["mtime"] = copy.deepcopy(data), os.path.getmtime(EXCEL_FILE)


@contextmanager
def editing():
    """with editing() as d: ... change d; it is saved to Excel when the block ends."""
    with _lock:
        data = load()
        before = copy.deepcopy(data)
        yield data
        if data != before:
            save(data)


def next_id(rows):
    return max([r["id"] for r in rows] + [0]) + 1


def init_store():
    if not os.path.exists(EXCEL_FILE):
        save({key: [] for key, _ in SHEETS.values()})


@app.errorhandler(PermissionError)
def excel_busy(_):
    flash("grocery_data.xlsx is open in Excel. Close it and try again (your change was not saved)", "error")
    return redirect(request.referrer or url_for("dashboard"))


@app.context_processor
def inject_dates():
    t = date.today()
    return {"today": t.isoformat(), "soon": (t + timedelta(days=30)).isoformat(),
            "store_name": STORE_NAME}


# ---------------------------------------------------------------- login
@app.before_request
def require_login():
    if request.endpoint in ("login", "static") or session.get("user"):
        return None
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u, p = request.form.get("username", ""), request.form.get("password", "")
        if hmac.compare_digest(u, ADMIN_USER) and hmac.compare_digest(p, ADMIN_PASS):
            session["user"] = u
            return redirect(url_for("dashboard"))
        flash("Wrong username or password", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------- dashboard
EMOJI = [("milk", "🥛"), ("butter", "🧈"), ("paneer", "🧀"), ("cheese", "🧀"), ("curd", "🥣"),
         ("egg", "🥚"), ("banana", "🍌"), ("apple", "🍎"), ("tomato", "🍅"), ("onion", "🧅"),
         ("potato", "🥔"), ("mango", "🥭"), ("rice", "🍚"), ("atta", "🌾"), ("flour", "🌾"),
         ("dal", "🫘"), ("salt", "🧂"), ("sugar", "🍬"), ("oil", "🫒"), ("bread", "🍞"),
         ("bakery", "🥐"), ("chips", "🍟"), ("noodle", "🍜"), ("maggi", "🍜"), ("biscuit", "🍪"),
         ("parle", "🍪"), ("cola", "🥤"), ("juice", "🧃"), ("tea", "🍵"), ("coffee", "☕"),
         ("water", "💧"), ("beverage", "🥤"), ("soap", "🧼"), ("surf", "🧺"), ("detergent", "🧺"),
         ("shampoo", "🧴"), ("personal", "🧴"), ("household", "🧹"), ("snack", "🍿"),
         ("dairy", "🥛"), ("fruit", "🍎"), ("veg", "🥬"), ("staple", "🌾"), ("grocery", "🛒")]


@app.template_global()
def emoji(name="", category=""):
    text = f"{name} {category}".lower()
    for key, e in EMOJI:
        if key in text:
            return e
    return "🛒"


@app.route("/")
def dashboard():
    d = load()
    prods, sales, items = d["products"], d["sales"], d["sale_items"]
    t = date.today()
    today_s, soon = t.isoformat(), (t + timedelta(days=30)).isoformat()
    todays = [s for s in sales if s["created"].startswith(today_s)]
    stats = {
        "products": len(prods),
        "today_sales": sum(s["total"] for s in todays),
        "today_bills": len(todays),
        "total_sales": sum(s["total"] for s in sales),
        "stock_value": sum(p["stock"] * p["cost"] for p in prods),
    }
    stats["khata_pending"] = sum(max(k["balance"], 0) for k in build_khata(d).values())
    low = sorted((p for p in prods if p["stock"] <= p["min_stock"]), key=lambda p: p["stock"])
    expiring = sorted((p for p in prods if p["expiry"] and p["expiry"] <= soon and p["stock"] > 0),
                      key=lambda p: p["expiry"])
    cats = [{"cat": c, "n": n} for c, n in
            Counter(p["category"] or "Other" for p in prods).most_common(12)]
    sold = Counter()
    for it in items:
        sold[it["product_id"]] += it["qty"]
    byid = {p["id"]: p for p in prods}
    top = [dict(byid[pid], sold=q) for pid, q in sold.most_common() if pid in byid][:8]
    if not top:
        top = [dict(p, sold=0) for p in sorted(prods, key=lambda p: -p["id"])[:8]]
    sums = Counter()
    for s in sales:
        sums[s["created"][:10]] += s["total"]
    week = []
    for i in range(6, -1, -1):
        day = t - timedelta(days=i)
        week.append({"label": day.strftime("%a"), "amount": sums.get(day.isoformat(), 0)})
    peak = max([w["amount"] for w in week] + [1])
    for w in week:
        w["pct"] = max(3, int(w["amount"] * 100 / peak)) if w["amount"] else 3
    recent = sorted(sales, key=lambda s: -s["id"])[:5]
    return render_template("dashboard.html", stats=stats, low=low, expiring=expiring, cats=cats,
                           top=top, week=week, week_total=sum(w["amount"] for w in week),
                           recent=recent)


DEMO_PRODUCTS = [  # name, category, unit, cost, price, gst, stock, min_stock, expiry (days from today)
    # Dairy
    ("Amul Milk 500ml", "Dairy", "pack", 24, 28, 0, 40, 10, 3),
    ("Amul Taaza Milk 1L", "Dairy", "pack", 50, 56, 0, 30, 10, 3),
    ("Amul Butter 100g", "Dairy", "pack", 50, 58, 12, 25, 8, 60),
    ("Amul Butter 500g", "Dairy", "pack", 250, 285, 12, 10, 3, 60),
    ("Paneer 200g", "Dairy", "pack", 70, 85, 5, 12, 5, 5),
    ("Amul Cheese Slices 200g", "Dairy", "pack", 120, 140, 12, 15, 5, 90),
    ("Curd 400g", "Dairy", "pack", 28, 35, 0, 25, 8, 5),
    ("Eggs (12 pcs)", "Dairy", "tray", 70, 84, 0, 20, 6, 10),
    # Staples
    ("Aashirvaad Atta 5kg", "Staples", "bag", 220, 260, 0, 18, 5, None),
    ("Aashirvaad Atta 10kg", "Staples", "bag", 420, 485, 0, 10, 3, None),
    ("Basmati Rice 1kg", "Staples", "kg", 85, 110, 0, 30, 8, None),
    ("Sona Masoori Rice 5kg", "Staples", "bag", 280, 330, 0, 12, 4, None),
    ("Toor Dal 1kg", "Staples", "kg", 130, 160, 0, 4, 6, None),
    ("Moong Dal 1kg", "Staples", "kg", 110, 135, 0, 20, 6, None),
    ("Chana Dal 1kg", "Staples", "kg", 85, 105, 0, 22, 6, None),
    ("Urad Dal 1kg", "Staples", "kg", 120, 150, 0, 15, 5, None),
    ("Rajma 1kg", "Staples", "kg", 120, 150, 0, 14, 5, None),
    ("Kabuli Chana 1kg", "Staples", "kg", 110, 140, 0, 16, 5, None),
    ("Sugar 1kg", "Staples", "kg", 40, 47, 0, 50, 10, None),
    ("Tata Salt 1kg", "Staples", "kg", 20, 28, 0, 60, 10, None),
    ("Poha 500g", "Staples", "pack", 28, 38, 5, 25, 6, 180),
    ("Besan 500g", "Staples", "pack", 45, 60, 5, 20, 6, 180),
    ("Sooji 500g", "Staples", "pack", 30, 40, 5, 20, 6, 180),
    ("Maida 1kg", "Staples", "kg", 36, 46, 0, 20, 6, None),
    # Oil & Ghee
    ("Fortune Sunflower Oil 1L", "Oil & Ghee", "pouch", 125, 145, 5, 30, 8, 240),
    ("Fortune Mustard Oil 1L", "Oil & Ghee", "bottle", 150, 175, 5, 20, 6, 240),
    ("Saffola Gold Oil 1L", "Oil & Ghee", "pouch", 160, 185, 5, 15, 5, 240),
    ("Amul Ghee 1L", "Oil & Ghee", "pack", 560, 620, 12, 12, 4, 180),
    # Spices
    ("Everest Turmeric Powder 100g", "Spices", "pack", 22, 30, 5, 30, 8, 300),
    ("Everest Chilli Powder 100g", "Spices", "pack", 35, 46, 5, 30, 8, 300),
    ("Everest Garam Masala 50g", "Spices", "pack", 40, 52, 5, 25, 8, 300),
    ("Coriander Powder 100g", "Spices", "pack", 24, 32, 5, 25, 8, 300),
    ("MDH Chana Masala 100g", "Spices", "pack", 55, 70, 5, 20, 6, 300),
    ("Jeera (Cumin) 100g", "Spices", "pack", 45, 60, 5, 20, 6, 300),
    # Fruits & Veg
    ("Banana", "Fruits & Veg", "dozen", 35, 50, 0, 20, 5, 3),
    ("Apple", "Fruits & Veg", "kg", 140, 180, 0, 15, 5, 7),
    ("Tomato", "Fruits & Veg", "kg", 20, 32, 0, 25, 8, 4),
    ("Onion", "Fruits & Veg", "kg", 28, 40, 0, 50, 10, 20),
    ("Potato", "Fruits & Veg", "kg", 18, 28, 0, 60, 15, 25),
    ("Carrot", "Fruits & Veg", "kg", 30, 45, 0, 15, 5, 6),
    ("Cauliflower", "Fruits & Veg", "kg", 25, 40, 0, 12, 4, 4),
    ("Green Chilli 250g", "Fruits & Veg", "pack", 10, 15, 0, 15, 5, 4),
    ("Ginger 100g", "Fruits & Veg", "pack", 8, 12, 0, 20, 6, 10),
    ("Garlic 250g", "Fruits & Veg", "pack", 25, 38, 0, 15, 5, 20),
    ("Coriander Bunch", "Fruits & Veg", "bunch", 8, 12, 0, 20, 6, 2),
    # Snacks
    ("Lays Chips 52g", "Snacks", "pack", 16, 20, 12, 80, 15, 90),
    ("Kurkure 40g", "Snacks", "pack", 8, 10, 12, 80, 15, 90),
    ("Parle-G Biscuit", "Snacks", "pack", 8, 10, 12, 100, 20, 120),
    ("Britannia Good Day 100g", "Snacks", "pack", 25, 30, 12, 40, 10, 120),
    ("Hide & Seek 100g", "Snacks", "pack", 30, 35, 12, 30, 8, 120),
    ("Maggi Noodles 70g", "Snacks", "pack", 11, 14, 12, 60, 12, 150),
    ("Haldiram Aloo Bhujia 200g", "Snacks", "pack", 55, 65, 12, 25, 8, 120),
    ("Cadbury Dairy Milk 40g", "Snacks", "pcs", 36, 45, 18, 40, 10, 150),
    # Beverages
    ("Coca-Cola 750ml", "Beverages", "btl", 32, 40, 28, 35, 10, 180),
    ("Pepsi 750ml", "Beverages", "btl", 32, 40, 28, 30, 10, 180),
    ("Maaza 600ml", "Beverages", "btl", 32, 40, 12, 30, 10, 150),
    ("Real Fruit Juice 1L", "Beverages", "pack", 85, 105, 12, 15, 5, 120),
    ("Bisleri Water 1L", "Beverages", "btl", 12, 20, 18, 60, 15, 300),
    ("Tata Tea Gold 250g", "Beverages", "pack", 115, 135, 5, 20, 6, 300),
    ("Red Label Tea 250g", "Beverages", "pack", 100, 120, 5, 20, 6, 300),
    ("Nescafe Classic 50g", "Beverages", "jar", 150, 175, 12, 12, 4, 360),
    # Bakery & Breakfast
    ("Bread", "Bakery", "pack", 28, 40, 0, 15, 5, 1),
    ("Pav (6 pcs)", "Bakery", "pack", 18, 25, 0, 15, 5, 1),
    ("Kissan Mixed Fruit Jam 200g", "Breakfast", "jar", 70, 85, 12, 15, 5, 240),
    ("Saffola Oats 400g", "Breakfast", "pack", 85, 105, 12, 15, 5, 240),
    ("Kellogg's Corn Flakes 250g", "Breakfast", "pack", 90, 110, 12, 12, 4, 240),
    # Household
    ("Surf Excel 1kg", "Household", "pack", 190, 230, 18, 10, 4, None),
    ("Vim Dishwash Gel 500ml", "Household", "btl", 95, 115, 18, 12, 4, None),
    ("Harpic Toilet Cleaner 500ml", "Household", "btl", 85, 105, 18, 10, 4, None),
    ("Lizol Floor Cleaner 500ml", "Household", "btl", 95, 115, 18, 10, 4, None),
    ("Matchbox (10 pcs)", "Household", "pack", 8, 12, 12, 30, 8, None),
    # Personal Care
    ("Dettol Soap", "Personal Care", "pcs", 28, 38, 18, 3, 5, None),
    ("Lux Soap 100g", "Personal Care", "pcs", 28, 35, 18, 30, 8, None),
    ("Dove Soap 100g", "Personal Care", "pcs", 50, 62, 18, 20, 6, None),
    ("Colgate Toothpaste 100g", "Personal Care", "pcs", 50, 62, 18, 25, 8, None),
    ("Clinic Plus Shampoo 340ml", "Personal Care", "btl", 150, 180, 18, 10, 4, None),
    ("Vaseline Body Lotion 100ml", "Personal Care", "btl", 85, 105, 18, 8, 3, None),
]


@app.route("/demo", methods=["POST"])
def demo():
    """Adds sample products (names that already exist are skipped)."""
    added = 0
    with editing() as d:
        existing = {p["name"].lower() for p in d["products"]}
        for n, c, u, cost, price, gst, stock_, mn, days in DEMO_PRODUCTS:
            if n.lower() in existing:
                continue
            exp = (date.today() + timedelta(days=days)).isoformat() if days is not None else ""
            d["products"].append({"id": next_id(d["products"]), "name": n, "category": c, "unit": u,
                                  "cost": cost, "price": price, "gst": gst, "stock": stock_,
                                  "min_stock": mn, "expiry": exp})
            added += 1
    flash(f"{added} sample products added" if added else "All sample products already exist", "ok")
    return redirect(url_for("products"))


# ---------------------------------------------------------------- products
def render_products(rows, q="", edit=None):
    cats = sorted({p["category"] for p in load()["products"] if p["category"]})
    return render_template("products.html", rows=rows, q=q, edit=edit, cats=cats)


@app.route("/products")
def products():
    q = request.args.get("q", "").strip().lower()
    rows = [p for p in load()["products"]
            if not q or q in p["name"].lower() or q in p["category"].lower()]
    return render_products(sorted(rows, key=lambda p: p["name"].lower()), q)


@app.route("/products/save", methods=["POST"])
def product_save():
    f = request.form
    name = f.get("name", "").strip()
    if not name:
        flash("Product name required", "error")
        return redirect(url_for("products"))
    rec = {"name": name, "category": f.get("category", "").strip(), "unit": f.get("unit", "pcs").strip() or "pcs",
           "cost": num(f.get("cost")), "price": num(f.get("price")), "gst": num(f.get("gst")),
           "stock": num(f.get("stock")), "min_stock": num(f.get("min_stock"), 5),
           "expiry": f.get("expiry", "").strip()}
    pid = f.get("id", type=int)
    with editing() as d:
        if any(p["name"].lower() == name.lower() and p["id"] != pid for p in d["products"]):
            flash("A product with this name already exists", "error")
            return redirect(url_for("products"))
        if pid:
            p = next((p for p in d["products"] if p["id"] == pid), None)
            if p is None:
                flash("Product not found", "error")
                return redirect(url_for("products"))
            p.update(rec)
            flash("Product updated", "ok")
        else:
            d["products"].append(dict(rec, id=next_id(d["products"])))
            flash("Product added", "ok")
    return redirect(url_for("products"))


@app.route("/products/<int:pid>/edit")
def product_edit(pid):
    d = load()
    p = next((p for p in d["products"] if p["id"] == pid), None)
    return render_products(sorted(d["products"], key=lambda p: p["name"].lower()), "", p)


@app.route("/products/<int:pid>/delete", methods=["GET", "POST"])
def product_delete(pid):
    if request.method == "GET":                      # first show an "are you sure?" page
        p = next((p for p in load()["products"] if p["id"] == pid), None)
        if p is None:
            return redirect(url_for("products"))
        return render_template("delete.html", p=p)
    with editing() as d:
        if any(i["product_id"] == pid for i in d["sale_items"]):
            flash("This product has sales records, so it cannot be deleted", "error")
        else:
            d["products"] = [p for p in d["products"] if p["id"] != pid]
            flash("Product deleted", "ok")
    return redirect(url_for("products"))


# ---------------------------------------------------------------- billing (no JavaScript)
# The cart lives in the session; after every "Add to bill" the page reloads and Python works out the totals.
def product_options(d):
    """All products (sorted by category, then name). Out of stock / expired ones are marked."""
    today_s = date.today().isoformat()
    out = []
    for p in sorted(d["products"], key=lambda p: ((p["category"] or "Other").lower(), p["name"].lower())):
        status = "Out of stock" if p["stock"] <= 0 else (
            "Expired" if p["expiry"] and p["expiry"] < today_s else "")
        out.append(dict(p, category=p["category"] or "Other", status=status))
    return out


def get_cart():
    """Cart from the session: keeps only valid [product_id, qty] pairs and ignores stale or odd data."""
    raw, cart = session.get("cart", []), []
    if isinstance(raw, list):
        for entry in raw:
            if isinstance(entry, (list, tuple)) and len(entry) == 2:
                try:
                    pid, qty = int(entry[0]), float(entry[1])
                except (TypeError, ValueError):
                    continue
                if qty > 0:
                    cart.append([pid, qty])
    return cart


def read_cart(d):
    byid = {p["id"]: p for p in d["products"]}
    lines = []
    for pid, qty in get_cart():
        p = byid.get(pid)
        if p:
            lines.append({"id": pid, "name": p["name"], "unit": p["unit"], "qty": qty, "price": p["price"],
                          "gst": p["gst"], "amount": round(p["price"] * qty, 2)})
    subtotal = sum(l["price"] * l["qty"] for l in lines)
    gst = sum(l["price"] * l["qty"] * l["gst"] / 100 for l in lines)
    return lines, subtotal, gst


@app.route("/billing")
def billing():
    d = load()
    lines, subtotal, gst = read_cart(d)
    q = request.args.get("q", "").strip()                 # search bar text
    everything = product_options(d)
    shown = [p for p in everything
             if not q or q.lower() in p["name"].lower() or q.lower() in p["category"].lower()]
    return render_template("billing.html", prods=shown, all_count=len(everything), q=q,
                           pre=request.args.get("add", type=int),    # e.g. from the dashboard "ADD" card
                           lines=lines, subtotal=subtotal, gst=gst, total=subtotal + gst,
                           owed=owed_map(d))


@app.route("/billing/add", methods=["POST"])
def billing_add():
    q = request.form.get("q", "").strip()                 # keep the search after adding an item

    def back():
        return redirect(url_for("billing", q=q) if q else url_for("billing"))

    pid = request.form.get("product_id", type=int)
    qty = num(request.form.get("qty"))
    if not pid or qty <= 0:
        flash("Choose an item and enter a quantity", "error")
        return back()
    p = next((p for p in load()["products"] if p["id"] == pid), None)
    if p is None:
        flash("Product not found", "error")
    elif p["expiry"] and p["expiry"] < date.today().isoformat():
        flash(f"'{p['name']}' has expired ({p['expiry']}), so it cannot be billed", "error")
    else:
        cart = get_cart()
        have = next((x for x in cart if x[0] == pid), None)
        total_q = qty + (have[1] if have else 0)
        if total_q > p["stock"]:
            flash(f"Not enough stock for '{p['name']}' (available: {p['stock']:g})", "error")
        else:
            if have:
                have[1] = total_q
            else:
                cart.append([pid, qty])
            session["cart"] = cart
    return back()


@app.route("/billing/remove/<int:pid>", methods=["POST"])
def billing_remove(pid):
    session["cart"] = [x for x in get_cart() if x[0] != pid]
    return redirect(url_for("billing"))


@app.route("/billing/clear", methods=["POST"])
def billing_clear():
    session.pop("cart", None)
    return redirect(url_for("billing"))


@app.route("/billing/checkout", methods=["POST"])
def billing_checkout():
    cart = get_cart()
    f = request.form
    if not cart:
        flash("Add at least one item", "error")
        return redirect(url_for("billing"))
    payment = f.get("payment", "Cash")
    if payment == "Udhar" and (not f.get("customer", "").strip() or len(digits(f.get("phone"))) < 10):
        flash("For Udhar please enter the customer name and a 10-digit phone number", "error")
        return redirect(url_for("billing"))
    with editing() as d:
        byid = {p["id"]: p for p in d["products"]}
        chosen = []
        for pid, q in cart:
            p = byid.get(pid)
            if p is None:
                continue
            if p["expiry"] and p["expiry"] < date.today().isoformat():
                flash(f"'{p['name']}' has expired ({p['expiry']}), so it cannot be billed", "error")
                return redirect(url_for("billing"))
            if p["stock"] < q:
                flash(f"Not enough stock for '{p['name']}' (available: {p['stock']:g})", "error")
                return redirect(url_for("billing"))
            chosen.append((p, q))
        if not chosen:
            flash("Add at least one item", "error")
            return redirect(url_for("billing"))
        subtotal = sum(p["price"] * q for p, q in chosen)
        gst = sum(p["price"] * q * p["gst"] / 100 for p, q in chosen)
        discount = max(0.0, min(num(f.get("discount")), subtotal + gst))
        total = round(subtotal + gst - discount, 2)
        paid = round(min(max(num(f.get("paid")), 0.0), total), 2) if payment == "Udhar" else total
        sid = next_id(d["sales"])
        d["sales"].append({"id": sid, "created": now(),
                           "customer": f.get("customer", "").strip() or "Walk-in",
                           "phone": f.get("phone", "").strip(), "subtotal": round(subtotal, 2),
                           "gst": round(gst, 2), "discount": discount,
                           "total": total, "payment": payment, "paid": paid})
        for p, q in chosen:
            d["sale_items"].append({"sale_id": sid, "product_id": p["id"], "name": p["name"],
                                    "unit": p["unit"], "qty": q, "price": p["price"],
                                    "cost": p["cost"], "gst_rate": p["gst"]})
            p["stock"] -= q
    session.pop("cart", None)
    flash(f"Bill #{sid} saved to Excel (grocery_data.xlsx)", "ok")
    return redirect(url_for("bill", sid=sid))


@app.route("/bill/<int:sid>")
def bill(sid):
    d = load()
    sale = next((s for s in d["sales"] if s["id"] == sid), None)
    if sale is None:
        return redirect(url_for("dashboard"))
    kh, ph = None, digits(sale["phone"])
    book = build_khata(d).get(ph) if ph else None
    e = next((e for e in book["entries"] if e["kind"] == "bill" and e["id"] == sid), None) if book else None
    if e:
        kh = {"phone": ph, "paid": sale["paid"], "pending": round(sale["total"] - sale["paid"], 2),
              "prev": max(e["before"], 0), "after": max(e["after"], 0)}
    return render_template("bill.html", sale=sale, kh=kh,
                           items=[i for i in d["sale_items"] if i["sale_id"] == sid])


# ---------------------------------------------------------------- khata (udhar / credit)
def digits(phone):
    """Phone number as digits only (last 10), so '+91 98765 43210' and '9876543210' are the same customer."""
    s = "".join(ch for ch in str(phone or "") if ch.isdigit())
    return s[-10:] if len(s) > 10 else s


def build_khata(d):
    """One ledger per customer phone.
    Every bill adds its Total (debit) and what was paid at billing time (credit);
    every payment received later is a credit. Balance = debit - credit."""
    books = {}

    def book(ph):
        return books.setdefault(ph, {"phone": ph, "name": "", "entries": [],
                                     "billed": 0.0, "received": 0.0, "balance": 0.0})

    for s in d["sales"]:
        ph = digits(s["phone"])
        if not ph:
            continue
        b = book(ph)
        if s["customer"] and s["customer"] != "Walk-in":
            b["name"] = s["customer"]
        b["entries"].append({"kind": "bill", "id": s["id"], "created": s["created"],
                             "text": f"Bill #{s['id']} ({s['payment'] or 'Cash'})",
                             "debit": s["total"], "credit": s.get("paid", s["total"])})
    for p in d["payments"]:
        ph = digits(p["phone"])
        if not ph:
            continue
        b = book(ph)
        if not b["name"] and p["customer"]:
            b["name"] = p["customer"]
        note = f" - {p['note']}" if p["note"] else ""
        b["entries"].append({"kind": "payment", "id": p["id"], "created": p["created"],
                             "text": f"Payment received ({p['mode'] or 'Cash'}){note}",
                             "debit": 0.0, "credit": p["amount"]})
    for b in books.values():
        b["entries"].sort(key=lambda e: (e["created"], 0 if e["kind"] == "bill" else 1, e["id"]))
        bal = 0.0
        for e in b["entries"]:
            e["before"] = round(bal, 2)
            bal += e["debit"] - e["credit"]
            e["after"] = round(bal, 2)
        b["balance"] = round(bal, 2)
        b["billed"] = round(sum(e["debit"] for e in b["entries"]), 2)
        b["received"] = round(sum(e["credit"] for e in b["entries"]), 2)
    return books


def owed_map(d):
    """{phone: {name, balance}} for customers who still owe money (used on the billing page)."""
    return {ph: {"name": b["name"], "balance": b["balance"]}
            for ph, b in build_khata(d).items() if b["balance"] > 0.005}


@app.route("/khata")
def khata():
    d = load()
    q = request.args.get("q", "").strip().lower()
    show_all = request.args.get("all") == "1"
    books = sorted(build_khata(d).values(), key=lambda b: (-b["balance"], b["name"].lower()))
    pending = [b for b in books if b["balance"] > 0.005]
    rows = [b for b in books
            if (show_all or q or b["balance"] > 0.005)
            and (not q or q in b["name"].lower() or q in b["phone"])]
    return render_template("khata.html", rows=rows, q=q, show_all=show_all, n_pending=len(pending),
                           total_pending=round(sum(b["balance"] for b in pending), 2))


@app.route("/khata/<phone>")
def khata_customer(phone):
    b = build_khata(load()).get(digits(phone))
    if b is None:
        flash("Customer not found in khata", "error")
        return redirect(url_for("khata"))
    return render_template("khata_customer.html", b=b)


@app.route("/khata/<phone>/pay", methods=["POST"])
def khata_pay(phone):
    ph = digits(phone)
    amount = round(num(request.form.get("amount")), 2)
    with editing() as d:
        b = build_khata(d).get(ph)
        if b is None:
            flash("Customer not found in khata", "error")
            return redirect(url_for("khata"))
        if amount <= 0:
            flash("Enter the amount received", "error")
        elif amount > b["balance"] + 0.005:
            flash(f"Amount is more than the pending balance (₹{b['balance']:.2f})", "error")
        else:
            d["payments"].append({"id": next_id(d["payments"]), "created": now(), "customer": b["name"],
                                  "phone": ph, "amount": amount,
                                  "mode": request.form.get("mode", "Cash"),
                                  "note": request.form.get("note", "").strip()})
            flash(f"₹{amount:.2f} received from {b['name'] or ph}", "ok")
    return redirect(url_for("khata_customer", phone=ph))


init_store()

if __name__ == "__main__":
    app.run(debug=True)