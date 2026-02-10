import re
import shutil
from io import BytesIO
from pathlib import Path
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
from emailer import send_owner_snapshot_email
# Charts (for PDFs)
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# PDF rendering
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
    PageBreak,
    KeepTogether,
    Flowable,
)

# Optional nicer fonts
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- IMPORTANT: uses your existing exporter ---
from getSalesReport import run_sales_report, store_abbr_map  # store_name -> "MV"

###############################################################################
# CONFIG (easy to change)
###############################################################################

REPORT_TZ = "America/Los_Angeles"

# Backfill window used only when RUN_EXPORT=True
BACKFILL_DAYS = 61

REPORTS_ROOT = Path("reports").resolve()
RAW_ROOT = REPORTS_ROOT / "raw_sales"
PDF_ROOT = REPORTS_ROOT / "pdf"

# If True: run Selenium export and archive fresh files
# If False: reuse latest RAW folder, do NOT run Selenium
RUN_EXPORT = False

# If RUN_EXPORT=True: delete existing /files downloads first?
CLEANUP_FILES_BEFORE_EXPORT = False

# If RUN_EXPORT=True: do you want to "move" files out of /files, or "copy" them?
ARCHIVE_ACTION = "move"  # "move" or "copy"

# Build combined PDF summary as well as per-store PDFs
GENERATE_ALL_STORES_SUMMARY_PDF = True

# Charts / tables
TREND_DAYS = 14
TOP_N = 20
CATEGORY_TOP_N = 10

# --- Dutchie export header row ---
FORCE_HEADER_ROW = True
EXPORT_HEADER_ROW_INDEX = 4  # Excel row 5

# Discover getSalesReport /files directory
import getSalesReport as gsr
FILES_DIR = Path(gsr.__file__).resolve().parent / "files"


###############################################################################
# Column candidates
###############################################################################

COLUMN_CANDIDATES = {
    "date": ["Order Time", "Transaction Date", "Transaction Date (Local)", "Date", "Sold At", "Created At", "Order Date"],
    "transaction_id": ["Order ID", "Transaction ID", "Order Number", "Receipt ID", "Ticket", "Ticket Number", "Sale ID", "Cart ID"],
    "employee": ["Budtender Name", "Budtender", "Employee", "Employee Name", "Cashier"],
    "customer_type": ["Customer Type"],
    "product": ["Product Name", "Product", "Item Name", "Item"],
    "category": ["Major Category", "Category", "Product Category", "Product Category Name"],  # prefer Major Category first
    "quantity": ["Total Inventory Sold", "Quantity", "Qty", "Items", "Item Quantity"],
    "gross_sales": ["Gross Sales", "Gross Revenue", "Subtotal", "Total", "Gross"],
    "net_sales": ["Net Sales", "Net Revenue", "Net Total", "Net", "Net Amount", "Total (Net)"],
    "discount_main": ["Discounted Amount", "Discount Amount", "Discount", "Total Discount"],
    "discount_loyalty": ["Loyalty as Discount"],
    "cogs": ["Inventory Cost", "COGS", "Cost of Goods Sold", "Cost"],
    "profit": ["Order Profit", "Profit", "Gross Profit", "Net Profit"],
    "return_date": ["Return Date"],
    "total_weight_sold": ["Total Weight Sold", "Total Weight", "Weight Sold"],
}


###############################################################################
# Brand Theme (your palette)
###############################################################################
THEME = {
    "yellow": colors.HexColor("#FFF200"),
    "green": colors.HexColor("#00AE6F"),
    "black": colors.HexColor("#000000"),
    "muted": colors.HexColor("#374151"),
    "light_bg": colors.HexColor("#F7F7F7"),
    "border": colors.HexColor("#E5E7EB"),
    "row_alt": colors.HexColor("#FAFAFA"),
    "soft_gray": colors.HexColor("#F3F4F6"),
}

# Compact layout
PAGE_MARGINS = {
    "left": 0.45 * inch,
    "right": 0.45 * inch,
    "top": 0.42 * inch,
    "bottom": 0.42 * inch,
}
SPACER = {"xs": 0.04 * inch, "sm": 0.07 * inch, "md": 0.10 * inch}

# Chart color hex
HEX_GREEN = "#00AE6F"
HEX_YELLOW = "#FFF200"
HEX_BLACK = "#000000"
HEX_GRAY_SHADOW = "#9CA3AF"


###############################################################################
# Font setup (nicer font if available)
###############################################################################

BASE_FONT = "Helvetica"
BASE_FONT_BOLD = "Helvetica-Bold"
USE_UNICODE_ARROWS = False

def _try_register_font(name: str, path: str) -> bool:
    try:
        p = Path(path)
        if p.exists():
            pdfmetrics.registerFont(TTFont(name, str(p)))
            return True
    except Exception:
        return False
    return False

def setup_fonts() -> None:
    """Try to use DejaVuSans (nice, readable, supports more glyphs)."""
    global BASE_FONT, BASE_FONT_BOLD, USE_UNICODE_ARROWS

    regular_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
    ]
    bold_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf",
    ]

    reg_ok = False
    bold_ok = False

    for p in regular_candidates:
        if _try_register_font("BuzzSans", p):
            reg_ok = True
            break
    for p in bold_candidates:
        if _try_register_font("BuzzSans-Bold", p):
            bold_ok = True
            break

    if reg_ok and bold_ok:
        BASE_FONT = "BuzzSans"
        BASE_FONT_BOLD = "BuzzSans-Bold"
        USE_UNICODE_ARROWS = True
    else:
        BASE_FONT = "Helvetica"
        BASE_FONT_BOLD = "Helvetica-Bold"
        USE_UNICODE_ARROWS = False


###############################################################################
# Formatting helpers
###############################################################################

def money(x: float) -> str:
    try:
        return f"${x:,.0f}"
    except Exception:
        return "$0"

def money2(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def pct1(x: float) -> str:
    try:
        return f"{x*100:,.1f}%"
    except Exception:
        return "0.0%"

def pp1(x: float) -> str:
    try:
        return f"{x*100:,.1f}pp"
    except Exception:
        return "0.0pp"

def fmt_signed_money(x: float) -> str:
    sign = "+" if x >= 0 else "-"
    return f"{sign}${abs(x):,.0f}"

def fmt_signed_int(x: float) -> str:
    sign = "+" if x >= 0 else "-"
    return f"{sign}{int(abs(x)):,}"

def safe_filename(s: str) -> str:
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-zA-Z0-9 _\-\(\)\.]", "_", s)
    return s

def store_label(store_name: str) -> str:
    label = store_name.replace("Buzz Cannabis", "").strip()
    label = label.replace("(", "").replace(")", "")
    label = re.sub(r"^\-+", "", label).strip()
    return (label or store_name).upper()

def to_number(series: pd.Series) -> pd.Series:
    if series is None:
        return series
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)

    s = series.astype(str)
    s = s.str.replace("$", "", regex=False).str.replace(",", "", regex=False)
    s = s.str.replace("(", "-", regex=False).str.replace(")", "", regex=False)
    s = s.replace({"nan": None, "None": None, "": None})
    return pd.to_numeric(s, errors="coerce")

def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols:
            return cols[key]
    return None

def dow_short(d: date) -> str:
    return d.strftime("%a")  # Sun, Mon...

def fmt_hour_ampm(h: int) -> str:
    """0..23 -> '12a', '1a', ..., '12p', '1p'"""
    h = int(h)
    if h == 0:
        return "12a"
    if 1 <= h <= 11:
        return f"{h}a"
    if h == 12:
        return "12p"
    return f"{h-12}p"

def parse_brand_from_product(product_name: Any) -> str:
    """
    Brand is the part before the first '|'
    Example: "Cold Fire | Cart 1g Pineapple" -> "Cold Fire"
             "Dab Daddy | Flower 14g | | LA Pop Rocks" -> "Dab Daddy"
    """
    s = str(product_name or "").strip()
    if not s:
        return "Unknown"
    parts = [p.strip() for p in s.split("|")]
    for p in parts:
        if p:
            return p
    return "Unknown"


###############################################################################
# Reading exports robustly (Row 5 header fix)
###############################################################################

def guess_header_row(path: Path, tokens: List[str], scan_rows: int = 60) -> int:
    preview = pd.read_excel(path, header=None, nrows=scan_rows, engine="openpyxl")
    token_lc = [t.lower() for t in tokens]
    for i in range(len(preview)):
        row_vals = [
            str(x).strip().lower()
            for x in preview.iloc[i].tolist()
            if str(x).strip() != "nan"
        ]
        joined = " | ".join(row_vals)
        if any(t in joined for t in token_lc):
            return i
    return 0

def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", case=False, regex=True)]
    df.columns = [str(c).strip() for c in df.columns]
    return df

def read_export(path: Path) -> pd.DataFrame:
    if FORCE_HEADER_ROW:
        try:
            df_try = pd.read_excel(path, header=EXPORT_HEADER_ROW_INDEX, engine="openpyxl")
            df_try = _clean_df(df_try)
            if any(c in df_try.columns for c in ["Order ID", "Order Time", "Net Sales", "Gross Sales"]):
                return df_try
        except Exception:
            pass

    header_row = guess_header_row(
        path,
        tokens=["Order ID", "Order Time", "Net Sales", "Gross Sales", "Category", "Budtender Name"],
        scan_rows=80,
    )
    df = pd.read_excel(path, header=header_row, engine="openpyxl")
    return _clean_df(df)


###############################################################################
# Date helpers
###############################################################################

def compute_date_window(backfill_days: int, tz_name: str) -> Tuple[date, date]:
    tz = ZoneInfo(tz_name)
    today = datetime.now(tz).date()
    end_d = today - timedelta(days=1)
    start_d = end_d - timedelta(days=backfill_days - 1)
    return start_d, end_d

def month_start(d: date) -> date:
    return date(d.year, d.month, 1)

def prev_month_same_window(end_d: date) -> Tuple[date, date]:
    """Previous month: day 1 -> same day-of-month (clamped)."""
    if end_d.month == 1:
        py, pm = end_d.year - 1, 12
    else:
        py, pm = end_d.year, end_d.month - 1

    start = date(py, pm, 1)
    first_this_month = date(end_d.year, end_d.month, 1)
    last_prev = first_this_month - timedelta(days=1)
    end_day = min(end_d.day, last_prev.day)
    end = date(py, pm, end_day)
    return start, end

def parse_range_from_folder_name(folder: Path) -> Optional[Tuple[date, date]]:
    """
    Expects folder name like: 2025-12-10_to_2026-02-08
    """
    m = re.match(r"^(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})$", folder.name.strip())
    if not m:
        return None
    try:
        a = datetime.strptime(m.group(1), "%Y-%m-%d").date()
        b = datetime.strptime(m.group(2), "%Y-%m-%d").date()
        return a, b
    except Exception:
        return None


###############################################################################
# Export -> archive
###############################################################################

def cleanup_files_dir(files_dir: Path) -> None:
    files_dir.mkdir(parents=True, exist_ok=True)
    for p in files_dir.iterdir():
        try:
            if p.is_file():
                p.unlink()
        except Exception as e:
            print(f"[WARN] Could not delete {p}: {e}")

def run_export_for_range(start_day: date, end_day: date) -> None:
    print(f"[EXPORT] Running run_sales_report({start_day} -> {end_day})")
    FILES_DIR.mkdir(parents=True, exist_ok=True)

    if CLEANUP_FILES_BEFORE_EXPORT:
        cleanup_files_dir(FILES_DIR)
    else:
        print("[EXPORT] Skipping /files cleanup (CLEANUP_FILES_BEFORE_EXPORT=False)")

    start_dt = datetime(start_day.year, start_day.month, start_day.day)
    end_dt = datetime(end_day.year, end_day.month, end_day.day)

    run_sales_report(start_dt, end_dt)
    print("[EXPORT] Done.")

def archive_exports(start_day: date, end_day: date) -> Tuple[Path, Dict[str, Path]]:
    range_dir = RAW_ROOT / f"{start_day.isoformat()}_to_{end_day.isoformat()}"
    range_dir.mkdir(parents=True, exist_ok=True)

    abbr_to_path: Dict[str, Path] = {}

    for store_name, abbr in store_abbr_map.items():
        src = FILES_DIR / f"sales{abbr}.xlsx"
        if not src.exists():
            print(f"[WARN] Missing export for {store_name} ({abbr}): {src}")
            continue

        nice = store_label(store_name)
        dst_name = f"{abbr} - Sales Export - {nice} - {start_day.isoformat()}_to_{end_day.isoformat()}.xlsx"
        dst = range_dir / safe_filename(dst_name)

        if ARCHIVE_ACTION.lower() == "copy":
            shutil.copy2(str(src), str(dst))
        else:
            shutil.move(str(src), str(dst))

        abbr_to_path[abbr] = dst
        print(f"[ARCHIVE] {abbr}: {dst}")

    return range_dir, abbr_to_path

def find_latest_raw_folder() -> Optional[Path]:
    if not RAW_ROOT.exists():
        return None
    folders = [p for p in RAW_ROOT.iterdir() if p.is_dir()]
    if not folders:
        return None
    return sorted(folders, key=lambda p: p.stat().st_mtime, reverse=True)[0]


###############################################################################
# Metrics
###############################################################################

METRIC_KEYS = [
    "net_revenue",
    "gross_sales",
    "tickets",
    "basket",
    "items",
    "items_per_ticket",
    "net_price_per_item",
    "discount",
    "discount_main",
    "loyalty_discount",
    "discount_rate",
    "profit",
    "margin",
    "cogs",
    "returns_net",
    "returns_tickets",
    "weight_sold",
]

def normalize(df: pd.DataFrame) -> pd.DataFrame:
    date_col = find_col(df, COLUMN_CANDIDATES["date"])
    if not date_col:
        raise RuntimeError(f"Could not find date column. Columns: {list(df.columns)}")

    df = df.copy()
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df[df[date_col].notna()]
    return df

def compute_daily_metrics(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize(df)

    date_col = find_col(df, COLUMN_CANDIDATES["date"])
    tx_col = find_col(df, COLUMN_CANDIDATES["transaction_id"])

    qty_col = find_col(df, COLUMN_CANDIDATES["quantity"])
    net_col = find_col(df, COLUMN_CANDIDATES["net_sales"])
    gross_col = find_col(df, COLUMN_CANDIDATES["gross_sales"])

    disc_main_col = find_col(df, COLUMN_CANDIDATES["discount_main"])
    disc_loyal_col = find_col(df, COLUMN_CANDIDATES["discount_loyalty"])

    cogs_col = find_col(df, COLUMN_CANDIDATES["cogs"])
    profit_col = find_col(df, COLUMN_CANDIDATES["profit"])

    return_col = find_col(df, COLUMN_CANDIDATES["return_date"])
    weight_col = find_col(df, COLUMN_CANDIDATES["total_weight_sold"])

    if not net_col:
        raise RuntimeError(f"Could not find Net Sales column. Columns: {list(df.columns)}")

    df["_date"] = df[date_col].dt.date

    df["_net"] = to_number(df[net_col]).fillna(0)
    df["_gross"] = to_number(df[gross_col]).fillna(0) if gross_col else 0.0
    df["_qty"] = to_number(df[qty_col]).fillna(0) if qty_col else 1.0

    df["_disc_main"] = to_number(df[disc_main_col]).fillna(0) if disc_main_col else 0.0
    df["_disc_loyal"] = to_number(df[disc_loyal_col]).fillna(0) if disc_loyal_col else 0.0
    df["_disc_total"] = df["_disc_main"] + df["_disc_loyal"]

    df["_cogs"] = to_number(df[cogs_col]).fillna(0) if cogs_col else 0.0
    if profit_col:
        df["_profit"] = to_number(df[profit_col]).fillna(0)
    else:
        # fallback profit estimate if missing
        df["_profit"] = df["_net"] - df["_cogs"]

    df["_weight"] = to_number(df[weight_col]).fillna(0) if weight_col else 0.0

    if tx_col:
        tickets = df.groupby("_date")[tx_col].nunique().rename("tickets")
    else:
        tickets = df.groupby("_date").size().rename("tickets")
        print("[WARN] No Order ID column found; ticket count may be inaccurate.")

    daily = df.groupby("_date").agg(
        net_revenue=("_net", "sum"),
        gross_sales=("_gross", "sum"),
        items=("_qty", "sum"),
        discount=("_disc_total", "sum"),
        discount_main=("_disc_main", "sum"),
        loyalty_discount=("_disc_loyal", "sum"),
        cogs=("_cogs", "sum"),
        profit=("_profit", "sum"),
        weight_sold=("_weight", "sum"),
    ).join(tickets)

    daily = daily.reset_index().rename(columns={"_date": "date"})

    # Returns
    if return_col:
        ret = df[df[return_col].notna()].copy()
        if not ret.empty:
            if tx_col:
                returns = ret.groupby("_date").agg(
                    returns_net=("_net", "sum"),
                    returns_tickets=(tx_col, "nunique"),
                ).reset_index().rename(columns={"_date": "date"})
            else:
                returns = ret.groupby("_date").agg(
                    returns_net=("_net", "sum"),
                    returns_tickets=("_net", "size"),
                ).reset_index().rename(columns={"_date": "date"})
            daily = daily.merge(returns, on="date", how="left")

    daily["returns_net"] = daily.get("returns_net", 0.0)
    daily["returns_tickets"] = daily.get("returns_tickets", 0.0)
    daily["returns_net"] = daily["returns_net"].fillna(0.0)
    daily["returns_tickets"] = daily["returns_tickets"].fillna(0.0)

    # Derived
    daily["basket"] = daily.apply(lambda r: r["net_revenue"] / r["tickets"] if r["tickets"] else 0.0, axis=1)
    daily["items_per_ticket"] = daily.apply(lambda r: r["items"] / r["tickets"] if r["tickets"] else 0.0, axis=1)
    daily["net_price_per_item"] = daily.apply(lambda r: r["net_revenue"] / r["items"] if r["items"] else 0.0, axis=1)
    daily["margin"] = daily.apply(lambda r: r["profit"] / r["net_revenue"] if r["net_revenue"] else 0.0, axis=1)

    # discount_rate: prefer gross if available, else approximate gross = net + discount
    def _disc_rate(row):
        g = row["gross_sales"]
        if g:
            return row["discount"] / g
        approx_g = row["net_revenue"] + row["discount"]
        return row["discount"] / approx_g if approx_g else 0.0

    daily["discount_rate"] = daily.apply(_disc_rate, axis=1)

    for k in METRIC_KEYS:
        if k not in daily.columns:
            daily[k] = 0.0

    return daily.sort_values("date")

def slice_range(daily: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    return daily[(daily["date"] >= start) & (daily["date"] <= end)].copy()

def metrics_for_day(daily: pd.DataFrame, day: date) -> Dict[str, float]:
    row = daily[daily["date"] == day]
    if row.empty:
        return {k: 0.0 for k in METRIC_KEYS}
    r = row.iloc[0]
    return {k: float(r.get(k)) if pd.notna(r.get(k)) else 0.0 for k in METRIC_KEYS}

def metrics_for_range(daily: pd.DataFrame, start: date, end: date) -> Dict[str, float]:
    sub = slice_range(daily, start, end)
    if sub.empty:
        return {k: 0.0 for k in METRIC_KEYS}

    sum_fields = [
        "net_revenue", "gross_sales", "tickets", "items", "discount",
        "discount_main", "loyalty_discount",
        "cogs", "profit",
        "returns_net", "returns_tickets",
        "weight_sold",
    ]
    out = {k: float(sub[k].sum()) if k in sub.columns else 0.0 for k in sum_fields}

    net = out["net_revenue"]
    gross = out["gross_sales"]
    tickets = out["tickets"]
    items = out["items"]
    profit = out["profit"]
    disc = out["discount"]

    out["basket"] = net / tickets if tickets else 0.0
    out["items_per_ticket"] = items / tickets if tickets else 0.0
    out["net_price_per_item"] = net / items if items else 0.0
    out["margin"] = profit / net if net else 0.0
    if gross:
        out["discount_rate"] = disc / gross
    else:
        approx_g = net + disc
        out["discount_rate"] = disc / approx_g if approx_g else 0.0

    for k in METRIC_KEYS:
        out.setdefault(k, 0.0)

    return out


###############################################################################
# Breakdowns & summaries
###############################################################################

def _filter_df_date_range(df: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    date_col = find_col(df, COLUMN_CANDIDATES["date"])
    if not date_col:
        return df.iloc[0:0].copy()
    tmp = df.copy()
    tmp[date_col] = pd.to_datetime(tmp[date_col], errors="coerce")
    tmp = tmp[tmp[date_col].notna()]
    tmp["_date"] = tmp[date_col].dt.date
    return tmp[(tmp["_date"] >= start) & (tmp["_date"] <= end)].copy()

def compute_breakdown_net(
    df: pd.DataFrame,
    group_candidates: List[str],
    start: date,
    end: date,
    top_n: Optional[int] = 10,
) -> Optional[pd.DataFrame]:
    group_col = find_col(df, group_candidates)
    net_col = find_col(df, COLUMN_CANDIDATES["net_sales"])
    if not group_col or not net_col:
        return None

    tmp = _filter_df_date_range(df, start, end)
    if tmp.empty:
        return pd.DataFrame(columns=[group_col, "net_revenue"])

    tmp["_net"] = to_number(tmp[net_col]).fillna(0)
    tmp[group_col] = tmp[group_col].fillna("Unknown").astype(str)

    out = tmp.groupby(group_col, as_index=False)["_net"].sum().rename(columns={"_net": "net_revenue"})
    out = out.sort_values("net_revenue", ascending=False)
    if top_n is not None:
        out = out.head(top_n)
    return out

def compute_brand_summary(
    df: pd.DataFrame,
    start: date,
    end: date,
    top_n: int = 10,
) -> Optional[pd.DataFrame]:
    """
    Brand is parsed from product name (before first '|').
    Returns weighted margin = profit_sum / net_sum
    """
    prod_col = find_col(df, COLUMN_CANDIDATES["product"])
    net_col = find_col(df, COLUMN_CANDIDATES["net_sales"])
    profit_col = find_col(df, COLUMN_CANDIDATES["profit"])
    cogs_col = find_col(df, COLUMN_CANDIDATES["cogs"])

    if not prod_col or not net_col:
        return None

    tmp = _filter_df_date_range(df, start, end)
    if tmp.empty:
        return None

    tmp["_net"] = to_number(tmp[net_col]).fillna(0)

    if profit_col:
        tmp["_profit"] = to_number(tmp[profit_col]).fillna(0)
    else:
        tmp["_cogs"] = to_number(tmp[cogs_col]).fillna(0) if cogs_col else 0.0
        tmp["_profit"] = tmp["_net"] - tmp["_cogs"]

    tmp["_brand"] = tmp[prod_col].apply(parse_brand_from_product)

    out = tmp.groupby("_brand", as_index=False).agg(
        net_revenue=("_net", "sum"),
        profit=("_profit", "sum"),
    )
    out["margin"] = out["profit"] / out["net_revenue"].replace({0: None})
    out["margin"] = out["margin"].fillna(0.0)

    out = out.sort_values("net_revenue", ascending=False).head(top_n)
    out = out.rename(columns={"_brand": "brand"})
    return out

def compute_customer_type_summary(df: pd.DataFrame, start: date, end: date) -> Optional[pd.DataFrame]:
    type_col = find_col(df, COLUMN_CANDIDATES["customer_type"])
    net_col = find_col(df, COLUMN_CANDIDATES["net_sales"])
    tx_col = find_col(df, COLUMN_CANDIDATES["transaction_id"])
    if not type_col or not net_col or not tx_col:
        return None

    tmp = _filter_df_date_range(df, start, end)
    if tmp.empty:
        return None

    tmp["_net"] = to_number(tmp[net_col]).fillna(0)
    tmp[type_col] = tmp[type_col].fillna("Unknown").astype(str)

    out = tmp.groupby(type_col, as_index=False).agg(
        net_revenue=("_net", "sum"),
        tickets=(tx_col, "nunique"),
    )
    out["basket"] = out["net_revenue"] / out["tickets"].replace({0: None})
    out["basket"] = out["basket"].fillna(0.0)
    return out.sort_values("net_revenue", ascending=False)

def compute_budtender_summary(df: pd.DataFrame, start: date, end: date) -> Optional[pd.DataFrame]:
    emp_col = find_col(df, COLUMN_CANDIDATES["employee"])
    net_col = find_col(df, COLUMN_CANDIDATES["net_sales"])
    tx_col = find_col(df, COLUMN_CANDIDATES["transaction_id"])
    gross_col = find_col(df, COLUMN_CANDIDATES["gross_sales"])
    disc_main_col = find_col(df, COLUMN_CANDIDATES["discount_main"])
    disc_loyal_col = find_col(df, COLUMN_CANDIDATES["discount_loyalty"])

    if not emp_col or not net_col or not tx_col:
        return None

    tmp = _filter_df_date_range(df, start, end)
    if tmp.empty:
        return None

    tmp["_net"] = to_number(tmp[net_col]).fillna(0)
    tmp["_gross"] = to_number(tmp[gross_col]).fillna(0) if gross_col else 0.0
    tmp["_disc_main"] = to_number(tmp[disc_main_col]).fillna(0) if disc_main_col else 0.0
    tmp["_disc_loyal"] = to_number(tmp[disc_loyal_col]).fillna(0) if disc_loyal_col else 0.0
    tmp["_disc_total"] = tmp["_disc_main"] + tmp["_disc_loyal"]

    tmp[emp_col] = tmp[emp_col].fillna("Unknown").astype(str)

    out = tmp.groupby(emp_col, as_index=False).agg(
        net_revenue=("_net", "sum"),
        gross_sales=("_gross", "sum"),
        discount=("_disc_total", "sum"),
        tickets=(tx_col, "nunique"),
    )
    out["basket"] = out["net_revenue"] / out["tickets"].replace({0: None})
    out["basket"] = out["basket"].fillna(0.0)

    # discount rate: prefer gross; else approximate gross = net+discount
    out["discount_rate"] = out.apply(
        lambda r: (r["discount"] / r["gross_sales"]) if r["gross_sales"] else (r["discount"] / (r["net_revenue"] + r["discount"]) if (r["net_revenue"] + r["discount"]) else 0.0),
        axis=1
    )

    out = out.sort_values("net_revenue", ascending=False).rename(columns={emp_col: "budtender"})
    return out

def compute_category_summary(df: pd.DataFrame, start: date, end: date) -> Optional[pd.DataFrame]:
    """
    Returns category-level metrics (for the 'Major Category Summary' tables).
    """
    cat_col = find_col(df, COLUMN_CANDIDATES["category"])
    date_col = find_col(df, COLUMN_CANDIDATES["date"])
    net_col = find_col(df, COLUMN_CANDIDATES["net_sales"])
    gross_col = find_col(df, COLUMN_CANDIDATES["gross_sales"])
    qty_col = find_col(df, COLUMN_CANDIDATES["quantity"])
    disc_main_col = find_col(df, COLUMN_CANDIDATES["discount_main"])
    disc_loyal_col = find_col(df, COLUMN_CANDIDATES["discount_loyalty"])
    cogs_col = find_col(df, COLUMN_CANDIDATES["cogs"])
    profit_col = find_col(df, COLUMN_CANDIDATES["profit"])

    if not cat_col or not date_col or not net_col:
        return None

    tmp = _filter_df_date_range(df, start, end)
    if tmp.empty:
        return None

    tmp["_net"] = to_number(tmp[net_col]).fillna(0)
    tmp["_gross"] = to_number(tmp[gross_col]).fillna(0) if gross_col else 0.0
    tmp["_qty"] = to_number(tmp[qty_col]).fillna(0) if qty_col else 1.0
    tmp["_disc_main"] = to_number(tmp[disc_main_col]).fillna(0) if disc_main_col else 0.0
    tmp["_disc_loyal"] = to_number(tmp[disc_loyal_col]).fillna(0) if disc_loyal_col else 0.0
    tmp["_disc"] = tmp["_disc_main"] + tmp["_disc_loyal"]

    tmp["_cogs"] = to_number(tmp[cogs_col]).fillna(0) if cogs_col else 0.0
    if profit_col:
        tmp["_profit"] = to_number(tmp[profit_col]).fillna(0)
    else:
        tmp["_profit"] = tmp["_net"] - tmp["_cogs"]

    tmp[cat_col] = tmp[cat_col].fillna("Unknown").astype(str)

    out = tmp.groupby(cat_col, as_index=False).agg(
        net_revenue=("_net", "sum"),
        gross_sales=("_gross", "sum"),
        profit=("_profit", "sum"),
        discount=("_disc", "sum"),
        cogs=("_cogs", "sum"),
        items=("_qty", "sum"),
    ).rename(columns={cat_col: "category"})

    total_net = float(out["net_revenue"].sum()) if not out.empty else 0.0
    total_profit = float(out["profit"].sum()) if not out.empty else 0.0

    out["pct_revenue"] = out["net_revenue"] / (total_net if total_net else 1.0)
    out["pct_profit"] = out["profit"] / (total_profit if total_profit else 1.0) if total_profit else 0.0

    # Derived
    def _disc_rate_row(r):
        if r["gross_sales"]:
            return r["discount"] / r["gross_sales"]
        approx_g = r["net_revenue"] + r["discount"]
        return r["discount"] / approx_g if approx_g else 0.0

    out["discount_rate"] = out.apply(_disc_rate_row, axis=1)
    out["margin"] = out["profit"] / out["net_revenue"].replace({0: None})
    out["margin"] = out["margin"].fillna(0.0)

    out["price_per_item"] = out["net_revenue"] / out["items"].replace({0: None})
    out["price_per_item"] = out["price_per_item"].fillna(0.0)

    out["profit_per_item"] = out["profit"] / out["items"].replace({0: None})
    out["profit_per_item"] = out["profit_per_item"].fillna(0.0)

    out["cogs_pct"] = out["cogs"] / out["net_revenue"].replace({0: None})
    out["cogs_pct"] = out["cogs_pct"].fillna(0.0)

    out = out.sort_values("net_revenue", ascending=False)
    return out

def compute_hourly_metrics(df: pd.DataFrame, day: date) -> Optional[pd.DataFrame]:
    """
    Hourly metrics for one day:
      - net_revenue, profit, tickets, basket, margin
    """
    date_col = find_col(df, COLUMN_CANDIDATES["date"])
    tx_col = find_col(df, COLUMN_CANDIDATES["transaction_id"])
    net_col = find_col(df, COLUMN_CANDIDATES["net_sales"])
    profit_col = find_col(df, COLUMN_CANDIDATES["profit"])
    cogs_col = find_col(df, COLUMN_CANDIDATES["cogs"])

    if not date_col or not net_col:
        return None

    tmp = df.copy()
    tmp[date_col] = pd.to_datetime(tmp[date_col], errors="coerce")
    tmp = tmp[tmp[date_col].notna()]
    tmp["_date"] = tmp[date_col].dt.date
    tmp = tmp[tmp["_date"] == day]
    if tmp.empty:
        return pd.DataFrame(columns=["hour", "net_revenue", "profit", "tickets", "basket", "margin"])

    tmp["_hour"] = tmp[date_col].dt.hour
    tmp["_net"] = to_number(tmp[net_col]).fillna(0)

    if profit_col:
        tmp["_profit"] = to_number(tmp[profit_col]).fillna(0)
    else:
        tmp["_cogs"] = to_number(tmp[cogs_col]).fillna(0) if cogs_col else 0.0
        tmp["_profit"] = tmp["_net"] - tmp["_cogs"]

    if tx_col:
        agg = tmp.groupby("_hour").agg(
            net_revenue=("_net", "sum"),
            profit=("_profit", "sum"),
            tickets=(tx_col, "nunique"),
        ).reset_index().rename(columns={"_hour": "hour"})
    else:
        agg = tmp.groupby("_hour").agg(
            net_revenue=("_net", "sum"),
            profit=("_profit", "sum"),
            tickets=("_net", "size"),
        ).reset_index().rename(columns={"_hour": "hour"})

    agg["basket"] = agg["net_revenue"] / agg["tickets"].replace({0: None})
    agg["basket"] = agg["basket"].fillna(0.0)

    agg["margin"] = agg["profit"] / agg["net_revenue"].replace({0: None})
    agg["margin"] = agg["margin"].fillna(0.0)

    return agg.sort_values("hour")


###############################################################################
# Charts (compact + visual)
###############################################################################

def _mpl_setup():
    plt.rcParams.update({
        "font.size": 8.3,
        "axes.titlesize": 10.2,
        "axes.labelsize": 8.0,
        "axes.edgecolor": "#D1D5DB",
        "axes.linewidth": 0.8,
        "grid.color": "#E5E7EB",
        "grid.linewidth": 0.8,
    })

def chart_trend_bar_with_labels(
    daily: pd.DataFrame,
    value_col: str,
    title: str,
    days: int = 14,
    kind: str = "money",
) -> BytesIO:
    """
    Full-width bar trend with:
      - x labels: YYYY-MM-DD DOW
      - value labels above bars
    """
    _mpl_setup()
    buf = BytesIO()
    if daily is None or daily.empty or value_col not in daily.columns:
        return buf

    tail = daily.tail(days).copy()
    labels = [f"{d.isoformat()} {dow_short(d)}" for d in tail["date"].tolist()]
    values = tail[value_col].fillna(0).astype(float).tolist()

    plt.figure(figsize=(7.25, 2.25))
    bars = plt.bar(range(len(labels)), values, color=HEX_GREEN)

    plt.title(title)
    plt.xticks(range(len(labels)), labels, rotation=35, ha="right", fontsize=7.2)
    plt.grid(True, axis="y", alpha=1.0)
    plt.tight_layout()

    # annotate values
    if values:
        vmax = max(values)
        pad = (vmax * 0.02) if vmax else 1.0
        for i, v in enumerate(values):
            if kind == "money":
                txt = money(v)
            elif kind == "int":
                txt = f"{int(v):,}"
            else:
                txt = pct1(v)
            plt.text(i, v + pad, txt, ha="center", va="bottom", fontsize=7.2)

    plt.savefig(buf, format="png", dpi=190)
    plt.close()
    buf.seek(0)
    return buf

def chart_rank_barh(
    df: pd.DataFrame,
    label_col: str,
    value_col: str,
    title: str,
    top_n: int = 10,
    figsize: Tuple[float, float] = (7.25, 2.7),
) -> BytesIO:
    _mpl_setup()
    buf = BytesIO()
    if df is None or df.empty:
        return buf

    d = df.head(top_n).copy()
    labels = d[label_col].astype(str).tolist()[::-1]
    values = d[value_col].astype(float).tolist()[::-1]

    plt.figure(figsize=figsize)
    plt.barh(range(len(labels)), values, color=HEX_GREEN)
    plt.title(title)
    plt.yticks(range(len(labels)), labels, fontsize=7.6)
    plt.grid(True, axis="x", alpha=1.0)
    plt.tight_layout()

    plt.savefig(buf, format="png", dpi=190)
    plt.close()
    buf.seek(0)
    return buf

def chart_hourly_shadow_compare(
    this_day: pd.DataFrame,
    last_week: pd.DataFrame,
    metric_col: str,
    title: str,
    kind: str,
    figsize: Tuple[float, float] = (3.55, 2.15),
) -> BytesIO:
    """
    Shadow comparison:
      - Last week: wide, semi-transparent YELLOW
      - Report day: narrower, solid GREEN on top
    """
    _mpl_setup()
    buf = BytesIO()

    if this_day is None or last_week is None:
        return buf
    if this_day.empty and last_week.empty:
        return buf

    # Union of hours, but keep sorted
    hours = sorted(set(this_day["hour"].tolist()) | set(last_week["hour"].tolist()))
    if not hours:
        return buf

    this_map = {int(h): float(v) for h, v in zip(this_day["hour"], this_day[metric_col])} if (metric_col in this_day.columns) else {}
    last_map = {int(h): float(v) for h, v in zip(last_week["hour"], last_week[metric_col])} if (metric_col in last_week.columns) else {}

    this_vals = [this_map.get(h, 0.0) for h in hours]
    last_vals = [last_map.get(h, 0.0) for h in hours]

    # For percent charts, convert to percent points on axis
    if kind == "pct":
        this_vals_plot = [v * 100.0 for v in this_vals]
        last_vals_plot = [v * 100.0 for v in last_vals]
        yfmt = lambda v: f"{v:,.0f}%"
    elif kind == "int":
        this_vals_plot = this_vals
        last_vals_plot = last_vals
        yfmt = lambda v: f"{int(v):,}"
    else:
        this_vals_plot = this_vals
        last_vals_plot = last_vals
        yfmt = lambda v: money(v)

    x = list(range(len(hours)))

    plt.figure(figsize=figsize)
    # shadow layer
    plt.bar(
        x,
        last_vals_plot,
        width=0.82,
        color=HEX_YELLOW,
        alpha=0.35,
        edgecolor=HEX_BLACK,
        linewidth=0.4,
        label="Last Week",
        zorder=1,
    )
    # top layer
    plt.bar(
        x,
        this_vals_plot,
        width=0.52,
        color=HEX_GREEN,
        alpha=1.0,
        edgecolor=HEX_BLACK,
        linewidth=0.3,
        label="Report Day",
        zorder=2,
    )

    xt = [fmt_hour_ampm(h) for h in hours]
    plt.xticks(x, xt, fontsize=7.2)
    plt.title(title)
    plt.grid(True, axis="y", alpha=1.0, zorder=0)
    plt.legend(loc="upper right", frameon=False, fontsize=7.4)
    plt.tight_layout()

    plt.savefig(buf, format="png", dpi=190)
    plt.close()
    buf.seek(0)
    return buf


###############################################################################
# PDF visuals: KPI + tables + category "bar cells"
###############################################################################

def build_styles() -> Dict[str, Any]:
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="TitleBig",
        parent=styles["Title"],
        fontName=BASE_FONT_BOLD,
        fontSize=16.5,
        textColor=THEME["black"],
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        name="Muted",
        parent=styles["Normal"],
        fontName=BASE_FONT,
        fontSize=8.6,
        textColor=THEME["muted"],
        leading=10.2,
    ))
    styles.add(ParagraphStyle(
        name="Section",
        parent=styles["Heading2"],
        fontName=BASE_FONT_BOLD,
        fontSize=11.0,
        textColor=THEME["black"],
        spaceBefore=5,
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        name="KpiLabel",
        parent=styles["Normal"],
        fontName=BASE_FONT_BOLD,
        fontSize=8.3,
        textColor=THEME["black"],
        leading=9.6,
    ))
    styles.add(ParagraphStyle(
        name="KpiValue",
        parent=styles["Normal"],
        fontName=BASE_FONT_BOLD,
        fontSize=13.6,
        textColor=THEME["black"],
        leading=14.6,
    ))
    styles.add(ParagraphStyle(
        name="KpiDelta",
        parent=styles["Normal"],
        fontName=BASE_FONT,
        fontSize=8.2,
        textColor=THEME["muted"],
        leading=9.6,
    ))
    styles.add(ParagraphStyle(
        name="Small",
        parent=styles["Normal"],
        fontName=BASE_FONT,
        fontSize=8.5,
        leading=10.2,
        textColor=THEME["black"],
    ))
    styles.add(ParagraphStyle(
        name="Tiny",
        parent=styles["Normal"],
        fontName=BASE_FONT,
        fontSize=7.8,
        leading=9.4,
        textColor=THEME["muted"],
    ))
    return styles

def _arrow(diff: float) -> str:
    if USE_UNICODE_ARROWS:
        return "▲" if diff >= 0 else "▼"
    return "+" if diff >= 0 else "-"

def delta_html_currency(current: float, baseline: float, label: str) -> str:
    if baseline == 0:
        return f"<font color='#374151'>vs {label}: n/a</font>"
    diff = current - baseline
    pct = diff / baseline
    arrow = _arrow(diff)
    color = "#00AE6F" if diff >= 0 else "#111827"
    return f"<font color='{color}'>{arrow} {fmt_signed_money(diff)} ({pct1(pct)})</font> <font color='#374151'>vs {label}</font>"

def delta_html_int(current: float, baseline: float, label: str) -> str:
    if baseline == 0:
        return f"<font color='#374151'>vs {label}: n/a</font>"
    diff = current - baseline
    pct = diff / baseline
    arrow = _arrow(diff)
    color = "#00AE6F" if diff >= 0 else "#111827"
    return f"<font color='{color}'>{arrow} {fmt_signed_int(diff)} ({pct1(pct)})</font> <font color='#374151'>vs {label}</font>"

def delta_html_pp(current: float, baseline: float, label: str) -> str:
    if baseline == 0 and current == 0:
        return f"<font color='#374151'>vs {label}: n/a</font>"
    diff = current - baseline
    arrow = _arrow(diff)
    color = "#00AE6F" if diff >= 0 else "#111827"
    return f"<font color='{color}'>{arrow} {pp1(diff)}</font> <font color='#374151'>vs {label}</font>"

def kpi_cell(styles, label: str, value: str, delta_html: str) -> List[Paragraph]:
    return [
        Paragraph(label, styles["KpiLabel"]),
        Paragraph(value, styles["KpiValue"]),
        Paragraph(delta_html, styles["KpiDelta"]),
    ]

def build_kpi_grid(styles, cells: List[List[Paragraph]], cols: int = 4) -> Table:
    while len(cells) % cols != 0:
        cells.append(kpi_cell(styles, "", "", ""))

    rows = [cells[i:i+cols] for i in range(0, len(cells), cols)]
    t = Table(rows, colWidths=[(7.6 * inch) / cols] * cols)

    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), THEME["light_bg"]),
        ("BOX", (0, 0), (-1, -1), 0.6, THEME["border"]),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, THEME["border"]),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return t

def build_table(headers: List[Any], rows: List[List[Any]], col_widths: Optional[List[float]] = None) -> Table:
    data = [headers] + rows
    t = Table(data, colWidths=col_widths, repeatRows=1)

    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), THEME["black"]),
        ("TEXTCOLOR", (0, 0), (-1, 0), THEME["yellow"]),
        ("FONTNAME", (0, 0), (-1, 0), BASE_FONT_BOLD),
        ("FONTNAME", (0, 1), (-1, -1), BASE_FONT),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.4, THEME["border"]),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, THEME["row_alt"]]),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t

def build_report_day_band(report_day: date, width: float) -> Table:
    p = Paragraph(
        f"<b>REPORT DAY:</b> {report_day.isoformat()} ({report_day.strftime('%A')})",
        ParagraphStyle(
            name="ReportBand",
            fontName=BASE_FONT_BOLD,
            fontSize=10.0,
            textColor=THEME["black"],
            leading=12,
        )
    )
    t = Table([[p]], colWidths=[width])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), THEME["yellow"]),
        ("BOX", (0, 0), (-1, -1), 0.8, THEME["black"]),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t

def make_footer(left_text: str, report_day: date):
    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(BASE_FONT, 8)
        canvas.setFillColor(THEME["muted"])
        canvas.drawString(doc.leftMargin, 0.30 * inch, f"{left_text} • {report_day.isoformat()} ({report_day.strftime('%A')})")
        canvas.drawRightString(letter[0] - doc.rightMargin, 0.30 * inch, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()
    return _footer


###############################################################################
# Category Summary "BarCell" to mimic the screenshot feel
###############################################################################

class BarCell(Flowable):
    """
    Draws a horizontal bar (ratio * cell width) with a value label on top.
    This makes tables feel like a dashboard (similar to your screenshot).
    """
    def __init__(
        self,
        text: str,
        ratio: float,
        bar_hex: str,
        font_name: str,
        font_size: float = 7.9,
        text_color_hex: str = "#111827",
    ):
        super().__init__()
        self.text = str(text)
        self.ratio = max(0.0, min(1.0, float(ratio)))
        self.bar_hex = bar_hex
        self.font_name = font_name
        self.font_size = font_size
        self.text_color_hex = text_color_hex

        self.width = 0
        self.height = 0

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        # compact row height
        self.height = 0.18 * inch
        return self.width, self.height

    def draw(self):
        c = self.canv
        # bar
        bar_w = self.width * self.ratio
        c.saveState()
        c.setFillColor(colors.HexColor(self.bar_hex))
        c.setStrokeColor(colors.HexColor(self.bar_hex))
        c.rect(0, 0, bar_w, self.height, fill=1, stroke=0)

        # text
        c.setFillColor(colors.HexColor(self.text_color_hex))
        c.setFont(self.font_name, self.font_size)
        # right align text with slight padding
        c.drawRightString(self.width - 2, (self.height / 2) - 3, self.text)
        c.restoreState()

def _safe_max(series: pd.Series) -> float:
    try:
        v = float(series.max())
        return v if v > 0 else 0.0
    except Exception:
        return 0.0
    
CATEGORY_MAX_ROWS = 8
def build_category_summary_table(
    styles,
    cat_df: pd.DataFrame,
    title: str,
    top_n: int = CATEGORY_MAX_ROWS,
) -> List[Any]:
    """
    Compact, page-safe Major Category Summary table.
    Guaranteed to fit on one page.
    """
    if cat_df is None or cat_df.empty:
        return []

    # --- limit rows strictly ---
    d_all = cat_df.copy()
    d = d_all.head(top_n).copy()

    # --- totals ---
    totals = {
        "category": "Totals",
        "net_revenue": float(d_all["net_revenue"].sum()),
        "profit": float(d_all["profit"].sum()),
        "discount": float(d_all["discount"].sum()),
        "cogs": float(d_all["cogs"].sum()),
        "items": float(d_all["items"].sum()),
    }

    gross_total = float(d_all["gross_sales"].sum()) if "gross_sales" in d_all.columns else 0.0
    if gross_total:
        totals["discount_rate"] = totals["discount"] / gross_total
    else:
        approx_g = totals["net_revenue"] + totals["discount"]
        totals["discount_rate"] = totals["discount"] / approx_g if approx_g else 0.0

    totals["margin"] = totals["profit"] / totals["net_revenue"] if totals["net_revenue"] else 0.0
    totals["price_per_item"] = totals["net_revenue"] / totals["items"] if totals["items"] else 0.0
    totals["profit_per_item"] = totals["profit"] / totals["items"] if totals["items"] else 0.0
    totals["cogs_pct"] = totals["cogs"] / totals["net_revenue"] if totals["net_revenue"] else 0.0

    total_net = float(d_all["net_revenue"].sum()) or 1.0
    total_profit = float(d_all["profit"].sum()) or 1.0

    d["pct_revenue"] = d["net_revenue"] / total_net
    d["pct_profit"] = d["profit"] / total_profit

    # --- max values for bar scaling ---
    max_rev = _safe_max(d["net_revenue"])
    max_profit = _safe_max(d["profit"].abs())
    max_items = _safe_max(d["items"])
    max_price = _safe_max(d["price_per_item"])
    max_ppi = _safe_max(d["profit_per_item"].abs())
    max_disc = _safe_max(d["discount_rate"])
    max_margin = _safe_max(d["margin"])
    max_cogs = _safe_max(d["cogs_pct"])

    headers = [
        "#",
        "Major Category",
        "Revenue",
        "% Rev",
        "Profit",
        "% Profit",
        "Discount %",
        "Margin %",
        "Items",
        "Price/Item",
        "Profit/Item",
        "% COGS",
    ]

    rows: List[List[Any]] = []

    for idx, r in enumerate(d.itertuples(index=False), start=1):
        rows.append([
            str(idx),
            str(r.category),
            BarCell(money(r.net_revenue), (r.net_revenue / max_rev) if max_rev else 0.0, HEX_GREEN, BASE_FONT, 7.4),
            BarCell(pct1(r.pct_revenue), (r.pct_revenue / d["pct_revenue"].max()) if d["pct_revenue"].max() else 0.0, HEX_GREEN, BASE_FONT, 7.4),
            BarCell(money(r.profit), (abs(r.profit) / max_profit) if max_profit else 0.0, HEX_GREEN, BASE_FONT, 7.4),
            BarCell(pct1(r.pct_profit), (abs(r.pct_profit) / d["pct_profit"].abs().max()) if d["pct_profit"].abs().max() else 0.0, HEX_GREEN, BASE_FONT, 7.4),
            BarCell(pct1(r.discount_rate), (r.discount_rate / max_disc) if max_disc else 0.0, HEX_YELLOW, BASE_FONT, 7.4),
            BarCell(pct1(r.margin), (r.margin / max_margin) if max_margin else 0.0, HEX_GREEN, BASE_FONT, 7.4),
            BarCell(f"{int(r.items):,}", (r.items / max_items) if max_items else 0.0, HEX_GREEN, BASE_FONT, 7.4),
            BarCell(money2(r.price_per_item), (r.price_per_item / max_price) if max_price else 0.0, HEX_YELLOW, BASE_FONT, 7.4),
            BarCell(money2(r.profit_per_item), (abs(r.profit_per_item) / max_ppi) if max_ppi else 0.0, HEX_GREEN, BASE_FONT, 7.4),
            BarCell(pct1(r.cogs_pct), (r.cogs_pct / max_cogs) if max_cogs else 0.0, HEX_YELLOW, BASE_FONT, 7.4),
        ])

    # --- totals row ---
    rows.append([
        "",
        "Totals<",
        money(totals["net_revenue"]),
        "100.0%",
        money(totals["profit"]),
        "100.0%",
        pct1(totals["discount_rate"]),
        pct1(totals["margin"]),
        f"{int(totals['items']):,}",
        money2(totals["price_per_item"]),
        money2(totals["profit_per_item"]),
        pct1(totals["cogs_pct"]),
    ])

    table = Table(
        [headers] + rows,
        repeatRows=1,
        colWidths=[
            0.18*inch,  # #
            1.25*inch,  # Major Category
            0.85*inch,  # Revenue
            0.55*inch,  # % Rev
            0.80*inch,  # Profit
            0.55*inch,  # % Profit
            0.60*inch,  # Discount %
            0.55*inch,  # Margin %
            0.55*inch,  # Items
            0.70*inch,  # Price / Item
            0.70*inch,  # Profit / Item
            0.50*inch,  # % COGS
        ],
    )

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), THEME["black"]),
        ("TEXTCOLOR", (0, 0), (-1, 0), THEME["yellow"]),
        ("FONTNAME", (0, 0), (-1, 0), BASE_FONT_BOLD),
        ("FONTSIZE", (0, 0), (-1, 0), 8.0),
        ("GRID", (0, 0), (-1, -1), 0.3, THEME["border"]),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, THEME["row_alt"]]),
        ("FONTNAME", (0, 1), (-1, -1), BASE_FONT),
        ("FONTSIZE", (0, 1), (-1, -1), 7.6),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, -1), (-1, -1), THEME["soft_gray"]),
        ("FONTNAME", (0, -1), (-1, -1), BASE_FONT_BOLD),
    ]))

    return [
        KeepTogether([
            Paragraph(title, styles["Section"]),
            table,
        ])
    ]

###############################################################################
# PDF: Store report
###############################################################################

def build_store_pdf(
    out_pdf: Path,
    store_name: str,
    abbr: str,
    df_raw: pd.DataFrame,
    daily: pd.DataFrame,
    start_day: date,
    end_day: date,
) -> None:
    styles = build_styles()
    label = store_label(store_name)
    generated_at = datetime.now(ZoneInfo(REPORT_TZ)).strftime("%B %d, %Y at %I:%M %p %Z")

    # Comparison windows
    last_week_day = end_day - timedelta(days=7)
    mtd_start = month_start(end_day)
    last_mtd_start, last_mtd_end = prev_month_same_window(end_day)

    # KPIs
    today = metrics_for_day(daily, end_day)
    last_week = metrics_for_day(daily, last_week_day)
    mtd = metrics_for_range(daily, mtd_start, end_day)
    last_mtd = metrics_for_range(daily, last_mtd_start, last_mtd_end)

    # Month pacing
    days_elapsed = (end_day - mtd_start).days + 1
    avg_per_day = (mtd["net_revenue"] / days_elapsed) if days_elapsed else 0.0

    # Trend chart
    trend = daily[daily["date"] <= end_day].copy().tail(max(TREND_DAYS, 1))
    net_trend = chart_trend_bar_with_labels(
        trend,
        "net_revenue",
        f"Net Sales Trend (Last {len(trend)} Days)",
        days=len(trend),
        kind="money",
    )

    # Hourly metrics charts (shadow compare)
    hourly_today = compute_hourly_metrics(df_raw, end_day)
    hourly_last = compute_hourly_metrics(df_raw, last_week_day)
    if hourly_today is None:
        hourly_today = pd.DataFrame(columns=["hour", "net_revenue", "profit", "tickets", "basket", "margin"])
    if hourly_last is None:
        hourly_last = pd.DataFrame(columns=["hour", "net_revenue", "profit", "tickets", "basket", "margin"])

    ch_rev = chart_hourly_shadow_compare(
        hourly_today, hourly_last, "net_revenue",
        title="Revenue by Hour",
        kind="money",
        figsize=(3.55, 2.15),
    )
    ch_tix = chart_hourly_shadow_compare(
        hourly_today, hourly_last, "tickets",
        title="Tickets by Hour",
        kind="int",
        figsize=(3.55, 2.15),
    )
    ch_profit = chart_hourly_shadow_compare(
        hourly_today, hourly_last, "profit",
        title="Profit by Hour",
        kind="money",
        figsize=(3.55, 2.15),
    )
    ch_basket = chart_hourly_shadow_compare(
        hourly_today, hourly_last, "basket",
        title="Basket by Hour",
        kind="money",
        figsize=(3.55, 2.15),
    )
    ch_margin = chart_hourly_shadow_compare(
        hourly_today, hourly_last, "margin",
        title="Margin by Hour",
        kind="pct",
        figsize=(3.55, 2.15),
    )

    # Drivers (MTD)
    cat_today = compute_category_summary(df_raw, end_day, end_day)
    cat_mtd = compute_category_summary(df_raw, mtd_start, end_day)

    prod_mtd = compute_breakdown_net(df_raw, COLUMN_CANDIDATES["product"], mtd_start, end_day, top_n=TOP_N)
    prod_chart = BytesIO()
    if prod_mtd is not None and not prod_mtd.empty:
        prod_chart = chart_rank_barh(
            prod_mtd.rename(columns={prod_mtd.columns[0]: "product"}),
            "product", "net_revenue",
            "Top Products (MTD)",
            top_n=TOP_N,
            figsize=(7.25, 2.8),
        )

    brand_mtd = compute_brand_summary(df_raw, mtd_start, end_day, top_n=TOP_N)
    brand_chart = BytesIO()
    if brand_mtd is not None and not brand_mtd.empty:
        brand_chart = chart_rank_barh(
            brand_mtd,
            "brand", "net_revenue",
            "Top Brands (MTD)",
            top_n=TOP_N,
            figsize=(7.25, 2.8),
        )

    # Budtenders: report day + MTD (ALL)
    bud_today = compute_budtender_summary(df_raw, end_day, end_day)
    bud_mtd = compute_budtender_summary(df_raw, mtd_start, end_day)

    bud_today_chart = BytesIO()
    if bud_today is not None and not bud_today.empty:
        bud_today_chart = chart_rank_barh(bud_today, "budtender", "net_revenue", "Top Budtenders (Report Day)", top_n=min(TOP_N, len(bud_today)), figsize=(7.25, 2.7))

    bud_mtd_chart = BytesIO()
    if bud_mtd is not None and not bud_mtd.empty:
        bud_mtd_chart = chart_rank_barh(bud_mtd, "budtender", "net_revenue", "Top Budtenders (MTD)", top_n=min(TOP_N, len(bud_mtd)), figsize=(7.25, 2.7))

    # PDF setup
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(out_pdf),
        pagesize=letter,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title=f"{abbr} Owner Snapshot - {label}",
        author="Buzz Automation",
    )
    footer = make_footer(f"{abbr} - {label}", end_day)

    story: List[Any] = []

    # Header
    story.append(Paragraph(f"{abbr} • Owner Snapshot • {label}", styles["TitleBig"]))
    story.append(build_report_day_band(end_day, width=7.6 * inch))
    story.append(Spacer(1, SPACER["xs"]))

    # Windows line (small)
    story.append(Paragraph(
        f"<b>Data Window:</b> {start_day.isoformat()} → {end_day.isoformat()} &nbsp;&nbsp; "
        f"<b>MTD Window:</b> {mtd_start.isoformat()} → {end_day.isoformat()} &nbsp;&nbsp; "
        f"<b>Last MTD Ref:</b> {last_mtd_start.isoformat()} → {last_mtd_end.isoformat()}",
        styles["Tiny"],
    ))
    story.append(Paragraph(f"<b>Generated:</b> {generated_at}", styles["Tiny"]))
    story.append(Spacer(1, SPACER["sm"]))

    # KPI grid
    kpis: List[List[Paragraph]] = []
    kpis.append(kpi_cell(styles, "TODAY • NET SALES", money(today["net_revenue"]), delta_html_currency(today["net_revenue"], last_week["net_revenue"], "last week")))
    kpis.append(kpi_cell(styles, "TODAY • TICKETS", f"{int(today['tickets']):,}", delta_html_int(today["tickets"], last_week["tickets"], "last week")))
    kpis.append(kpi_cell(styles, "TODAY • BASKET", money2(today["basket"]), delta_html_currency(today["basket"], last_week["basket"], "last week")))
    kpis.append(kpi_cell(styles, "TODAY • DISC RATE", pct1(today["discount_rate"]), delta_html_pp(today["discount_rate"], last_week["discount_rate"], "last week")))

    kpis.append(kpi_cell(styles, "MTD • NET SALES", money(mtd["net_revenue"]), delta_html_currency(mtd["net_revenue"], last_mtd["net_revenue"], "last MTD")))
    kpis.append(kpi_cell(styles, "MTD • TICKETS", f"{int(mtd['tickets']):,}", delta_html_int(mtd["tickets"], last_mtd["tickets"], "last MTD")))
    kpis.append(kpi_cell(styles, "MTD • BASKET", money2(mtd["basket"]), delta_html_currency(mtd["basket"], last_mtd["basket"], "last MTD")))
    kpis.append(kpi_cell(styles, "MTD • MARGIN", pct1(mtd["margin"]), delta_html_pp(mtd["margin"], last_mtd["margin"], "last MTD")))

    story.append(build_kpi_grid(styles, kpis, cols=4))
    story.append(Spacer(1, SPACER["sm"]))

    # Pulse line
    story.append(Paragraph(
        f"<b>MTD Avg/Day:</b> {money(avg_per_day)}"
        f"&nbsp;&nbsp; <b>MTD Discount:</b> {money(mtd['discount'])}"
        f"&nbsp;&nbsp; <b>MTD Returns:</b> {money(mtd['returns_net'])}",
        styles["Muted"],
    ))
    story.append(Spacer(1, SPACER["sm"]))

    # Trend (keep together so header doesn't orphan)
    story.append(KeepTogether([
        Paragraph("Trends", styles["Section"]),
        Image(net_trend, width=7.25 * inch, height=2.25 * inch) if net_trend.getbuffer().nbytes > 0 else Spacer(1, 0),
    ]))
    story.append(Spacer(1, SPACER["xs"]))

    # Hourly quick compare (two-up)
    story.append(Paragraph(f"Hourly Snapshot (Report Day vs {last_week_day.isoformat()} {dow_short(last_week_day)})", styles["Section"]))
    hourly_grid = Table(
        [[
            Image(ch_rev, width=3.55*inch, height=2.15*inch) if ch_rev.getbuffer().nbytes > 0 else Spacer(1, 0),
            Image(ch_tix, width=3.55*inch, height=2.15*inch) if ch_tix.getbuffer().nbytes > 0 else Spacer(1, 0),
        ]],
        colWidths=[3.8*inch, 3.8*inch],
    )
    hourly_grid.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(hourly_grid)

    # Page 2: Hourly deep-dive (more metrics)
    story.append(PageBreak())
    story.append(Paragraph("Hourly Performance", styles["TitleBig"]))
    story.append(Paragraph("Shadow compare = Last Week (Yellow) behind Report Day (Green).", styles["Tiny"]))
    story.append(Spacer(1, SPACER["sm"]))

    hourly_grid2 = Table(
        [
            [
                Image(ch_profit, width=3.55*inch, height=2.15*inch) if ch_profit.getbuffer().nbytes > 0 else Spacer(1, 0),
                Image(ch_basket, width=3.55*inch, height=2.15*inch) if ch_basket.getbuffer().nbytes > 0 else Spacer(1, 0),
            ],
            [
                Image(ch_margin, width=3.55*inch, height=2.15*inch) if ch_margin.getbuffer().nbytes > 0 else Spacer(1, 0),
                Spacer(1, 0),
            ],
        ],
        colWidths=[3.8*inch, 3.8*inch],
    )
    hourly_grid2.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(hourly_grid2)

    # Page 3: Category summary + products/brands
    story.append(PageBreak())
    story.append(Paragraph("Drivers", styles["TitleBig"]))
    story.append(Paragraph("Major Categories + Products + Brands (MTD).", styles["Tiny"]))
    story.append(Spacer(1, SPACER["sm"]))

    # Category tables (today + mtd)
    if cat_today is not None and not cat_today.empty:
        story += build_category_summary_table(styles, cat_today, "Major Category Summary — Today", top_n=CATEGORY_TOP_N)
        story.append(Spacer(1, SPACER["sm"]))

    if cat_mtd is not None and not cat_mtd.empty:
        story += build_category_summary_table(styles, cat_mtd, "Major Category Summary — MTD", top_n=CATEGORY_TOP_N)
        story.append(Spacer(1, SPACER["sm"]))

    # Top Products (keep header+chart together to prevent orphaned title page)
    if prod_mtd is not None and not prod_mtd.empty and prod_chart.getbuffer().nbytes > 0:
        prod_rows = [[str(r[0]), money(float(r.net_revenue))] for r in prod_mtd.itertuples(index=False)]
        story.append(KeepTogether([
            Paragraph("Top Products (MTD)", styles["Section"]),
            Image(prod_chart, width=7.25 * inch, height=2.8 * inch),
            build_table(["Product", "MTD Net"], prod_rows, [5.85*inch, 1.4*inch]),
        ]))
        story.append(Spacer(1, SPACER["sm"]))

    # Top Brands + margin
    if brand_mtd is not None and not brand_mtd.empty and brand_chart.getbuffer().nbytes > 0:
        brand_rows = [[str(r.brand), money(float(r.net_revenue)), pct1(float(r.margin))] for r in brand_mtd.itertuples(index=False)]
        story.append(KeepTogether([
            Paragraph("Top Brands (MTD)", styles["Section"]),
            Image(brand_chart, width=7.25 * inch, height=2.8 * inch),
            build_table(["Brand", "MTD Net", "Avg Margin"], brand_rows, [4.95*inch, 1.4*inch, 1.25*inch]),
        ]))

    # Page 4: Budtenders (Daily + MTD)
    story.append(PageBreak())
    story.append(Paragraph("Staff Performance", styles["TitleBig"]))
    story.append(Paragraph("Budtenders — Report Day and MTD (full lists).", styles["Tiny"]))
    story.append(Spacer(1, SPACER["sm"]))

    if bud_today is not None and not bud_today.empty:
        story.append(Paragraph(f"Budtenders — Report Day ({end_day.isoformat()} {dow_short(end_day)})", styles["Section"]))
        if bud_today_chart.getbuffer().nbytes > 0:
            story.append(Image(bud_today_chart, width=7.25*inch, height=2.7*inch))
        bud_today_rows = []
        for r in bud_today.itertuples(index=False):
            bud_today_rows.append([str(r.budtender), money(float(r.net_revenue)), f"{int(r.tickets):,}", money2(float(r.basket)), pct1(float(r.discount_rate))])
        story.append(build_table(
            ["Budtender", "Net", "Tickets", "Basket", "Disc Rate"],
            bud_today_rows,
            [2.65*inch, 1.25*inch, 1.0*inch, 1.25*inch, 1.2*inch],
        ))
        story.append(Spacer(1, SPACER["sm"]))

    if bud_mtd is not None and not bud_mtd.empty:
        story.append(Paragraph("Budtenders — MTD", styles["Section"]))
        if bud_mtd_chart.getbuffer().nbytes > 0:
            story.append(Image(bud_mtd_chart, width=7.25*inch, height=2.7*inch))
        bud_mtd_rows = []
        for r in bud_mtd.itertuples(index=False):
            bud_mtd_rows.append([str(r.budtender), money(float(r.net_revenue)), f"{int(r.tickets):,}", money2(float(r.basket)), pct1(float(r.discount_rate))])
        story.append(build_table(
            ["Budtender", "MTD Net", "MTD Tickets", "MTD Basket", "Disc Rate"],
            bud_mtd_rows,
            [2.65*inch, 1.25*inch, 1.05*inch, 1.25*inch, 1.15*inch],
        ))

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    print(f"✅ PDF created: {out_pdf}")


###############################################################################
# PDF: All stores summary (kept simple but consistent)
###############################################################################

def build_all_stores_summary_pdf(out_pdf: Path, store_daily_map: Dict[str, pd.DataFrame], end_day: date, start_day: date) -> None:
    styles = build_styles()
    generated_at = datetime.now(ZoneInfo(REPORT_TZ)).strftime("%B %d, %Y at %I:%M %p %Z")
    mtd_start = month_start(end_day)
    last_week_day = end_day - timedelta(days=7)
    last_mtd_start, last_mtd_end = prev_month_same_window(end_day)

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(out_pdf),
        pagesize=letter,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title=f"All Stores Owner Snapshot - {end_day.isoformat()}",
        author="Buzz Automation",
    )

    footer = make_footer("ALL STORES", end_day)
    story: List[Any] = []

    story.append(Paragraph("All Stores • Owner Snapshot", styles["TitleBig"]))
    story.append(build_report_day_band(end_day, width=7.6 * inch))
    story.append(Spacer(1, SPACER["xs"]))

    story.append(Paragraph(
        f"<b>Data Window:</b> {start_day.isoformat()} → {end_day.isoformat()} &nbsp;&nbsp; "
        f"<b>MTD Window:</b> {mtd_start.isoformat()} → {end_day.isoformat()} &nbsp;&nbsp; "
        f"<b>Last MTD Ref:</b> {last_mtd_start.isoformat()} → {last_mtd_end.isoformat()}",
        styles["Tiny"],
    ))
    story.append(Paragraph(f"<b>Generated:</b> {generated_at}", styles["Tiny"]))
    story.append(Spacer(1, SPACER["sm"]))

    rows = []
    totals_today_net = totals_today_tickets = totals_mtd_net = totals_mtd_tickets = 0.0
    store_rank = []

    for store_name, abbr in store_abbr_map.items():
        daily = store_daily_map.get(abbr)
        if daily is None or daily.empty:
            continue

        today = metrics_for_day(daily, end_day)
        last_week = metrics_for_day(daily, last_week_day)
        mtd = metrics_for_range(daily, mtd_start, end_day)
        last_mtd = metrics_for_range(daily, last_mtd_start, last_mtd_end)

        totals_today_net += today["net_revenue"]
        totals_today_tickets += today["tickets"]
        totals_mtd_net += mtd["net_revenue"]
        totals_mtd_tickets += mtd["tickets"]

        d_today = today["net_revenue"] - last_week["net_revenue"]
        d_mtd = mtd["net_revenue"] - last_mtd["net_revenue"]

        rows.append([
            f"{abbr} - {store_label(store_name)}",
            money(today["net_revenue"]),
            fmt_signed_money(d_today),
            money(mtd["net_revenue"]),
            fmt_signed_money(d_mtd),
            pct1(mtd["margin"]),
            f"{int(mtd['tickets']):,}",
        ])
        store_rank.append([f"{abbr} - {store_label(store_name)}", float(mtd["net_revenue"])])

    story.append(Paragraph(
        f"<b>Totals Today:</b> {money(totals_today_net)} • {int(totals_today_tickets):,} tickets"
        f"&nbsp;&nbsp; <b>Totals MTD:</b> {money(totals_mtd_net)} • {int(totals_mtd_tickets):,} tickets",
        styles["Muted"],
    ))
    story.append(Spacer(1, SPACER["sm"]))

    if store_rank:
        rank_df = pd.DataFrame(store_rank, columns=["store", "net_revenue"]).sort_values("net_revenue", ascending=False)
        rank_chart = chart_rank_barh(rank_df, "store", "net_revenue", "Store Ranking (MTD Net Sales)", top_n=min(10, len(rank_df)), figsize=(7.25, 2.7))
        story.append(KeepTogether([
            Paragraph("MTD Ranking", styles["Section"]),
            Image(rank_chart, width=7.25 * inch, height=2.7 * inch),
        ]))
        story.append(Spacer(1, SPACER["sm"]))

    story.append(Paragraph("Store Table", styles["Section"]))
    story.append(build_table(
        headers=["Store", "Today Net", "Δ vs LW", "MTD Net", "Δ vs Last MTD", "MTD Margin", "MTD Tickets"],
        rows=rows,
        col_widths=[2.70*inch, 0.9*inch, 0.85*inch, 0.9*inch, 1.05*inch, 0.85*inch, 0.75*inch],
    ))

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    print(f"✅ All-stores summary PDF created: {out_pdf}")


###############################################################################
# MAIN
###############################################################################

def main():
    setup_fonts()

    REPORTS_ROOT.mkdir(parents=True, exist_ok=True)
    RAW_ROOT.mkdir(parents=True, exist_ok=True)
    PDF_ROOT.mkdir(parents=True, exist_ok=True)

    abbr_to_file: Dict[str, Path] = {}

    if RUN_EXPORT:
        print("⚠️ RUN_EXPORT=True → Running Selenium export")
        start_day, end_day = compute_date_window(BACKFILL_DAYS, REPORT_TZ)
        run_export_for_range(start_day, end_day)
        _, abbr_to_file = archive_exports(start_day, end_day)
    else:
        print("✅ RUN_EXPORT=False → Reusing latest raw export folder")
        raw_folder = find_latest_raw_folder()
        if raw_folder is None:
            raise SystemExit("No raw export folders found in reports/raw_sales and RUN_EXPORT=False.")

        parsed = parse_range_from_folder_name(raw_folder)
        if parsed:
            start_day, end_day = parsed
        else:
            start_day, end_day = compute_date_window(BACKFILL_DAYS, REPORT_TZ)

        for store_name, abbr in store_abbr_map.items():
            matches = list(raw_folder.glob(f"{abbr}*Sales Export*.xlsx"))
            if matches:
                abbr_to_file[abbr] = matches[0]

    if not abbr_to_file:
        raise SystemExit("No store exports found. Check getSalesReport output /files or raw archive.")

    # Parse and compute daily
    store_daily_map: Dict[str, pd.DataFrame] = {}
    store_raw_df_map: Dict[str, pd.DataFrame] = {}

    for store_name, abbr in store_abbr_map.items():
        path = abbr_to_file.get(abbr)
        if not path:
            continue

        print(f"[PARSE] {abbr}: {path.name}")
        df = read_export(path)
        store_raw_df_map[abbr] = df

        daily = compute_daily_metrics(df)
        daily = daily[(daily["date"] >= start_day) & (daily["date"] <= end_day)]
        store_daily_map[abbr] = daily

    # Output folder for this run
    pdf_run_dir = PDF_ROOT / f"{start_day.isoformat()}_to_{end_day.isoformat()}"
    pdf_run_dir.mkdir(parents=True, exist_ok=True)

    # Per-store PDFs
    for store_name, abbr in store_abbr_map.items():
        daily = store_daily_map.get(abbr)
        df_raw = store_raw_df_map.get(abbr)
        if daily is None or daily.empty or df_raw is None:
            print(f"[SKIP] {abbr} missing data")
            continue

        out_pdf = pdf_run_dir / safe_filename(
            f"{abbr} - Owner Snapshot - {store_label(store_name)} - {end_day.isoformat()}.pdf"
        )
        build_store_pdf(out_pdf, store_name, abbr, df_raw, daily, start_day, end_day)

    # All stores summary
    if GENERATE_ALL_STORES_SUMMARY_PDF:
        out_pdf = pdf_run_dir / safe_filename(f"ALL STORES - Owner Snapshot - {end_day.isoformat()}.pdf")
        build_all_stores_summary_pdf(out_pdf, store_daily_map, end_day=end_day, start_day=start_day)
    # Collect PDFs from this run
    pdfs = sorted(str(p) for p in pdf_run_dir.glob("*.pdf"))

    send_owner_snapshot_email(
        pdf_paths=pdfs,
        report_day=end_day,
        data_start=start_day,
        data_end=end_day,
        to_email="anthony@buzzcannabis.com",
    )
    print("\nDone ✅")


if __name__ == "__main__":
    main()
