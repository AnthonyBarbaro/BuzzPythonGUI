#!/usr/bin/env python3
import pandas as pd
import pprint
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

CSV_PATH = r"files\11-28-2025_MV.csv"

# Column names in that CSV (adjust if yours differ)
BRAND_COLUMN   = "Brand"
VENDOR_COLUMN  = "Vendor"
PRODUCT_COLUMN_CANDIDATES = ["Product", "Product Name", "product name"]

ALL_DAYS = [
    "Monday", "Tuesday", "Wednesday",
    "Thursday", "Friday", "Saturday", "Sunday"
]

DEFAULT_DISCOUNT = 0.50   # 50% off
DEFAULT_KICKBACK = 0.30   # 30% kickback

# These are the logical brand names / keys in your criteria dict
BRANDS_TO_BUILD = [
    "American Weed",
    "Sluggers",
    "BLEM",
    "Ball Family Farms",
    "Turtle Pie Co.",
    "Eureka",
    "Treesap",
    "Drops",
    "Quiet Kings",
    "Pure Beauty",
    "Not Your Father",
    "Dab Daddy",
    "Dabwoods",
    "PBR (Pabst)",
    "Green Dawg",
    "Nasha",
    "Claybourne Co.",
    "St. Ides",
    "Cannabiotix (CBX)",
    "Keef",
    "Turn",
    "Yada Yada",
    "Emerald Sky",
    "Hashish",
    "Jeeter",
    "CLSICS",
    "Dr. Norms",
    "Pearl Pharma",
    "Almora Farm",
    "Smokiez",
    "Kushy Punch",
    "Seed Junky",
    "Papa & Barkley (P&B)",
    "Uncle Arnie's",
    "Punch",
    "KANHA",
    "Pacific Stone",
    "Heavy Hitters",
    "Stiiizy",
]


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def load_catalog(csv_path: str) -> pd.DataFrame:
    """
    Read the MV menu CSV and ensure Brand/Vendor/Product columns exist.
    """
    path = Path(csv_path)
    if not path.is_file():
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    df = pd.read_csv(path)

    # Check Brand/Vendor columns
    missing = [c for c in (BRAND_COLUMN, VENDOR_COLUMN) if c not in df.columns]
    if missing:
        raise ValueError(
            f"CSV is missing required columns: {missing}. "
            f"Available columns: {list(df.columns)}"
        )

    # Find product-name column
    product_col = None
    for cand in PRODUCT_COLUMN_CANDIDATES:
        if cand in df.columns:
            product_col = cand
            break
    if product_col is None:
        raise ValueError(
            f"Could not find a product-name column. Tried: {PRODUCT_COLUMN_CANDIDATES}. "
            f"Available columns: {list(df.columns)}"
        )

    # Normalize strings
    df[BRAND_COLUMN]  = df[BRAND_COLUMN].astype(str).str.strip()
    df[VENDOR_COLUMN] = df[VENDOR_COLUMN].astype(str).str.strip()
    df[product_col]   = df[product_col].astype(str).str.strip()

    # Rename for easier handling
    df = df.rename(columns={product_col: "ProductName"})
    return df


def tokens_from_product_names(product_series: pd.Series) -> list[str]:
    """
    Given a Series of product names, return sorted unique brand tokens based on:

        token = part before the first '|', stripped of whitespace

    So for "CBX | Rosin 1G | T2..." → "CBX"
    """
    tokens = set()

    for raw in product_series.dropna():
        s = str(raw)
        if "|" in s:
            first = s.split("|", 1)[0].strip()
        else:
            first = s.strip()
        if first:
            tokens.add(first)

    return sorted(tokens)


def build_brand_criteria_from_df(df: pd.DataFrame,
                                 brands: list[str]) -> dict:
    """
    Build a brand_criteria-style dict compatible with your deals script:

        brand_criteria = {
            'Cannabiotix (CBX)': {
                'vendors': [...],
                'days': [...],
                'discount': 0.50,
                'kickback': 0.30,
                'brands': ['CBX'],   # <-- derived from ProductName ("CBX | ...")
            },
            ...
        }
    """
    criteria: dict[str, dict] = {}

    for brand_key in brands:
        # Match logical brand key against Brand column
        mask = df[BRAND_COLUMN].str.casefold() == brand_key.casefold()
        sub = df[mask]

        # Vendors for this brand
        if not sub.empty:
            vendors = sorted(
                v for v in sub[VENDOR_COLUMN].dropna().unique().tolist()
                if str(v).strip() not in ("", "nan")
            )
            brand_tokens = tokens_from_product_names(sub["ProductName"])
        else:
            vendors = []
            brand_tokens = []

        # If we couldn't derive tokens, fall back to the brand key itself
        if not brand_tokens:
            brand_tokens = [brand_key]

        # This is what your deals script uses against 'product name'
        criteria[brand_key] = {
            "vendors": vendors,
            "days": ALL_DAYS[:],
            "discount": DEFAULT_DISCOUNT,
            "kickback": DEFAULT_KICKBACK,
            "brands": brand_tokens,
        }

    return criteria


def main():
    df = load_catalog(CSV_PATH)
    crit = build_brand_criteria_from_df(df, BRANDS_TO_BUILD)

    print("# === AUTO-GENERATED brand_criteria (50% / 30% Everyday) ===")
    pprint.pprint(crit, sort_dicts=False)


if __name__ == "__main__":
    main()
