"""
Feedback Correction AI-Agent Style Tool

Purpose:
- Read a multi-month data workbook.
- Read a feedback workbook.
- Refine/standardize the feedback text into a corrected feedback output file.
- Apply safe corrections into the correct data workbook month/sheet.
- Keep the logic query-first, so we only correct what was actually asked.

Inputs:
- DATA_PATH
- FEEDBACK_PATH
- OUTPUT_DIR

Outputs:
- Corrected Data File
- Corrected Feedback File

Important rules implemented:
- Only 3 main decision classes:
    1. Corrected
    2. Query Not Answered
    3. Not Corrected
- Previous stock/current previous-stock handling:
    If feedback says previous stock is X:
      a) Add feedback tracking in the current month.
      b) Correct current month's Previous Stock column if it exists.
      c) Still go to the previous month and correct the configured stock column there.
- If correction belongs to a previous month:
    The current month still receives the feedback and a status such as "Corrected in April".
    The actual target month receives status "Corrected".
- If feedback mentions a month but gives no clear value for that month:
    The feedback is still tracked, but status becomes "Not Corrected".
- If query asks one metric but feedback answers another metric:
    Status becomes "Query Not Answered".
- Prices cannot be negative.
- Multiple price columns are handled using project_configs.py and row-level logic.

Keep project_configs.py in the same folder as this script.
"""

import os
import re
import time
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

try:
    from project_configs import PROJECT_CONFIGS
except Exception:
    PROJECT_CONFIGS = {}


# ═══════════════════════════════════════════════════════════════
# CONFIG — update these before running
# ═══════════════════════════════════════════════════════════════

PROJECT_NAME = "Usafi-Uganda"  # Must match project_configs.py where possible
DATA_PATH = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\Usafi-Uganda\Usafi-Uganda-Data-Files\2026\May2026\Usafi_Uganda_Data_May26-Batch 4.xlsx"
FEEDBACK_PATH = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\Usafi-Uganda\Usafi-Uganda-Feedbacks\May2026\Usafi Uganda-May2026-4-Queries Feedback.xlsx"
OUTPUT_DIR = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\Usafi-Uganda\Usafi-Uganda-Feedbacks\May2026-2"

PREFERRED_FEEDBACK_SHEET = "Data Queries"

QUERIES_COL = "Queries"
FEEDBACK_COL = "Feedback"
STATUS_COL = "Correction Status"

STATUS_CORRECTED = "Corrected"
STATUS_QUERY_NOT_ANSWERED = "Query Not Answered"
STATUS_NOT_CORRECTED = "Not Corrected"


# ═══════════════════════════════════════════════════════════════
# BASIC NORMALIZATION
# ═══════════════════════════════════════════════════════════════


def normalize(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_blank(value: Any) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip().lower() in ["", "nan", "none", "null"]


def make_key(outlet: Any, sku: Any) -> str:
    return f"{clean_text(outlet)}{clean_text(sku)}"


def append_tracking(existing: Any, new_text: str) -> str:
    existing_text = clean_text(existing)
    new_text = clean_text(new_text)
    if not new_text:
        return existing_text
    if not existing_text:
        return new_text
    if new_text in existing_text:
        return existing_text
    return existing_text + " | " + new_text


def format_value(value: Any) -> Any:
    if value is None or pd.isna(value):
        return np.nan
    try:
        f = float(value)
        if f.is_integer():
            return int(f)
        return f
    except Exception:
        return value


# ═══════════════════════════════════════════════════════════════
# COLUMN RESOLUTION
# ═══════════════════════════════════════════════════════════════


def resolve_col(df: pd.DataFrame, candidates: Any) -> Optional[str]:
    if candidates is None:
        return None
    if isinstance(candidates, str):
        candidates = [candidates]

    columns = list(df.columns)
    norm_map = {normalize(c): c for c in columns}

    # Exact and normalized match first.
    for cand in candidates:
        if not cand:
            continue
        if cand in df.columns:
            return cand
        n = normalize(cand)
        if n in norm_map:
            return norm_map[n]

    # Flexible contains match.
    for cand in candidates:
        if not cand:
            continue
        cn = normalize(cand)
        for col in columns:
            n = normalize(col)
            if cn and (cn in n or n in cn):
                return col
    return None


def resolve_all_cols(df: pd.DataFrame, candidates: List[str]) -> List[str]:
    found: List[str] = []
    seen = set()
    for cand in candidates:
        col = resolve_col(df, cand)
        if col and col not in seen:
            found.append(col)
            seen.add(col)
    return found


def resolve_feedback_columns(feedback_df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Feedback file is controlled by field teams, so resolve flexibly."""
    fb_outlet = resolve_col(feedback_df, [
        "Outlet ID", "OutletID", "Outlet Id", "Outlet Number", "Out Number", "OUTNUMBER",
        "outletid", "outlet_id", "wh_outletid", "projectOutletid", "Outlet",
    ])
    fb_sku = resolve_col(feedback_df, [
        "SKU ID", "Sku ID", "SKU_ID", "SKU", "Sku", "Prod Code", "Prodcode", "Product Code",
        "product code", "wh_skuid", "sku_id",
    ])
    fb_query = resolve_col(feedback_df, [
        "Queries", "Query", "Question", "Questions", "Issue", "Title", "Data Query", "Data Queries",
    ])
    fb_feedback = resolve_col(feedback_df, [
        "Feedback", "Feedbacks", "Comment", "Comments", "Response", "Correction", "Field Feedback",
    ])
    return {
        "fb_outlet": fb_outlet,
        "fb_sku": fb_sku,
        "fb_query": fb_query,
        "fb_feedback": fb_feedback,
    }


def resolve_data_columns(data_df: pd.DataFrame, project_name: str) -> Dict[str, Any]:
    """Data file is guided mainly by project_configs.py."""
    config = PROJECT_CONFIGS.get(project_name, {}) if PROJECT_CONFIGS else {}

    data_outlet = resolve_col(data_df, [
        config.get("outlet_id"), "Outlet ID", "Out Number", "Outlet Number", "outletid", "OutletID",
        "wh_outletid", "projectOutletid", "outlet_id", "OUTNUMBER",
    ])
    data_sku = resolve_col(data_df, [
        config.get("sku_id"), "SKU ID", "Prod Code", "Prodcode", "SKU_ID", "wh_skuid", "Sku ID",
        "product code", "sku_id",
    ])
    data_stock = resolve_col(data_df, [
        config.get("stock_col"), "Total Stock", "Stock", "Current Stock", "Total_Stock", "TOTALSTOCK",
    ])
    data_purchase = resolve_col(data_df, [
        config.get("purchase_col"), "Purchases", "Purchase", "Total Purchases", "Total Purchase",
        "Stock Increase", "PURCHASES",
    ])

    previous_stock_cols = resolve_all_cols(data_df, [
        "Previous Stock", "Prev Stock", "Previous Month Stock", "Last Month Stock",
        "Opening Stock", "Previous Total Stock", "Prev Total Stock",
    ])
    data_previous_stock = previous_stock_cols[0] if previous_stock_cols else None

    price_candidates: List[str] = []
    for p in config.get("price_cols", []) or []:
        price_candidates.append(p)
    if config.get("buying_price_col"):
        price_candidates.append(config.get("buying_price_col"))
    price_candidates += [
        "Selling Price", "Selling Price per Sku", "Price", "Retail Price", "Buying Price", "Buying Price per Sku",
        "Capture Price", "Capture Price Excl Container", "Capture Price Excluding Container",
        "Capture price excluding the container", "Capture price including the container",
    ]

    data_price_cols = resolve_all_cols(data_df, price_candidates)
    data_price = data_price_cols[0] if data_price_cols else None

    return {
        "data_outlet": data_outlet,
        "data_sku": data_sku,
        "data_stock": data_stock,
        "data_previous_stock": data_previous_stock,
        "data_purchase": data_purchase,
        "data_price": data_price,
        "data_price_cols": data_price_cols,
        "project_config": config,
    }


def resolve_feedback_value_columns(feedback_df: pd.DataFrame, data_cols: Dict[str, Any]) -> Dict[str, List[str]]:
    """Detect possible corrected values already typed into feedback file columns."""
    price_candidates = list(data_cols.get("data_price_cols", []) or []) + [
        "Selling Price", "Selling Price per Sku", "Price", "Retail Price", "Buying Price", "Buying Price per Sku",
        "Capture Price", "Capture Price Excl Container", "Capture Price Excluding Container",
    ]
    return {
        "price": resolve_all_cols(feedback_df, price_candidates),
        "purchase": resolve_all_cols(feedback_df, [
            data_cols.get("data_purchase"), "Purchases", "Purchase", "Total Purchases", "Total Purchase", "Stock Increase",
        ]),
        "stock": resolve_all_cols(feedback_df, [
            data_cols.get("data_stock"), "Total Stock", "Stock", "Current Stock", "Total_Stock", "TOTALSTOCK",
        ]),
        "previous_stock": resolve_all_cols(feedback_df, [
            "Previous Stock", "Prev Stock", "Previous Month Stock", "Last Month Stock", "Opening Stock",
        ]),
    }


# ═══════════════════════════════════════════════════════════════
# EXCEL LOADING
# ═══════════════════════════════════════════════════════════════


def read_excel_sheet(path: str, preferred_sheet: Optional[str] = None) -> Tuple[pd.DataFrame, str]:
    xl = pd.ExcelFile(path)
    sheet = None
    if preferred_sheet:
        for s in xl.sheet_names:
            if normalize(s) == normalize(preferred_sheet):
                sheet = s
                break
    if sheet is None:
        sheet = xl.sheet_names[0]

    # Try normal header first.
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    unnamed = sum(1 for c in df.columns if str(c).lower().startswith("unnamed"))
    if unnamed > max(2, len(df.columns) // 2):
        # Some feedback files have the useful header on row 2.
        df = pd.read_excel(path, sheet_name=sheet, header=1)
        df.columns = [str(c).strip() for c in df.columns]

    return df, sheet


def load_data_workbook(path: str) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    xl = pd.ExcelFile(path)
    sheets: Dict[str, pd.DataFrame] = {}
    for s in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=s)
        df.columns = [str(c).strip() for c in df.columns]
        sheets[s] = df
    return sheets, xl.sheet_names


# ═══════════════════════════════════════════════════════════════
# MONTH DETECTION
# ═══════════════════════════════════════════════════════════════


MONTH_ALIASES = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

MONTH_NAMES = {
    1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June",
    7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December",
}


def month_number_from_text(text: str) -> Optional[int]:
    t = str(text).lower()
    for name, num in MONTH_ALIASES.items():
        if re.search(rf"\b{re.escape(name)}\b", t):
            return num
    return None


def months_mentioned_in_text(text: str) -> List[int]:
    t = str(text).lower()
    found: List[int] = []
    # Sort longer names first so "march" is handled before "mar".
    for name, num in sorted(MONTH_ALIASES.items(), key=lambda x: -len(x[0])):
        if re.search(rf"\b{re.escape(name)}\b", t) and num not in found:
            found.append(num)
    return found


def sheet_month_number(sheet_name: str) -> Optional[int]:
    return month_number_from_text(sheet_name)


def find_sheet_for_month(month_num: int, sheet_names: List[str]) -> Optional[str]:
    for s in sheet_names:
        if sheet_month_number(s) == month_num:
            return s
    return None


def previous_sheet(sheet_names: List[str], current_sheet: str, back: int = 1) -> Optional[str]:
    try:
        idx = sheet_names.index(current_sheet)
    except ValueError:
        return None
    target_idx = idx - back
    return sheet_names[target_idx] if target_idx >= 0 else None


def display_month_name(sheet: str) -> str:
    n = sheet_month_number(sheet)
    return MONTH_NAMES.get(n, sheet)


# ═══════════════════════════════════════════════════════════════
# FEEDBACK REFINER
# ═══════════════════════════════════════════════════════════════


SPELL_FIXES = [
    (r"\bbougt\b", "bought"),
    (r"\bboight\b", "bought"),
    (r"\brespodent\b", "respondent"),
    (r"\bRespondant\b", "Respondent"),
    (r"\bstockwd\b", "stocked"),
    (r"\bstockd\b", "stocked"),
    (r"\bcompletly\b", "completely"),
    (r"\bafew\b", "a few"),
    (r"\balot\b", "a lot"),
    (r"\bresturaunt\b", "restaurant"),
    (r"\bresturant\b", "restaurant"),
    (r"\bhiden\b", "hidden"),
    (r"\batock\b", "stock"),
    (r"\bztock\b", "stock"),
    (r"\bpurchse\b", "purchase"),
    (r"\bpurchsed\b", "purchased"),
    (r"\bhad ran\b", "had run"),
    (r"\bran low\b", "run low"),
    (r"\bprefering\b", "preferring"),
    (r"!\s*(\d)", r"1\1"),
    (r"\bOUtlet\b", "Outlet"),
]


def fix_spelling(text: str) -> str:
    for pattern, replacement in SPELL_FIXES:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return text


def sentence_case(text: str) -> str:
    text = clean_text(text)
    if not text:
        return text
    text = text[0].upper() + text[1:]
    if text[-1] not in ".!?":
        text += "."
    return text


def extract_positive_integer_text(text: str) -> Optional[str]:
    m = re.search(r"(\d[\d,]*)", str(text))
    return m.group(1).replace(",", "") if m else None


def is_shorthand_feedback(text: str) -> bool:
    t = clean_text(text)
    patterns = [
        r"^[\-–—]?\s*[\d,]+$",
        r"^[\-–—]?\s*[\d,]+\s*(shs|ugx|ksh|kes|naira|ngn)\.?$",
        r"^(shs|ugx|ksh|kes|naira|ngn)\s*[\-–—]?\s*[\d,]+\.?$",
        r"^[\d,]+\s*\/?=?$",
        r"^price[-\s:]*[\d,]+\.?$",
        r"^selling\s+price\s+is\s*[\d,]+\/?=?$",
        r"^correct\s+(amount|price)\s+is\s+[\d,]+$",
        r"^[\d,]+\s*ugx\.?$",
    ]
    return any(re.match(p, t, re.IGNORECASE) for p in patterns)


def infer_query_metrics(query_text: str) -> List[str]:
    q = clean_text(query_text).lower()
    metrics: List[str] = []

    # Negative sales is special; it can be answered by stock, previous stock, or purchases.
    if re.search(r"negative\s+sales|neg\s*sales", q):
        metrics.append("negative_sales")
        return metrics

    if re.search(r"selling|price|capture price|retail price|buying price", q):
        metrics.append("price")
    if re.search(r"purchase|purchased|stock increase|bought|buy", q):
        metrics.append("purchase")
    if re.search(r"previous\s+stock|prev\s+stock", q):
        metrics.append("previous_stock")
    elif re.search(r"stock|count|available|unit count|total stock|current stock", q):
        metrics.append("stock")

    return list(dict.fromkeys(metrics))


def infer_feedback_metrics(feedback_text: str) -> List[str]:
    t = clean_text(feedback_text).lower()
    metrics: List[str] = []

    if re.search(r"previous\s+stock|prev\s+stock|opening\s+stock|last\s+month\s+stock", t):
        metrics.append("previous_stock")
    if re.search(r"selling\s+price|capture\s+price|retail\s+price|buying\s+price|\bprice\b|ugx|shs|ksh|kes|naira|ngn", t):
        metrics.append("price")
    if re.search(r"purchase|purchased|bought|buy|stock increase", t):
        metrics.append("purchase")
    if re.search(r"total\s+stock|current\s+stock|front\s+stock|back\s+stock|\bstock\b|counted|available|unit count", t):
        # Avoid adding ordinary stock if the phrase is only previous stock.
        if "previous_stock" not in metrics or re.search(r"total\s+stock|current\s+stock|front\s+stock|back\s+stock", t):
            metrics.append("stock")

    return list(dict.fromkeys(metrics))


def query_allows_feedback_metric(query_metrics: List[str], feedback_metric: str) -> bool:
    if not query_metrics:
        # If query is missing/unclear, allow feedback parsing but stay conservative later.
        return True
    if "negative_sales" in query_metrics:
        return feedback_metric in ["purchase", "stock", "previous_stock"]
    if feedback_metric == "previous_stock":
        return "previous_stock" in query_metrics or "stock" in query_metrics
    return feedback_metric in query_metrics


def refine_feedback(query_text: str, feedback_text: str) -> str:
    q = clean_text(query_text)
    f = clean_text(feedback_text)
    if not f:
        return f

    query_metrics = infer_query_metrics(q)
    has_price = "price" in query_metrics
    has_purchase = "purchase" in query_metrics or "negative_sales" in query_metrics
    has_stock = "stock" in query_metrics or "previous_stock" in query_metrics or "negative_sales" in query_metrics

    # Selling price shorthand.
    if is_shorthand_feedback(f) and has_price:
        n = extract_positive_integer_text(f)
        return f"The correct selling price is {n}." if n else sentence_case(fix_spelling(f))

    m = re.match(r"^Price[-\s:]*([\d,]+)[.,]?\s*(.*)$", f, re.IGNORECASE)
    if m and has_price:
        rest = clean_text(m.group(2))
        base = f"The correct selling price is {m.group(1).replace(',', '')}."
        return f"{base} {sentence_case(fix_spelling(rest))}" if rest else base

    m = re.match(r"^([\d,]+)\s+selling price[,.]?\s*(.+)$", f, re.IGNORECASE)
    if m and has_price:
        return f"The correct selling price is {m.group(1).replace(',', '')}. {sentence_case(fix_spelling(m.group(2)))}"

    m = re.match(r"^selling\s+price\s+is\s*([\d,]+)\/?=?$", f, re.IGNORECASE)
    if m and has_price:
        return f"The correct selling price is {m.group(1).replace(',', '')}."

    m = re.match(r"^correct\s+(?:amount|price)\s+is\s+([\d,]+)$", f, re.IGNORECASE)
    if m and has_price:
        return f"The correct selling price is {m.group(1).replace(',', '')}."

    m = re.match(r"^typing error,?\s+correct price is\s+([\d,]+)$", f, re.IGNORECASE)
    if m and has_price:
        return f"There was a typing error; the correct selling price is {m.group(1).replace(',', '')}."

    m = re.search(r"selling price is\s+([\d,]+)", f, re.IGNORECASE)
    if m and has_price:
        rest = re.sub(r"\s*(?:and\s+)?(?:its\s+)?selling price is\s+[\d,]+", "", f, flags=re.IGNORECASE).strip()
        base = f"The correct selling price is {m.group(1).replace(',', '')}."
        return f"{sentence_case(fix_spelling(rest))} {base}" if rest else base

    m = re.match(r"^([\d,]+)\s*(ugx|shs|ksh|kes|naira|ngn)\.?$", f, re.IGNORECASE)
    if m and has_price:
        return f"The correct selling price is {m.group(1).replace(',', '')}."

    # Purchases.
    m = re.match(r"^[Pp]urchased?\s+([\d,]+)\s*(units?|pieces?|packets?|sachets?|buckets?)?\s*\.?$", f)
    if m and has_purchase:
        unit = (m.group(2) or "units")
        return f"The respondent purchased {m.group(1).replace(',', '')} {unit}."

    m = re.match(r"^[Pp]urchased?\s+([\d,]+)\s+in\s+week\s+(\d+)\.?$", f)
    if m and has_purchase:
        return f"The respondent purchased {m.group(1).replace(',', '')} units in week {m.group(2)}."

    m = re.match(r"^([\d,]+)\s+purchased\s+in\s+week\s+(\d+)\.?$", f, re.IGNORECASE)
    if m and has_purchase:
        return f"The respondent purchased {m.group(1).replace(',', '')} units in week {m.group(2)}."

    m = re.match(r"^purchases?\s+(?:made\s+)?were\s+([\d,]+)\s*\.?$", f, re.IGNORECASE)
    if m and has_purchase:
        return f"The respondent made a purchase of {m.group(1).replace(',', '')} units."

    m = re.match(r"^([\d,]+)\s+(?:bought|purchased)\.?$", f, re.IGNORECASE)
    if m and has_purchase:
        return f"The respondent purchased {m.group(1).replace(',', '')} units."

    m = re.match(r"^[Tt]hey\s+purchased\s+([\d,]+)\s*(units?|pieces?)?\s*\.?$", f)
    if m and has_purchase:
        unit = m.group(2) or "units"
        return f"The respondent purchased {m.group(1).replace(',', '')} {unit}."

    m = re.match(r"^([\d,]+)\s+(pieces?|units?|sachets?|packets?)\s+were\s+purchased\.?$", f, re.IGNORECASE)
    if m and has_purchase:
        return f"The respondent purchased {m.group(1).replace(',', '')} {m.group(2)}."

    m = re.match(r"^[Pp]urchase\s+is\s+([\d,]+)\.?$", f)
    if m and has_purchase:
        return f"The respondent purchased {m.group(1).replace(',', '')} units."

    # Previous stock.
    m = re.match(r"^[Pp]rev(?:ious)?\s+[A-Za-z]?tock\s+(?:is|was)\s*([\d,]+)\.?$", f, re.IGNORECASE)
    if m and has_stock:
        return f"The previous stock was {m.group(1).replace(',', '')} units."

    m = re.match(r"^[Pp]rev(?:ious)?\s+stock\s+(?:is|was)\s*([\d,]+)\.?$", f, re.IGNORECASE)
    if m and has_stock:
        return f"The previous stock was {m.group(1).replace(',', '')} units."

    m = re.match(r"^[Pp]rev\s+([\d,]+)\.?$", f)
    if m and has_stock:
        return f"The previous stock was {m.group(1).replace(',', '')} units."

    m = re.match(r"^[Ss]tock\s+is\s+([\d,]+)(?:\s+pieces?)?\s*\.?$", f)
    if m and has_stock:
        return f"The stock was {m.group(1).replace(',', '')} units."

    fixed = fix_spelling(f)
    fixed = re.sub(r"\b1 units\b", "1 unit", fixed)
    fixed = re.sub(r"\b1 pieces\b", "1 piece", fixed)
    return sentence_case(fixed)


# ═══════════════════════════════════════════════════════════════
# NUMBER EXTRACTION
# ═══════════════════════════════════════════════════════════════


def extract_first_number(text: str, allow_negative: bool = False) -> Optional[float]:
    cleaned = clean_text(text).replace(",", "")
    if allow_negative:
        m = re.search(r"#?\s*(-?\d+(?:\.\d+)?)", cleaned)
    else:
        # Treat hyphen before a number as punctuation, not as negative.
        m = re.search(r"#?\s*(\d+(?:\.\d+)?)", cleaned)
    if not m:
        return None
    try:
        return float(m.group(1))
    except Exception:
        return None


def extract_number_near_keywords(text: str, keywords: List[str], allow_negative: bool = False) -> Optional[float]:
    t = clean_text(text).lower().replace(",", "")
    sign = r"-?" if allow_negative else r""

    # keyword before number: price is 500, purchase 12, stock was 8.
    for kw in keywords:
        pattern = rf"\b{re.escape(kw)}\w*\b[^0-9#]{{0,50}}#?\s*{sign}(\d+(?:\.\d+)?)"
        m = re.search(pattern, t)
        if m:
            return float(m.group(1))

    # number before keyword: 12 units purchased.
    for kw in keywords:
        pattern = rf"#?\s*{sign}(\d+(?:\.\d+)?)\s*(?:units?|pieces?|buckets?|packets?|sachets?)?[^a-z0-9]{{0,25}}\b{re.escape(kw)}\w*\b"
        m = re.search(pattern, t)
        if m:
            return float(m.group(1))

    return None


def safe_numeric(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        if pd.isna(value):
            return None
        return float(value)
    text = clean_text(value).replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", text)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════════
# DECISION OBJECTS
# ═══════════════════════════════════════════════════════════════


@dataclass
class Action:
    metric: str
    value: Optional[float]
    target_sheet: str
    target_column_kind: str  # price | purchase | stock | previous_stock
    apply_value: bool
    current_sheet_status: str
    target_sheet_status: str
    reason: str


@dataclass
class Decision:
    main_status: str
    reason: str
    refined_feedback: str
    query_text: str
    original_feedback: str
    current_sheet_status: str
    actions: List[Action] = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════
# DECISION ENGINE
# ═══════════════════════════════════════════════════════════════


def get_feedback_column_value(
    fb_row: pd.Series,
    metric: str,
    feedback_value_cols: Dict[str, List[str]],
    preferred_names: Optional[List[str]] = None,
) -> Optional[float]:
    cols = feedback_value_cols.get(metric, [])
    if preferred_names:
        preferred_norms = {normalize(x) for x in preferred_names}
        cols = sorted(cols, key=lambda c: 0 if normalize(c) in preferred_norms else 1)
    for col in cols:
        if col in fb_row.index:
            n = safe_numeric(fb_row.get(col))
            if n is not None:
                return n
    return None


def parse_metric_value_from_text(metric: str, text: str, query_metrics: List[str]) -> Optional[float]:
    t = clean_text(text)
    tl = t.lower()

    if metric == "price":
        value = extract_number_near_keywords(t, [
            "selling price", "capture price", "retail price", "buying price", "price", "naira", "ngn", "shs", "ugx", "ksh", "kes",
        ], allow_negative=False)
        if value is None and ("price" in query_metrics or is_shorthand_feedback(t)):
            value = extract_first_number(t, allow_negative=False)
        if value is not None and value < 0:
            return abs(value)
        return value

    if metric == "purchase":
        if re.search(r"\b(no|none|zero)\s+(new\s+)?purchases?\b|\bno\s+new\s+purchases?\b|\bno\s+purchase\b", tl):
            return 0.0
        return extract_number_near_keywords(t, [
            "purchase", "purchased", "bought", "buy", "stock increase",
        ], allow_negative=False)

    if metric in ["stock", "previous_stock"]:
        if metric == "previous_stock":
            value = extract_number_near_keywords(t, [
                "previous stock", "prev stock", "opening stock", "last month stock",
            ], allow_negative=False)
            if value is None and re.search(r"\bprev\b", tl):
                value = extract_first_number(t, allow_negative=False)
            return value

        return extract_number_near_keywords(t, [
            "total stock", "current stock", "front stock", "back stock", "stock", "counted", "available", "unit count",
        ], allow_negative=False)

    return None


def target_sheets_from_text(text: str, current_sheet: str, sheet_names: List[str]) -> List[str]:
    months = months_mentioned_in_text(text)
    sheets: List[str] = []
    for m in months:
        s = find_sheet_for_month(m, sheet_names)
        if s and s not in sheets:
            sheets.append(s)

    if sheets:
        return sheets

    tl = clean_text(text).lower()
    if re.search(r"\b(previous|last)\s+month\b", tl):
        prev = previous_sheet(sheet_names, current_sheet, 1)
        return [prev] if prev else [current_sheet]
    if re.search(r"\b(two months ago|other month)\b", tl):
        prev2 = previous_sheet(sheet_names, current_sheet, 2)
        return [prev2] if prev2 else [current_sheet]

    return [current_sheet]


def mentions_previous_stock(text: str) -> bool:
    tl = clean_text(text).lower()
    return bool(re.search(r"previous\s+stock|prev\s+stock|\bprev\b|opening\s+stock|last\s+month\s+stock", tl))


def looks_like_already_corrected(text: str) -> bool:
    tl = clean_text(text).lower()
    return bool(re.search(r"already\s+(been\s+)?corrected|has\s+been\s+corrected|corrected\s+in\s+the\s+columns?|updated\s+in\s+the\s+columns?", tl))


def choose_decision_metric(query_metrics: List[str], feedback_metrics: List[str]) -> Tuple[Optional[str], bool]:
    """Return metric and whether the feedback answered the query.

    Important distinction:
    - If feedback talks about a DIFFERENT metric, that is Query Not Answered.
    - If feedback gives a general explanation with no clear metric/value, that is Not Corrected,
      because the query may have been acknowledged but no correction value was supplied.
    """
    if not query_metrics:
        if len(feedback_metrics) == 1:
            return feedback_metrics[0], True
        return None, False

    # If feedback has no detectable metric, do not call it Query Not Answered yet.
    # Treat it as an attempted answer to the query; it will become Not Corrected
    # later if no clear correction value is found.
    if not feedback_metrics:
        if "negative_sales" in query_metrics:
            return "stock", True
        return query_metrics[0], True

    if "negative_sales" in query_metrics:
        # Negative sales may be answered through previous stock, stock, or purchase.
        for m in ["previous_stock", "stock", "purchase"]:
            if m in feedback_metrics:
                return m, True
        return None, False

    for m in feedback_metrics:
        if query_allows_feedback_metric(query_metrics, m):
            return m, True

    return None, False


def build_decision(
    fb_row: pd.Series,
    query_text: str,
    original_feedback: str,
    refined_feedback: str,
    current_sheet: str,
    sheet_names: List[str],
    feedback_value_cols: Dict[str, List[str]],
) -> Decision:
    query_metrics = infer_query_metrics(query_text)
    feedback_metrics = infer_feedback_metrics(refined_feedback or original_feedback)

    # If feedback is a bare number, infer its metric from a single query metric.
    if not feedback_metrics and len(query_metrics) == 1 and extract_first_number(refined_feedback or original_feedback) is not None:
        feedback_metrics = query_metrics[:]

    metric, answered = choose_decision_metric(query_metrics, feedback_metrics)

    if not answered or not metric:
        return Decision(
            main_status=STATUS_QUERY_NOT_ANSWERED,
            reason="Feedback does not answer the metric asked in the query",
            refined_feedback=refined_feedback,
            query_text=query_text,
            original_feedback=original_feedback,
            current_sheet_status=STATUS_QUERY_NOT_ANSWERED,
            actions=[],
        )

    # Special previous stock logic.
    if metric == "previous_stock" or mentions_previous_stock(refined_feedback):
        value = parse_metric_value_from_text("previous_stock", refined_feedback, query_metrics)
        # Only trust values from feedback file columns when the text clearly says the file/columns
        # were already corrected. Otherwise an ordinary original value in the feedback file can be
        # mistaken for a correction and incorrectly counted as Corrected.
        if value is None and looks_like_already_corrected(refined_feedback):
            value = get_feedback_column_value(fb_row, "previous_stock", feedback_value_cols)
            if value is None:
                value = get_feedback_column_value(fb_row, "stock", feedback_value_cols)

        if value is None:
            return Decision(
                main_status=STATUS_NOT_CORRECTED,
                reason="Previous stock was mentioned, but no clear correction value was found",
                refined_feedback=refined_feedback,
                query_text=query_text,
                original_feedback=original_feedback,
                current_sheet_status=STATUS_NOT_CORRECTED,
                actions=[],
            )

        actions: List[Action] = []
        prev_sheet = previous_sheet(sheet_names, current_sheet, 1)

        # Current sheet: update Previous Stock column only if it exists later.
        # The action is still recorded; the applier decides if the column exists.
        actions.append(Action(
            metric="previous_stock",
            value=value,
            target_sheet=current_sheet,
            target_column_kind="previous_stock",
            apply_value=True,
            current_sheet_status=STATUS_CORRECTED,
            target_sheet_status=STATUS_CORRECTED,
            reason="Previous stock value should be placed in current month's Previous Stock column if available",
        ))

        if prev_sheet:
            actions.append(Action(
                metric="stock",
                value=value,
                target_sheet=prev_sheet,
                target_column_kind="stock",
                apply_value=True,
                current_sheet_status=f"Corrected in {display_month_name(prev_sheet)}",
                target_sheet_status=STATUS_CORRECTED,
                reason="Previous stock also updates the previous month's stock value",
            ))
            current_status = f"Corrected in {display_month_name(prev_sheet)}"
        else:
            current_status = STATUS_CORRECTED

        return Decision(
            main_status=STATUS_CORRECTED,
            reason="Previous stock correction planned",
            refined_feedback=refined_feedback,
            query_text=query_text,
            original_feedback=original_feedback,
            current_sheet_status=current_status,
            actions=actions,
        )

    # Normal price / purchase / stock correction.
    value = parse_metric_value_from_text(metric, refined_feedback, query_metrics)

    # Only use values already typed in the feedback file columns when the feedback text clearly says
    # the correction was already made in the columns/file. Do NOT use feedback-file values as a
    # general fallback, because many feedback files contain the original queried values there.
    if value is None and looks_like_already_corrected(refined_feedback):
        value = get_feedback_column_value(fb_row, metric, feedback_value_cols)

    if metric == "price" and value is not None:
        value = abs(value)

    target_sheets = target_sheets_from_text(refined_feedback, current_sheet, sheet_names)

    if value is None:
        return Decision(
            main_status=STATUS_NOT_CORRECTED,
            reason="Query was answered, but no clear correction value was found",
            refined_feedback=refined_feedback,
            query_text=query_text,
            original_feedback=original_feedback,
            current_sheet_status=STATUS_NOT_CORRECTED,
            actions=[],
        )

    actions = []
    for sheet in target_sheets:
        if sheet == current_sheet:
            current_status = STATUS_CORRECTED
        else:
            current_status = f"Corrected in {display_month_name(sheet)}"

        actions.append(Action(
            metric=metric,
            value=value,
            target_sheet=sheet,
            target_column_kind=metric,
            apply_value=True,
            current_sheet_status=current_status,
            target_sheet_status=STATUS_CORRECTED,
            reason=f"Clear {metric} correction value found",
        ))

    # For current sheet tracking, if target is previous/explicit month, keep "Corrected in Month".
    current_sheet_status = actions[0].current_sheet_status if actions else STATUS_CORRECTED

    return Decision(
        main_status=STATUS_CORRECTED,
        reason="Correction planned",
        refined_feedback=refined_feedback,
        query_text=query_text,
        original_feedback=original_feedback,
        current_sheet_status=current_sheet_status,
        actions=actions,
    )


# ═══════════════════════════════════════════════════════════════
# DATA APPLY LOGIC
# ═══════════════════════════════════════════════════════════════


def ensure_tracking_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Keep the output tracking columns together at the end of the data file.
    # Order: Queries, Feedback, Correction Status.
    # Queries comes directly from the feedback file's Queries column.
    if QUERIES_COL not in df.columns:
        df[QUERIES_COL] = ""
    if FEEDBACK_COL not in df.columns:
        df[FEEDBACK_COL] = ""
    if STATUS_COL not in df.columns:
        df[STATUS_COL] = ""

    data_columns = [c for c in df.columns if c not in [QUERIES_COL, FEEDBACK_COL, STATUS_COL]]
    tracking_columns = [QUERIES_COL, FEEDBACK_COL, STATUS_COL]
    return df[data_columns + tracking_columns]


def build_row_key_map(df: pd.DataFrame, outlet_col: str, sku_col: str) -> Dict[str, List[int]]:
    keys: Dict[str, List[int]] = {}
    for idx, row in df.iterrows():
        key = make_key(row.get(outlet_col), row.get(sku_col))
        if key:
            keys.setdefault(key, []).append(idx)
    return keys


def choose_price_column_for_row(
    df: pd.DataFrame,
    row_idx: int,
    data_cols: Dict[str, Any],
    query_text: str,
    feedback_text: str,
) -> Tuple[Optional[str], str]:
    price_cols = [c for c in data_cols.get("data_price_cols", []) if c in df.columns]
    if not price_cols:
        return None, "No price column was found in the data file"

    combined = f"{query_text} {feedback_text}".lower()

    # Try explicit text match first.
    for col in price_cols:
        cn = normalize(col)
        if cn and cn in normalize(combined):
            return col, "Price column matched by query/feedback wording"

    # Common semantic hints.
    excl_cols = [c for c in price_cols if re.search(r"excl|excluding|without", c, re.IGNORECASE)]
    incl_cols = [c for c in price_cols if re.search(r"incl|including|with container", c, re.IGNORECASE)]

    if re.search(r"excl|excluding|without\s+container", combined) and excl_cols:
        return excl_cols[0], "Excluding-container price column selected"
    if re.search(r"incl|including|with\s+container", combined) and incl_cols:
        return incl_cols[0], "Including-container price column selected"

    # If only one populated price column on this row, update that one.
    populated = [c for c in price_cols if not is_blank(df.at[row_idx, c])]
    if len(populated) == 1:
        return populated[0], "Only one populated price column found on this row"

    # If only one price column exists, use it.
    if len(price_cols) == 1:
        return price_cols[0], "Only one configured price column exists"

    # If multiple price columns exist but only one is configured first and others blank, use first.
    if not populated:
        return None, "Multiple price columns exist, but none is populated on this row"

    return None, "Multiple price columns are populated and feedback did not specify which one"


def target_column_for_action(
    action: Action,
    df: pd.DataFrame,
    row_idx: int,
    data_cols: Dict[str, Any],
    query_text: str,
    feedback_text: str,
) -> Tuple[Optional[str], str]:
    if action.target_column_kind == "price":
        return choose_price_column_for_row(df, row_idx, data_cols, query_text, feedback_text)

    if action.target_column_kind == "purchase":
        col = data_cols.get("data_purchase")
        return (col, "Purchase column selected") if col in df.columns else (None, "Purchase column missing")

    if action.target_column_kind == "stock":
        col = data_cols.get("data_stock")
        return (col, "Stock column selected") if col in df.columns else (None, "Stock column missing")

    if action.target_column_kind == "previous_stock":
        col = data_cols.get("data_previous_stock")
        return (col, "Previous Stock column selected") if col in df.columns else (None, "Previous Stock column missing")

    return None, "Unknown target column kind"


def style_worksheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#4472C4",
        "font_color": "white",
        "border": 1,
    })
    wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})

    worksheet.freeze_panes(1, 0)
    for col_idx, col_name in enumerate(df.columns):
        worksheet.write(0, col_idx, col_name, header_fmt)
        width = min(max(len(str(col_name)) + 3, 14), 35)
        if col_name in [QUERIES_COL, FEEDBACK_COL, STATUS_COL, "Reason"]:
            width = 35
        worksheet.set_column(col_idx, col_idx, width, wrap_fmt)


# ═══════════════════════════════════════════════════════════════
# MAIN ENGINE
# ═══════════════════════════════════════════════════════════════


def apply_feedback_corrections_ai_agent(
    project_name: str,
    data_path: str,
    feedback_path: str,
    output_dir: str,
) -> Tuple[str, str]:
    start = time.time()

    print("\n" + "=" * 78)
    print(f"FEEDBACK CORRECTION AI AGENT — {project_name}")
    print("=" * 78)

    data_sheets, sheet_names = load_data_workbook(data_path)
    current_sheet = sheet_names[-1]
    current_df = data_sheets[current_sheet]

    feedback_df, feedback_sheet = read_excel_sheet(feedback_path, PREFERRED_FEEDBACK_SHEET)

    print(f"Data current sheet: {current_sheet}")
    print(f"Feedback sheet used: {feedback_sheet}")

    data_cols = resolve_data_columns(current_df, project_name)
    fb_cols = resolve_feedback_columns(feedback_df)
    feedback_value_cols = resolve_feedback_value_columns(feedback_df, data_cols)

    required_data = ["data_outlet", "data_sku"]
    missing_data = [k for k in required_data if not data_cols.get(k)]
    if missing_data:
        raise ValueError(f"Missing required data columns: {missing_data}\nResolved data columns: {data_cols}")

    required_fb = ["fb_outlet", "fb_sku", "fb_feedback"]
    missing_fb = [k for k in required_fb if not fb_cols.get(k)]
    if missing_fb:
        raise ValueError(f"Missing required feedback columns: {missing_fb}\nResolved feedback columns: {fb_cols}")

    print("\nResolved data columns:")
    for k, v in data_cols.items():
        if k != "project_config":
            print(f"  {k}: {v}")

    print("\nResolved feedback columns:")
    for k, v in fb_cols.items():
        print(f"  {k}: {v}")

    print("\nDetected possible feedback value columns:")
    for k, v in feedback_value_cols.items():
        print(f"  {k}: {v}")

    corrected_feedback_df = feedback_df.copy()

    corrected_sheets: Dict[str, pd.DataFrame] = {}
    sheet_key_maps: Dict[str, Dict[str, List[int]]] = {}
    summary_rows: List[Dict[str, Any]] = []

    def get_sheet_df(sheet: str) -> pd.DataFrame:
        if sheet not in corrected_sheets:
            corrected_sheets[sheet] = ensure_tracking_columns(data_sheets[sheet])
        return corrected_sheets[sheet]

    def get_key_map(sheet: str) -> Dict[str, List[int]]:
        if sheet not in sheet_key_maps:
            df = get_sheet_df(sheet)
            sheet_key_maps[sheet] = build_row_key_map(df, data_cols["data_outlet"], data_cols["data_sku"])
        return sheet_key_maps[sheet]

    query_col = fb_cols.get("fb_query")
    feedback_col = fb_cols["fb_feedback"]
    outlet_col = fb_cols["fb_outlet"]
    sku_col = fb_cols["fb_sku"]

    counters = {
        STATUS_CORRECTED: 0,
        STATUS_QUERY_NOT_ANSWERED: 0,
        STATUS_NOT_CORRECTED: 0,
    }

    for fb_idx, fb_row in feedback_df.iterrows():
        outlet = fb_row.get(outlet_col)
        sku = fb_row.get(sku_col)
        key = make_key(outlet, sku)
        original_feedback = clean_text(fb_row.get(feedback_col))
        query_text = clean_text(fb_row.get(query_col)) if query_col else ""

        if not key:
            continue

        if not original_feedback:
            # No feedback was supplied. Track only if there was a query.
            if query_text:
                counters[STATUS_NOT_CORRECTED] += 1
                summary_rows.append({
                    "Feedback Row": fb_idx + 2,
                    "Outlet": outlet,
                    "SKU": sku,
                    "Queries": query_text,
                    "Original Feedback": original_feedback,
                    "Refined Feedback": original_feedback,
                    "Main Status": STATUS_NOT_CORRECTED,
                    "Tracking Status": STATUS_NOT_CORRECTED,
                    "Target Sheet": current_sheet,
                    "Target Column": "",
                    "Old Value": "",
                    "New Value": "",
                    "Reason": "No feedback was supplied for this query",
                })
            continue

        refined_feedback = refine_feedback(query_text, original_feedback)
        corrected_feedback_df.at[fb_idx, feedback_col] = refined_feedback

        decision = build_decision(
            fb_row=fb_row,
            query_text=query_text,
            original_feedback=original_feedback,
            refined_feedback=refined_feedback,
            current_sheet=current_sheet,
            sheet_names=sheet_names,
            feedback_value_cols=feedback_value_cols,
        )

        counters[decision.main_status] += 1

        # Always track the feedback in the current month, even when correction belongs to previous/explicit month.
        current_df_out = get_sheet_df(current_sheet)
        current_km = get_key_map(current_sheet)
        current_rows = current_km.get(key, [])
        for r in current_rows:
            current_df_out.at[r, QUERIES_COL] = append_tracking(current_df_out.at[r, QUERIES_COL], query_text)
            current_df_out.at[r, FEEDBACK_COL] = append_tracking(current_df_out.at[r, FEEDBACK_COL], refined_feedback)
            current_df_out.at[r, STATUS_COL] = append_tracking(current_df_out.at[r, STATUS_COL], decision.current_sheet_status)

        if decision.main_status != STATUS_CORRECTED or not decision.actions:
            # Even when no correction is applied, if the feedback clearly mentions a previous/explicit
            # month, track that feedback in that month too. This makes it auditable in the actual
            # month sheet instead of hiding it only in the current month.
            tracking_sheets = target_sheets_from_text(refined_feedback, current_sheet, sheet_names)
            if not tracking_sheets:
                tracking_sheets = [current_sheet]

            for tracking_sheet in tracking_sheets:
                if tracking_sheet not in data_sheets:
                    continue
                tracking_df = get_sheet_df(tracking_sheet)
                tracking_km = get_key_map(tracking_sheet)
                tracking_rows = tracking_km.get(key, [])
                for r in tracking_rows:
                    tracking_df.at[r, QUERIES_COL] = append_tracking(tracking_df.at[r, QUERIES_COL], query_text)
                    tracking_df.at[r, FEEDBACK_COL] = append_tracking(tracking_df.at[r, FEEDBACK_COL], refined_feedback)
                    tracking_df.at[r, STATUS_COL] = append_tracking(tracking_df.at[r, STATUS_COL], decision.main_status)

            summary_rows.append({
                "Feedback Row": fb_idx + 2,
                "Outlet": outlet,
                "SKU": sku,
                "Queries": query_text,
                "Original Feedback": original_feedback,
                "Refined Feedback": refined_feedback,
                "Main Status": decision.main_status,
                "Tracking Status": decision.main_status,
                "Target Sheet": ", ".join(tracking_sheets),
                "Target Column": "",
                "Old Value": "",
                "New Value": "",
                "Reason": decision.reason,
            })
            continue

        # Apply each planned correction action.
        applied_any = False
        for action in decision.actions:
            target_sheet = action.target_sheet
            if target_sheet not in data_sheets:
                summary_rows.append({
                    "Feedback Row": fb_idx + 2,
                    "Outlet": outlet,
                    "SKU": sku,
                    "Queries": query_text,
                    "Original Feedback": original_feedback,
                    "Refined Feedback": refined_feedback,
                    "Main Status": STATUS_NOT_CORRECTED,
                    "Tracking Status": STATUS_NOT_CORRECTED,
                    "Target Sheet": target_sheet,
                    "Target Column": "",
                    "Old Value": "",
                    "New Value": action.value,
                    "Reason": "Target sheet was not found in the data workbook",
                })
                continue

            target_df = get_sheet_df(target_sheet)
            target_km = get_key_map(target_sheet)
            target_rows = target_km.get(key, [])

            if not target_rows:
                summary_rows.append({
                    "Feedback Row": fb_idx + 2,
                    "Outlet": outlet,
                    "SKU": sku,
                    "Queries": query_text,
                    "Original Feedback": original_feedback,
                    "Refined Feedback": refined_feedback,
                    "Main Status": STATUS_NOT_CORRECTED,
                    "Tracking Status": STATUS_NOT_CORRECTED,
                    "Target Sheet": target_sheet,
                    "Target Column": "",
                    "Old Value": "",
                    "New Value": action.value,
                    "Reason": "No matching Outlet ID + SKU ID row found in target sheet",
                })
                continue

            for r in target_rows:
                target_col, col_reason = target_column_for_action(
                    action=action,
                    df=target_df,
                    row_idx=r,
                    data_cols=data_cols,
                    query_text=query_text,
                    feedback_text=refined_feedback,
                )

                # Always track feedback on target sheet too.
                target_df.at[r, QUERIES_COL] = append_tracking(target_df.at[r, QUERIES_COL], query_text)
                target_df.at[r, FEEDBACK_COL] = append_tracking(target_df.at[r, FEEDBACK_COL], refined_feedback)

                if not target_col:
                    # Previous Stock missing is allowed; do not force into Total Stock in current month.
                    if action.target_column_kind == "previous_stock" and target_sheet == current_sheet:
                        target_df.at[r, STATUS_COL] = append_tracking(target_df.at[r, STATUS_COL], action.current_sheet_status)
                        summary_rows.append({
                            "Feedback Row": fb_idx + 2,
                            "Outlet": outlet,
                            "SKU": sku,
                            "Queries": query_text,
                            "Original Feedback": original_feedback,
                            "Refined Feedback": refined_feedback,
                            "Main Status": STATUS_CORRECTED,
                            "Tracking Status": action.current_sheet_status,
                            "Target Sheet": target_sheet,
                            "Target Column": "",
                            "Old Value": "",
                            "New Value": action.value,
                            "Reason": "Current month Previous Stock column not found; feedback was tracked and previous month correction will still be applied if possible",
                        })
                        continue

                    target_df.at[r, STATUS_COL] = append_tracking(target_df.at[r, STATUS_COL], STATUS_NOT_CORRECTED)
                    summary_rows.append({
                        "Feedback Row": fb_idx + 2,
                        "Outlet": outlet,
                        "SKU": sku,
                        "Queries": query_text,
                        "Original Feedback": original_feedback,
                        "Refined Feedback": refined_feedback,
                        "Main Status": STATUS_NOT_CORRECTED,
                        "Tracking Status": STATUS_NOT_CORRECTED,
                        "Target Sheet": target_sheet,
                        "Target Column": "",
                        "Old Value": "",
                        "New Value": action.value,
                        "Reason": col_reason,
                    })
                    continue

                old_value = target_df.at[r, target_col]
                new_value = abs(action.value) if action.target_column_kind == "price" and action.value is not None else action.value
                new_value = format_value(new_value)

                target_df.at[r, target_col] = new_value
                target_df.at[r, STATUS_COL] = append_tracking(target_df.at[r, STATUS_COL], action.target_sheet_status)
                applied_any = True

                summary_rows.append({
                    "Feedback Row": fb_idx + 2,
                    "Outlet": outlet,
                    "SKU": sku,
                    "Queries": query_text,
                    "Original Feedback": original_feedback,
                    "Refined Feedback": refined_feedback,
                    "Main Status": STATUS_CORRECTED,
                    "Tracking Status": action.target_sheet_status,
                    "Target Sheet": target_sheet,
                    "Target Column": target_col,
                    "Old Value": old_value,
                    "New Value": new_value,
                    "Reason": f"{action.reason}; {col_reason}",
                })

        if not applied_any and decision.main_status == STATUS_CORRECTED:
            # It was understood as correctable, but could not actually apply anywhere.
            # The feedback is already tracked in current sheet from above.
            pass

    os.makedirs(output_dir, exist_ok=True)

    data_base = os.path.splitext(os.path.basename(data_path))[0]
    feedback_base = os.path.splitext(os.path.basename(feedback_path))[0]

    corrected_data_file = os.path.join(output_dir, f"{data_base}-Corrected Data File.xlsx")
    corrected_feedback_file = os.path.join(output_dir, f"{feedback_base}-Corrected Feedback File.xlsx")

    if not corrected_sheets:
        corrected_sheets[current_sheet] = ensure_tracking_columns(data_sheets[current_sheet])

    # Write corrected data workbook.
    with pd.ExcelWriter(corrected_data_file, engine="xlsxwriter") as writer:
        for sheet in sheet_names:
            if sheet in corrected_sheets:
                out_df = corrected_sheets[sheet]
                safe_sheet = str(sheet)[:31]
                out_df.to_excel(writer, sheet_name=safe_sheet, index=False)
                style_worksheet(writer, out_df, safe_sheet)

        summary_df = pd.DataFrame(summary_rows)
        if summary_df.empty:
            summary_df = pd.DataFrame({"Message": ["No feedback corrections processed."]})
        summary_df.to_excel(writer, sheet_name="Correction Summary", index=False)
        style_worksheet(writer, summary_df, "Correction Summary")

    # Write corrected feedback file.
    with pd.ExcelWriter(corrected_feedback_file, engine="xlsxwriter") as writer:
        corrected_feedback_df.to_excel(writer, sheet_name=str(feedback_sheet)[:31], index=False)
        style_worksheet(writer, corrected_feedback_df, str(feedback_sheet)[:31])

        summary_df = pd.DataFrame(summary_rows)
        if summary_df.empty:
            summary_df = pd.DataFrame({"Message": ["No feedback corrections processed."]})
        summary_df.to_excel(writer, sheet_name="Correction Summary", index=False)
        style_worksheet(writer, summary_df, "Correction Summary")

    elapsed = time.time() - start

    print("\n" + "=" * 78)
    print("RUN COMPLETE")
    print("=" * 78)
    print(f"Corrected: {counters[STATUS_CORRECTED]}")
    print(f"Query Not Answered: {counters[STATUS_QUERY_NOT_ANSWERED]}")
    print(f"Not Corrected: {counters[STATUS_NOT_CORRECTED]}")
    print(f"Runtime: {elapsed:.1f}s")
    print("\nOutput files:")
    print(f"  Corrected Data File:     {corrected_data_file}")
    print(f"  Corrected Feedback File: {corrected_feedback_file}")

    return corrected_data_file, corrected_feedback_file


if __name__ == "__main__":
    apply_feedback_corrections_ai_agent(
        PROJECT_NAME,
        DATA_PATH,
        FEEDBACK_PATH,
        OUTPUT_DIR,
    )
