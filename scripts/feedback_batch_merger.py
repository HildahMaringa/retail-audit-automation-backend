"""
Unified Multi-Sheet Feedback Batch Merger + Feedback Refiner
===========================================================

What this script does:
- Reads Batch 1/2/3/etc query workbooks and feedback workbooks.
- Discovers whatever sheets actually exist in the files.
- Processes only sheets that exist in both query and feedback files for a given batch.
- Keeps the old feedback refinement/correction logic.
- Supports multiple corrections inside one feedback response.
- Lets you choose the output format source globally or per sheet:
      query Batch 1, feedback Batch 1, query Batch 2, etc.
- For Nigeria projects only, applies the special Outlet Stock / Outlet Purchases rule:
      * middle feedback column = matched text feedback, kept and moved to the end
      * last feedback column = numeric replacement feedback, used to replace month value
      * blank numeric feedback does not overwrite existing value
      * numeric replacement column is not kept in final output

Dependencies:
    pip install pandas numpy openpyxl xlsxwriter

How to run:
    1. Update CONFIG paths below.
    2. Run:
          python unified_feedback_batch_merger_dynamic.py
"""

import json
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


# ═══════════════════════════════════════════════════════════════
# PROJECT CONFIGS — imported from separate file
# ═══════════════════════════════════════════════════════════════

# Keep your project configuration code in project_configs.py in the same folder
# as this script. Do not paste PROJECT_CONFIGS directly into this file.
from project_configs import PROJECT_CONFIGS, get_stock_thresholds


# ═══════════════════════════════════════════════════════════════
# CONFIG — UPDATE THESE VALUES
# ═══════════════════════════════════════════════════════════════

PROJECT_NAME = "Nigeria-Ville"

# Put all feedback files for the month in this folder.
FEEDBACK_FOLDER = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\Nigeria Ville\Nigeria Ville-Feedbacks\May2026"

# Put all query files for the month in this folder.
QUERY_FOLDER = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\Nigeria Ville\Nigeria Ville-Queries\May2026"

BATCH_NUMBERS = [1, 2, 3]

OUTPUT_DIR = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\Nigeria Ville\Nigeria Ville-Feedbacks\May2026-2"
OUTPUT_FILE_NAME = "Nigeria-Ville-Merged-Feedback-May2026.xlsx"

# The engine discovers sheets dynamically from the actual workbooks.
PROCESS_ALL_AVAILABLE_SHEETS = True

# Optional. Leave empty to process all discovered sheets.
# Example: ONLY_SHEETS = ["Data Queries", "Outlet Stock"]
ONLY_SHEETS: List[str] = []

# Optional. Sheets to ignore.
EXCLUDED_SHEETS = ["Merge Summary", "Summary"]

# Global output format source.
# source_type can be: "query" or "feedback".
FORMAT_SOURCE_DEFAULT = {
    "source_type": "batch",
    "batch": 3,
}

# Per-sheet override.
# Example:
# FORMAT_SOURCE_BY_SHEET = {
#     "Outlet Stock": {"source_type": "feedback", "batch": 1},
#     "Outlet Purchases": {"source_type": "query", "batch": 2},
# }
FORMAT_SOURCE_BY_SHEET: Dict[str, Dict[str, Any]] = {}

BATCH_COL = "Batch"
QUERY_COL = "Queries"
FEEDBACK_COL = "Feedback"
ORIGINAL_FEEDBACK_COL = "Original Feedback"
AI_NOTES_COL = "AI Notes"

# Nigeria-only rule for outlet stock/purchases two feedback columns.
NIGERIA_PROJECTS = {"NG-MRA", "Nigeria-MRA", "Nigeria-Ville"}
NIGERIA_NUMERIC_FEEDBACK_SHEETS = {"outletstock", "outletpurchases"}

# Columns named xx / xx.1 / xx.2 are separator/helper columns in some Nigeria files.
# They should not be used for matching and are removed from final output by default.
DROP_XX_HELPER_COLUMNS = True

AI_CACHE_FILE = "feedback_refine_cache.json"
INTERPRETATION_CACHE_VERSION = "dynamic_multisheet_v1"
USE_AI_AGENT = False
OPENAI_MODEL = "gpt-4o-mini"


# ═══════════════════════════════════════════════════════════════
# BASIC HELPERS
# ═══════════════════════════════════════════════════════════════


def normalize(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def is_blank(value: Any) -> bool:
    return clean_text(value).lower() in ["", "nan", "none", "null"]


def safe_number(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        if pd.isna(value):
            return None
        return float(value)
    text = clean_text(value).replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", text)
    if m:
        try:
            return float(m.group(0))
        except Exception:
            pass
    return number_from_words(value)


def format_value(value: Optional[float]) -> Any:
    if value is None:
        return ""
    try:
        f = float(value)
        if f.is_integer():
            return int(f)
        return f
    except Exception:
        return value


NUMBER_WORDS = {
    "zero": 0, "none": 0, "nil": 0,
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
    "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20,
    "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70,
    "eighty": 80, "ninety": 90,
}


def number_from_words(text: Any) -> Optional[float]:
    t = clean_text(text).lower()
    for word, value in NUMBER_WORDS.items():
        if re.search(rf"\b{re.escape(word)}\b", t):
            return float(value)
    return None


def append_unique(existing: Any, new_text: str, sep: str = ", ") -> str:
    existing_text = clean_text(existing)
    new_text = clean_text(new_text)
    if not new_text:
        return existing_text
    if not existing_text:
        return new_text
    parts = [p.strip() for p in existing_text.split(sep) if p.strip()]
    if new_text in parts or new_text in existing_text:
        return existing_text
    return existing_text + sep + new_text


# ═══════════════════════════════════════════════════════════════
# MONTH HELPERS
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

MONTH_DISPLAY = {
    1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June",
    7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December",
}


def add_months(year: int, month: int, offset: int) -> Tuple[int, int]:
    total = year * 12 + (month - 1) + offset
    return total // 12, total % 12 + 1


def month_year_tokens(year: int, month: int) -> List[str]:
    full = MONTH_DISPLAY[month]
    short = full[:3]
    yy = str(year)[-2:]
    yyyy = str(year)
    return [
        f"{full}{yyyy}", f"{full}{yy}", f"{short}{yyyy}", f"{short}{yy}",
        f"{full} {yyyy}", f"{full} {yy}", f"{short} {yyyy}", f"{short} {yy}",
        f"{full}-{yyyy}", f"{full}-{yy}", f"{short}-{yyyy}", f"{short}-{yy}",
        f"{full}_{yyyy}", f"{full}_{yy}", f"{short}_{yyyy}", f"{short}_{yy}",
        f"({short}{yy})", f"({full}{yy})", full, short,
    ]


def month_year_in_text(text: str) -> Optional[Tuple[int, int]]:
    raw = clean_text(text)
    norm = normalize(raw)

    for name, month in sorted(MONTH_ALIASES.items(), key=lambda x: -len(x[0])):
        for m in re.finditer(rf"{re.escape(name)}\s*[-_/ ]?\s*(20\d{{2}}|\d{{2}})", raw, flags=re.IGNORECASE):
            ytxt = m.group(1)
            year = int(ytxt) if len(ytxt) == 4 else 2000 + int(ytxt)
            return year, month
        for m in re.finditer(rf"{re.escape(name)}(20\d{{2}}|\d{{2}})", norm, flags=re.IGNORECASE):
            ytxt = m.group(1)
            year = int(ytxt) if len(ytxt) == 4 else 2000 + int(ytxt)
            return year, month

    return None


def month_number_from_text(text: str) -> Optional[int]:
    t = clean_text(text).lower()
    for name, num in sorted(MONTH_ALIASES.items(), key=lambda x: -len(x[0])):
        if re.search(rf"\b{re.escape(name)}\b", t):
            return num
    return None


def infer_current_month(*texts: str) -> Optional[Tuple[int, int]]:
    for text in texts:
        ym = month_year_in_text(text)
        if ym:
            return ym
    return None


def explicit_month_from_feedback(feedback_text: str, current_month: Optional[Tuple[int, int]]) -> Optional[Tuple[int, int]]:
    ym = month_year_in_text(feedback_text)
    if ym:
        return ym
    m = month_number_from_text(feedback_text)
    if m and current_month:
        year = current_month[0]
        if current_month[1] == 1 and m == 12:
            year -= 1
        return year, m
    return None


def column_mentions_month(col: str, target_month: Optional[Tuple[int, int]]) -> bool:
    if not target_month:
        return False
    year, month = target_month
    cn = normalize(col)
    for tok in month_year_tokens(year, month):
        if normalize(tok) and normalize(tok) in cn:
            return True
    return False


def column_mentions_any_month(col: str) -> bool:
    return month_year_in_text(col) is not None or month_number_from_text(col) is not None


def is_metric_excluded_col(col: str) -> bool:
    """Columns that must never be treated as value-correction targets."""
    cn = normalize(col)
    if not cn:
        return True
    if re.fullmatch(r"xx(?:\d+)?", cn):
        return True
    excluded_parts = [
        "feedback", "feedbacks", "originalfeedback", "ainotes", "query", "queries",
        "outlier", "outliers", "status", "flag", "remark", "comment", "comments",
        "description", "desc", "name", "channel", "auditor", "fwname",
    ]
    return any(part in cn for part in excluded_parts)


def metric_months_in_columns(columns: List[str], metric: str, hint: str = "", prefer_individual: bool = False) -> List[Tuple[int, int]]:
    """Find all month/year values in columns that look like targets for a metric."""
    months: List[Tuple[int, int]] = []
    for col in columns:
        if is_metric_excluded_col(col):
            continue
        if prefer_individual and "sum" in normalize(col):
            continue
        if metric == "purchase" and not is_purchase_col(col):
            continue
        if metric == "stock" and not is_stock_col(col):
            continue
        if metric == "price" and not is_price_col(col):
            continue
        if hint and not column_base_matches_hint(col, hint):
            continue
        ym = month_year_in_text(col)
        if ym and ym not in months:
            months.append(ym)
    months.sort()
    return months


def latest_metric_month_from_columns(columns: List[str], metric: str, hint: str = "", current_month: Optional[Tuple[int, int]] = None, prefer_individual: bool = False) -> Optional[Tuple[int, int]]:
    """Use current month only if it exists for the target metric; otherwise choose latest available metric month."""
    months = metric_months_in_columns(columns, metric, hint, prefer_individual=prefer_individual)
    if current_month and current_month in months:
        return current_month
    if months:
        return months[-1]
    return current_month


# ═══════════════════════════════════════════════════════════════
# EXCEL DISCOVERY / READING
# ═══════════════════════════════════════════════════════════════


def excel_files_in_folder(folder_path: str, label: str) -> List[Path]:
    folder = Path(folder_path)
    if folder.is_file() and folder.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
        return [folder]
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"{label} path does not exist or is not a folder/file: {folder_path}")

    files: List[Path] = []
    for ext in ["*.xlsx", "*.xlsm", "*.xls"]:
        files.extend(folder.glob(ext))
    files = [f for f in files if not f.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"No Excel files found in {label} folder: {folder_path}")
    return files


def filename_mentions_batch(filename: str, batch_no: int) -> bool:
    n = str(batch_no)
    stem = Path(filename).stem.lower()
    compact = normalize(stem)
    patterns = [
        rf"\bbatch\s*[-_ ]*{n}\b",
        rf"\bbatch{n}\b",
        rf"\bb{n}\b",
        rf"\b{n}(?:st|nd|rd|th)?\s*batch\b",
        rf"(?<!\d){n}(?!\d)",
    ]
    if any(re.search(p, stem, flags=re.IGNORECASE) for p in patterns):
        return True
    compact_patterns = [f"batch{n}", f"b{n}"]
    return any(p in compact for p in compact_patterns)


def find_batch_excel_file(folder_path: str, batch_no: int, label: str) -> str:
    files = excel_files_in_folder(folder_path, label)
    batch_matches = [f for f in files if filename_mentions_batch(f.name, batch_no)]

    if not batch_matches:
        available = "\n  - ".join(f.name for f in files)
        raise FileNotFoundError(
            f"Could not find Batch {batch_no} file in {label} folder: {folder_path}\n"
            f"Available Excel files:\n  - {available}"
        )

    label_norm = normalize(label)
    preferred: List[Path] = []
    for f in batch_matches:
        fn = normalize(f.name)
        if "query" in label_norm and "quer" in fn:
            preferred.append(f)
        elif "feedback" in label_norm and ("feedback" in fn or "fb" in fn):
            preferred.append(f)

    chosen_list = preferred if preferred else batch_matches
    chosen = max(chosen_list, key=lambda x: x.stat().st_mtime)
    return str(chosen)


def get_workbook_sheets(path: str) -> List[str]:
    xl = pd.ExcelFile(path)
    return xl.sheet_names


def find_actual_sheet_name(path: str, sheet_name: str) -> Optional[str]:
    try:
        xl = pd.ExcelFile(path)
        for s in xl.sheet_names:
            if normalize(s) == normalize(sheet_name):
                return s
    except Exception:
        return None
    return None


def workbook_has_sheet(path: str, sheet_name: str) -> bool:
    return find_actual_sheet_name(path, sheet_name) is not None


def discover_available_sheets(query_files: Dict[int, str], feedback_files: Dict[int, str]) -> List[str]:
    seen: List[str] = []
    seen_norm = set()

    for file_map in [query_files, feedback_files]:
        for _, path in file_map.items():
            try:
                for sheet in get_workbook_sheets(path):
                    n = normalize(sheet)
                    if not n:
                        continue
                    if any(n == normalize(x) for x in EXCLUDED_SHEETS):
                        continue
                    if ONLY_SHEETS and not any(n == normalize(x) for x in ONLY_SHEETS):
                        continue
                    if n not in seen_norm:
                        seen.append(sheet)
                        seen_norm.add(n)
            except Exception as e:
                print(f"Could not read sheets from {path}: {e}")

    return seen


def read_sheet(path: str, sheet_name: str) -> Tuple[pd.DataFrame, str]:
    actual_sheet = find_actual_sheet_name(path, sheet_name)
    if actual_sheet is None:
        xl = pd.ExcelFile(path)
        raise ValueError(f"Could not find sheet '{sheet_name}' in {path}. Sheets found: {xl.sheet_names}")

    df = pd.read_excel(path, sheet_name=actual_sheet)
    df.columns = [str(c).strip() for c in df.columns]
    unnamed = sum(1 for c in df.columns if str(c).lower().startswith("unnamed"))
    if unnamed > max(2, len(df.columns) // 2):
        df = pd.read_excel(path, sheet_name=actual_sheet, header=1)
        df.columns = [str(c).strip() for c in df.columns]
    return df, actual_sheet


def build_norm_col_map(df: pd.DataFrame) -> Dict[str, str]:
    return {normalize(c): c for c in df.columns}


def resolve_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    norm_map = build_norm_col_map(df)
    for cand in candidates:
        if not cand:
            continue
        if cand in df.columns:
            return cand
        n = normalize(cand)
        if n in norm_map:
            return norm_map[n]
    for cand in candidates:
        if not cand:
            continue
        cn = normalize(cand)
        for col in df.columns:
            n = normalize(col)
            if cn and (cn in n or n in cn):
                return col
    return None


def align_to_format(source_df: pd.DataFrame, format_cols: List[str]) -> pd.DataFrame:
    norm_map = build_norm_col_map(source_df)
    out = pd.DataFrame(index=source_df.index)
    for fmt_col in format_cols:
        src_col = norm_map.get(normalize(fmt_col))
        out[fmt_col] = source_df[src_col] if src_col is not None else ""
    return out


def remove_excel_line_breaks(value: Any) -> Any:
    """Remove embedded line breaks so Excel does not display wrapped/multiline cells.

    This is different from xlsxwriter's text_wrap option. Even when text_wrap=False,
    Excel can still show multiple lines if the cell text itself contains '\n' or '\r'.
    So we clean the actual cell values before writing the workbook. Numeric/date values
    are left untouched.
    """
    if value is None:
        return value
    try:
        if pd.isna(value):
            return value
    except Exception:
        pass
    if isinstance(value, str):
        cleaned = re.sub(r"[\r\n]+", " ", value)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        return cleaned
    return value


def normalize_output_dataframe_text(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of df with no embedded line breaks in string cells.

    This is applied to every output sheet immediately before writing to Excel.
    It prevents tall wrapped-looking rows like Outlet Details/Product Description
    even when the cell format has text_wrap disabled.
    """
    if df.empty:
        return df.copy()
    out = df.copy()
    object_cols = out.select_dtypes(include=["object"]).columns
    for col in object_cols:
        out[col] = out[col].map(remove_excel_line_breaks)
    return out


# ═══════════════════════════════════════════════════════════════
# DYNAMIC COOLER / POS QUERY CORRECTIONS
# ═══════════════════════════════════════════════════════════════

def is_dynamic_attribute_sheet(sheet_name: str) -> bool:
    """True for flexible attribute sheets such as Cooler Queries and POS Queries.

    Data Queries, Outlet Stock and Outlet Purchases already have strict, working
    correction logic. Do not route them through this dynamic attribute logic.
    """
    sn = normalize(sheet_name)
    if sn in {"dataqueries", "outletstock", "outletpurchases"}:
        return False
    return "cooler" in sn or "pos" in sn


def split_camel_and_words(text: Any) -> List[str]:
    """Split labels such as FridgeStockWithoutCooler into comparable tokens."""
    raw = clean_text(text)
    raw = re.sub(r"([a-z])([A-Z])", r"\1 \2", raw)
    raw = re.sub(r"[^A-Za-z0-9]+", " ", raw)
    return [w.lower() for w in raw.split() if w.strip()]


DYNAMIC_LABEL_STOPWORDS = {
    "number", "numbers", "num", "no", "of", "the", "a", "an", "and", "or",
    "total", "count", "counts", "query", "queries", "current", "previous",
    "month", "sum", "without", "with", "stock", "fridge", "captured", "capture",
    "recorded", "observed", "outlet", "visit", "auditor", "time",
}


def dynamic_content_tokens(text: Any) -> List[str]:
    """Return useful comparable tokens for a dynamic attribute label/feedback phrase."""
    tokens = split_camel_and_words(text)
    cleaned: List[str] = []
    for t in tokens:
        # Keep words that identify the object, drop generic words.
        if t in DYNAMIC_LABEL_STOPWORDS:
            continue
        if t.endswith("s") and len(t) > 3:
            t = t[:-1]
        cleaned.append(t)
    return cleaned


def dynamic_query_parts(query_text: str) -> List[str]:
    """Split query labels while preserving original wording."""
    parts = [p.strip() for p in re.split(r"\s*,\s*|\s*\|\s*|\s*/\s*", clean_text(query_text)) if p.strip()]
    return parts if parts else ([clean_text(query_text)] if clean_text(query_text) else [])


def dynamic_feedback_number_mentions(feedback_text: str) -> List[Tuple[float, str]]:
    """Extract numeric mentions and the nearby wording around each number.

    Examples:
      "2 branded coolers, 1 unbranded cooler" ->
        [(2, "branded coolers"), (1, "unbranded cooler")]
      "There is 1 chest freezer in the outlet..." ->
        [(1, "chest freezer in the outlet")]
    """
    text = fix_spelling(feedback_text)
    if not text:
        return []

    mentions: List[Tuple[float, str]] = []
    pattern = re.compile(r"(?P<num>\d+(?:\.\d+)?)")
    for m in pattern.finditer(text):
        try:
            value = float(m.group("num"))
        except Exception:
            continue

        after = text[m.end():]
        before = text[:m.start()]

        # Take nearby words after the number first because feedback usually says
        # "2 branded coolers" / "1 chest freezer".
        after_words = re.findall(r"[A-Za-z]+", after)[:8]
        before_words = re.findall(r"[A-Za-z]+", before)[-4:]
        phrase = " ".join(after_words if after_words else before_words)
        if not phrase and before_words:
            phrase = " ".join(before_words)
        mentions.append((value, phrase))

    # Word-number fallback for simple phrases like "two branded coolers".
    if not mentions:
        for word, value in NUMBER_WORDS.items():
            for m in re.finditer(rf"\b{re.escape(word)}\b", text, flags=re.IGNORECASE):
                after = text[m.end():]
                after_words = re.findall(r"[A-Za-z]+", after)[:8]
                if after_words:
                    mentions.append((float(value), " ".join(after_words)))
    return mentions




def dynamic_feedback_yes_no_mentions(feedback_text: str) -> List[Tuple[str, str]]:
    """Extract Yes/No POS-style feedback mentions and nearby wording.

    Examples:
      "Yes branded chairs or tables, No branded table runners" ->
        [("Yes", "branded chairs or tables"), ("No", "branded table runners")]
      "Yes shop sign" -> [("Yes", "shop sign")]

    This is mainly for POS Queries, where the correction value is usually Yes/No
    instead of a number. The labels remain generic: we do not hardcode specific
    POS columns; the wording is matched against whatever columns exist in the sheet.
    """
    text = fix_spelling(feedback_text)
    if not text:
        return []

    mentions: List[Tuple[str, str]] = []

    # Split on strong separators so multiple feedback corrections are handled independently.
    parts = [p.strip() for p in re.split(r"\s*,\s*|\s*\|\s*|\n+|\r+", text) if p.strip()]
    if not parts:
        parts = [text]

    for part in parts:
        m = re.search(r"\b(yes|no)\b\s*(?P<label>.+)$", part, flags=re.IGNORECASE)
        if m:
            value = "Yes" if m.group(1).lower() == "yes" else "No"
            label = clean_text(m.group("label"))
            # Remove generic filler after the target label if present.
            label = re.split(r"\b(?:in|at|during|where|which|that|because)\b", label, maxsplit=1, flags=re.IGNORECASE)[0].strip()
            if label:
                mentions.append((value, label))
            continue

        # Also support wording like "branded signage: yes" or "branded signage - no".
        m = re.search(r"(?P<label>.+?)\s*(?:=|:|-|is|was)\s*\b(yes|no)\b\s*$", part, flags=re.IGNORECASE)
        if m:
            value = "Yes" if m.group(2).lower() == "yes" else "No"
            label = clean_text(m.group("label"))
            if label:
                mentions.append((value, label))

    return mentions


def dynamic_value_for_target(row: pd.Series, target_col: str, value: Any) -> Any:
    """Return the value to write for a dynamic correction.

    Numeric cooler-style corrections remain numeric. POS-style Yes/No corrections
    are written as Yes/No unless the existing cell is clearly coded as 1/0.
    """
    if isinstance(value, str) and value.lower() in {"yes", "no"}:
        existing = row.get(target_col, "")
        num = safe_number(existing)
        if num in [0.0, 1.0]:
            return 1 if value.lower() == "yes" else 0
        return "Yes" if value.lower() == "yes" else "No"
    return format_value(value)

def dynamic_column_is_candidate(col: str) -> bool:
    """Columns eligible for dynamic cooler/POS correction."""
    if is_metric_excluded_col(col):
        return False
    cn = normalize(col)
    if not cn:
        return False
    # Avoid identity/descriptive columns.
    if is_outlet_id_like_col(col) or is_sku_id_like_col(col):
        return False
    if any(p in cn for p in ["outlet", "audit", "interview", "enumerator", "fwname", "name", "description", "desc", "channel"]):
        return False
    # Prefer numeric attribute/history columns. Month columns are ideal, but some
    # projects may have current columns without months, so do not require a month.
    return bool(dynamic_content_tokens(col))


def dynamic_match_score(label_or_phrase: str, col: str) -> int:
    """Score how well a query/feedback phrase matches a column name."""
    label_tokens = set(dynamic_content_tokens(label_or_phrase))
    col_tokens = set(dynamic_content_tokens(col))
    if not label_tokens or not col_tokens:
        return 0
    overlap = label_tokens & col_tokens
    score = len(overlap) * 10
    # Strong bonus for exact normalized containment.
    ln = normalize(label_or_phrase)
    cn = normalize(col)
    if ln and (ln in cn or cn in ln):
        score += 80
    # Bonus if all useful label tokens are represented in the column.
    if label_tokens and label_tokens.issubset(col_tokens):
        score += 40
    return score


def latest_dynamic_month_for_columns(columns: List[str]) -> Optional[Tuple[int, int]]:
    months: List[Tuple[int, int]] = []
    for col in columns:
        ym = month_year_in_text(col)
        if ym and ym not in months:
            months.append(ym)
    months.sort()
    return months[-1] if months else None


def choose_dynamic_target_column(
    columns: List[str],
    query_labels: List[str],
    feedback_phrase: str,
    target_month: Optional[Tuple[int, int]],
) -> Optional[str]:
    """Choose the best column for a dynamic cooler/POS feedback value.

    The function is intentionally generic: it does not hardcode branded/unbranded,
    Pepsi coolers, dispensers, chest freezers, signage, etc. It compares the words
    in the Queries cell and the nearby feedback phrase against whatever columns
    exist in the sheet.
    """
    candidate_cols = [c for c in columns if dynamic_column_is_candidate(c)]
    if not candidate_cols:
        return None

    scored: List[Tuple[int, int, str]] = []
    for col in candidate_cols:
        score = 0
        # Match feedback wording first. This handles FridgeStockWithoutCooler where
        # the feedback says chest freezer / unbranded cooler / branded cooler.
        if feedback_phrase:
            score = max(score, dynamic_match_score(feedback_phrase, col) + 15)

        # Also respect the actual query labels, especially when the query contains
        # multiple requested attributes.
        for label in query_labels:
            score = max(score, dynamic_match_score(label, col))

        if target_month and column_mentions_month(col, target_month):
            score += 25
        elif column_mentions_any_month(col):
            # Older month columns are still candidates, but less preferred.
            score += 5

        if score > 0:
            # Stable column order tie-breaker.
            scored.append((score, -columns.index(col), col))

    if not scored:
        return None

    scored.sort(key=lambda x: (-x[0], -x[1]))
    return scored[0][2]


def apply_dynamic_attribute_corrections(
    row: pd.Series,
    sheet_name: str,
    query_text: str,
    feedback_text: str,
    current_month: Optional[Tuple[int, int]] = None,
) -> Tuple[pd.Series, str, int]:
    """Apply flexible Cooler/POS corrections using query labels + feedback text.

    Only used for Cooler Queries / POS Queries. It leaves Data Queries, Outlet
    Stock and Outlet Purchases untouched.

    Cooler-style feedback is usually numeric, for example:
        "2 branded coolers, 1 unbranded cooler"

    POS-style feedback is usually Yes/No, for example:
        "Yes branded chairs or tables, No branded table runners"

    Both are handled generically by comparing the query/feedback wording against
    the actual columns in the sheet. No fixed list of POS/cooler names is used.
    """
    if not is_dynamic_attribute_sheet(sheet_name):
        return row, "", 0

    columns = list(row.index)
    query_labels = dynamic_query_parts(query_text)
    target_month = explicit_month_from_feedback(feedback_text, current_month)
    if target_month is None:
        target_month = latest_dynamic_month_for_columns([c for c in columns if dynamic_column_is_candidate(c)]) or current_month

    # Numeric corrections for Cooler-style sheets, plus Yes/No corrections for POS-style sheets.
    feedback_mentions: List[Tuple[Any, str]] = []
    for value, phrase in dynamic_feedback_number_mentions(feedback_text):
        feedback_mentions.append((value, phrase))
    for value, phrase in dynamic_feedback_yes_no_mentions(feedback_text):
        feedback_mentions.append((value, phrase))

    if not feedback_mentions:
        return row, "No dynamic Cooler/POS correction value found.", 0

    notes: List[str] = []
    applied = 0
    used_cols: set = set()

    for value, phrase in feedback_mentions:
        target_col = choose_dynamic_target_column(columns, query_labels, phrase, target_month)
        if not target_col or target_col in used_cols:
            continue
        write_value = dynamic_value_for_target(row, target_col, value)
        row[target_col] = write_value
        used_cols.add(target_col)
        applied += 1
        month_note = f" for {MONTH_DISPLAY[target_month[1]]} {target_month[0]}" if target_month else ""
        notes.append(f"Dynamic {sheet_name} correction: updated '{target_col}'{month_note} to {write_value}")

    if applied == 0:
        return row, "Dynamic Cooler/POS feedback had values, but no matching output column was found.", 0
    return row, "; ".join(notes) + ".", applied


# ═══════════════════════════════════════════════════════════════
# COLUMN DETECTION / MATCHING
# ═══════════════════════════════════════════════════════════════


def feedback_like_columns(df: pd.DataFrame) -> List[str]:
    out = []
    for c in df.columns:
        n = normalize(c)
        if n in ["feedback", "feedbacks", "fieldfeedback", "response", "responses", "comment", "comments", "correction"]:
            out.append(c)
        elif n.startswith("feedback") or n.startswith("feedbacks"):
            out.append(c)
    return out


def detect_feedback_columns_for_sheet(df: pd.DataFrame, sheet_name: str) -> Dict[str, Optional[str]]:
    fb_cols = feedback_like_columns(df)
    if not fb_cols:
        return {"text_feedback": None, "numeric_feedback": None}

    is_nigeria_numeric_sheet = (
        PROJECT_NAME in NIGERIA_PROJECTS and normalize(sheet_name) in NIGERIA_NUMERIC_FEEDBACK_SHEETS
    )

    if is_nigeria_numeric_sheet and len(fb_cols) >= 2:
        return {
            "text_feedback": fb_cols[0],      # middle matched feedback
            "numeric_feedback": fb_cols[-1],  # last numeric replacement feedback
        }

    return {"text_feedback": fb_cols[0], "numeric_feedback": None}


def value_format_signature(value: Any) -> str:
    """Return a simple signature for checking if an ID column has a consistent format."""
    text = clean_text(value)
    if not text:
        return "blank"
    compact = normalize(text)
    if re.fullmatch(r"\d+", text):
        return f"digits:{len(text)}"
    if re.fullmatch(r"[A-Za-z]+\d+", text):
        letters = re.match(r"[A-Za-z]+", text).group(0)
        digits = re.search(r"\d+$", text).group(0)
        return f"letters_digits:{len(letters)}:{len(digits)}"
    if re.fullmatch(r"[A-Za-z0-9_-]+", text):
        return f"alnum:{len(compact)}"
    if " " in text:
        return "text_with_spaces"
    return "other"


def column_format_consistency(series: pd.Series, mask: Optional[pd.Series] = None) -> float:
    """How consistently formatted the nonblank values are, from 0 to 1."""
    if mask is not None:
        series = series[mask]
    values = [v for v in series.tolist() if not is_blank(v)]
    if not values:
        return 0.0
    counts: Dict[str, int] = {}
    for v in values:
        sig = value_format_signature(v)
        counts[sig] = counts.get(sig, 0) + 1
    return max(counts.values()) / len(values)


def bad_identifier_column_name(col: str, role: str) -> bool:
    """Exclude descriptive columns that can accidentally match loose names like SKU/Outlet."""
    cn = normalize(col)
    if re.fullmatch(r"xx(?:\d+)?", cn):
        return True
    bad_parts = ["name", "description", "desc", "channel", "auditor", "fwname", "interviewer"]
    if role in {"outlet", "sku"} and any(p in cn for p in bad_parts):
        return True
    return False


def resolve_col_best_nonblank(
    df: pd.DataFrame,
    candidates: List[str],
    allow_contains: bool = True,
    role: str = "generic",
    scope_mask: Optional[pd.Series] = None,
) -> Optional[str]:
    """Resolve a matching column by choosing the most reliable populated ID column.

    This is deliberately NOT hardcoded to outletid.1 or any one project name.
    The rule is:
      - among possible ID columns, choose the one that is populated across the
        individual/detail rows being matched;
      - avoid visually merged display columns that are blank on repeated SKU rows;
      - avoid helper/separator columns such as xx;
      - avoid descriptive columns such as outlet name or SKU description;
      - prefer values with a consistent ID-like format.

    This matters for Nigeria Outlet Stock / Outlet Purchases, where the first
    outlet column can be a visually merged outlet display column while another
    outlet column is populated on every SKU row. The populated one is the one to
    use for matching.
    """
    wanted = [normalize(c) for c in candidates if c]
    if not wanted:
        return None

    if scope_mask is not None:
        try:
            scope_mask = scope_mask.reindex(df.index).fillna(False)
        except Exception:
            scope_mask = None

    scored: List[Tuple[int, int, float, int, int, str]] = []
    for position, col in enumerate(df.columns):
        cn = normalize(col)
        if not cn or bad_identifier_column_name(col, role):
            continue

        exact_rank = None
        for cand in wanted:
            if cn == cand:
                exact_rank = 4
                break
            # Duplicate columns read by pandas usually become outletid.1 / sku_id.1.
            if re.fullmatch(rf"{re.escape(cand)}\d+", cn) or re.fullmatch(rf"{re.escape(cand)}\.\d+", clean_text(col).lower()):
                exact_rank = 4
                break
            if allow_contains and cand and (cand in cn or cn in cand):
                exact_rank = 2
                break

        if exact_rank is None:
            continue

        series = df[col]
        if scope_mask is not None and int(scope_mask.sum()) > 0:
            scope_series = series[scope_mask]
        else:
            scope_series = series

        filled_in_scope = int(scope_series.apply(lambda x: not is_blank(x)).sum())
        filled_total = int(series.apply(lambda x: not is_blank(x)).sum())
        consistency = column_format_consistency(series, scope_mask)

        # Prefer more filled values in the rows that matter, then total filled,
        # then format consistency, then stronger name match. Later pandas duplicate
        # columns do not win by name; they win only if they are actually populated.
        scored.append((filled_in_scope, filled_total, consistency, exact_rank, -position, col))

    if not scored:
        return resolve_col(df, candidates)

    scored.sort(key=lambda x: (-x[0], -x[1], -x[2], -x[3], -x[4]))
    return scored[0][5]


def nonblank_mask_for_column(df: pd.DataFrame, col: Optional[str]) -> Optional[pd.Series]:
    if not col or col not in df.columns:
        return None
    return df[col].apply(lambda x: not is_blank(x))


def detect_identity_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    config = PROJECT_CONFIGS.get(PROJECT_NAME, {})

    outlet_candidates = [
        config.get("outlet_id"),
        "Outlet ID", "OutletID", "Outlet Id", "Outlet Number", "Out Number", "OUTNUMBER",
        "outletid", "outlet_id", "wh_outletid", "projectOutletid", "Outlet No", "OutletNo",
    ]
    sku_candidates = [
        config.get("sku_id"),
        "SKU ID", "Sku ID", "SKU_ID", "SKU", "Sku", "Prod Code", "Prodcode", "Product Code",
        "product code", "wh_skuid", "sku_id",
    ]

    # Use the column populated across detail rows. Do not use xx/helper columns.
    # First pick SKU; then use SKU-filled rows as the scope for choosing outlet ID.
    sku = resolve_col_best_nonblank(df, sku_candidates, role="sku")
    detail_scope = nonblank_mask_for_column(df, sku)
    outlet = resolve_col_best_nonblank(df, outlet_candidates, role="outlet", scope_mask=detail_scope)

    query = resolve_col(df, [
        "Queries", "Query", "Question", "Questions", "Issue", "Title", "Data Query", "Data Queries",
    ])
    return {"outlet": outlet, "sku": sku, "query": query}

def row_key(row: pd.Series, cols: Dict[str, Optional[str]], include_query: bool = True) -> str:
    outlet = clean_text(row.get(cols.get("outlet"), "")) if cols.get("outlet") else ""
    sku = clean_text(row.get(cols.get("sku"), "")) if cols.get("sku") else ""
    query = clean_text(row.get(cols.get("query"), "")) if include_query and cols.get("query") else ""
    return normalize(f"{outlet}|{sku}|{query}")


def fallback_key(row: pd.Series, cols: Dict[str, Optional[str]]) -> str:
    outlet = clean_text(row.get(cols.get("outlet"), "")) if cols.get("outlet") else ""
    sku = clean_text(row.get(cols.get("sku"), "")) if cols.get("sku") else ""
    if sku:
        return normalize(f"{outlet}|{sku}")
    return normalize(f"{outlet}")


def build_feedback_index(feedback_df: pd.DataFrame, fb_cols: Dict[str, Optional[str]]) -> Tuple[Dict[str, List[int]], Dict[str, List[int]]]:
    exact: Dict[str, List[int]] = {}
    loose: Dict[str, List[int]] = {}
    for idx, row in feedback_df.iterrows():
        k = row_key(row, fb_cols, include_query=True)
        lk = fallback_key(row, fb_cols)
        if k:
            exact.setdefault(k, []).append(idx)
        if lk:
            loose.setdefault(lk, []).append(idx)
    return exact, loose


# ═══════════════════════════════════════════════════════════════
# FEEDBACK REFINEMENT / INTERPRETATION
# ═══════════════════════════════════════════════════════════════

SPELL_FIXES = [
    (r"\bam\s+error\b", "an error"),
    (r"\ba\s+error\b", "an error"),
    (r"\bauditor\s+error\b", "auditor error"),
    (r"\bcorrectly\s+counted\b", "correctly counted"),
    (r"\bunit\b", "unit"),
    (r"(?<=\d)\s+un\b", " units"),
    (r"\bunts\b", "units"),
    (r"\bunites\b", "units"),
    (r"\buniits\b", "units"),
    (r"\bbougt\b", "bought"),
    (r"\bboight\b", "bought"),
    (r"\burchasd\b", "purchased"),
    (r"\burceased\b", "purchased"),
    (r"\burchased\b", "purchased"),
    (r"\bpurchse\b", "purchase"),
    (r"\bpurchsed\b", "purchased"),
    (r"\brespodent\b", "respondent"),
    (r"\brespondant\b", "respondent"),
    (r"\bprevioux\b", "previous"),
    (r"\bprevios\b", "previous"),
    (r"\bprevous\b", "previous"),
    (r"\bpreviuos\b", "previous"),
    (r"\bpreviouss\b", "previous"),
    (r"\bprevstock\b", "previous stock"),
    (r"\bprevstok\b", "previous stock"),
    (r"\bprevstpck\b", "previous stock"),
    (r"\bprevious\s+stpck\b", "previous stock"),
    (r"\bprevious\s+stok\b", "previous stock"),
    (r"\bpreviousstpck\b", "previous stock"),
    (r"\bpreviousstock\b", "previous stock"),
    (r"\bprepstock\b", "previous stock"),
    (r"\bprestock\b", "previous stock"),
    (r"\bprep\s+stock\b", "previous stock"),
    (r"\bpre\s+stock\b", "previous stock"),
    (r"\bprev\s+tock\b", "previous stock"),
    (r"\bstpck\b", "stock"),
    (r"\bstok\b", "stock"),
    (r"\btock\b", "stock"),
    (r"\batock\b", "stock"),
    (r"\bztock\b", "stock"),
    (r"\bstockwd\b", "stocked"),
    (r"\bstockd\b", "stocked"),
    (r"\bcompletly\b", "completely"),
    (r"\bafew\b", "a few"),
    (r"\balot\b", "a lot"),
    (r"\bhad ran\b", "had run"),
    (r"\bran low\b", "run low"),
    (r"!\s*(\d)", r"1\1"),
]


def fix_spelling(text: str) -> str:
    fixed = clean_text(text)
    for pattern, replacement in SPELL_FIXES:
        fixed = re.sub(pattern, replacement, fixed, flags=re.IGNORECASE)
    fixed = re.sub(r"\s+", " ", fixed).strip()
    return fixed


def sentence_case(text: str) -> str:
    text = clean_text(text)
    if not text:
        return text
    text = text[0].upper() + text[1:]
    if text[-1] not in ".!?":
        text += "."
    return text


def infer_query_metric(query_text: str) -> str:
    q = clean_text(query_text).lower()
    if re.search(r"negative\s+sales|neg\s*sales", q):
        return "negative_sales"
    if re.search(r"selling|price|capture\s+price|retail\s+price|buying\s+price", q):
        return "price"
    if re.search(r"purchase|purchased|stock increase|bought|buy", q):
        return "purchase"
    if re.search(r"previous\s+stock|prev\s+stock|opening\s+stock", q):
        return "previous_stock"
    if re.search(r"stock|count|available|unit count|total stock|current stock|front stock|back stock|chest freezer|cooler bank", q):
        return "stock"
    if re.search(r"cooler|freezer", q):
        return "cooler"
    if re.search(r"signage|chair|table|runner|shop sign|pos", q):
        return "pos"
    return "other"


def extract_first_positive_number(text: str) -> Optional[float]:
    cleaned = clean_text(text).replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)", cleaned)
    if m:
        try:
            return float(m.group(1))
        except Exception:
            pass
    return number_from_words(text)


def word_or_digit_to_int(value: Any) -> Optional[int]:
    if value is None:
        return None
    s = clean_text(value).lower()
    if re.fullmatch(r"\d+", s):
        return int(s)
    n = number_from_words(s)
    if n is not None:
        return int(n)
    return None


def extract_purchase_period(text: str) -> Dict[str, Optional[int]]:
    fixed = fix_spelling(text).lower().replace(",", "")
    period: Dict[str, Optional[int]] = {"week": None, "day": None}
    week_word_pattern = "|".join(NUMBER_WORDS.keys())
    week_patterns = [
        rf"\bweek\s*({week_word_pattern}|\d+)\b",
        rf"\bwk\s*({week_word_pattern}|\d+)\b",
    ]
    for pattern in week_patterns:
        m = re.search(pattern, fixed, flags=re.IGNORECASE)
        if m:
            period["week"] = word_or_digit_to_int(m.group(1))
            break

    day_patterns = [
        r"\bon\s+(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)?\b",
        r"\bon\s+(\d{1,2})\s*/\s*\d{1,2}\b",
        r"\bday\s+(\d{1,2})\b",
        r"\b(\d{1,2})(?:st|nd|rd|th)\s+(?:of\s+)?(?:this\s+month|may|april|march|january|february|june|july|august|september|october|november|december)\b",
    ]
    for pattern in day_patterns:
        m = re.search(pattern, fixed, flags=re.IGNORECASE)
        if m:
            try:
                day = int(m.group(1))
                if 1 <= day <= 31:
                    period["day"] = day
                    break
            except Exception:
                pass
    return period


def extract_purchase_number(text: str) -> Optional[float]:
    fixed = fix_spelling(text).lower().replace(",", "")

    before_patterns = [
        r"\b(\d+(?:\.\d+)?)\s*(?:more\s+)?(?:pieces?|units?|buckets?|sachets?|packets?)?\s*(?:were\s+)?\b(?:purchased|bought)\b",
        r"\b(\d+(?:\.\d+)?)\s*(?:pieces?|units?|buckets?|sachets?|packets?)\s+(?:were\s+)?(?:purchased|bought)\b",
    ]
    for pattern in before_patterns:
        m = re.search(pattern, fixed, flags=re.IGNORECASE)
        if m:
            return float(m.group(1))

    after_patterns = [
        r"\b(?:purchase|purchases|purchased|bought|buy)\w*\b\s*(?:of|was|were|made|is|=|:)?\s*(\d+(?:\.\d+)?)\s*(?:pieces?|units?|buckets?|sachets?|packets?)?",
        r"\b(?:there\s+was\s+a\s+)?purchase\s+of\s+(\d+(?:\.\d+)?)",
        r"\bpurchase\s+(?:was|were)\s+(\d+(?:\.\d+)?)",
        r"\b\w+\s+purchases\s+were\s+(\d+(?:\.\d+)?)",
        r"\bmissing\s+purchase\s+(?:stock\s+)?of\s+(\d+(?:\.\d+)?)",
        r"\btyping\s+error\s+(\d+(?:\.\d+)?)\s+missing\s+purchase\b",
    ]
    for pattern in after_patterns:
        m = re.search(pattern, fixed, flags=re.IGNORECASE)
        if m:
            return float(m.group(1))

    word_value = number_from_words(fixed)
    if word_value is not None and re.search(r"\b(?:purchase|purchases|purchased|bought|buy)\w*\b", fixed):
        return word_value
    return None


def is_bare_number_feedback(text: str) -> bool:
    t = clean_text(text).lower().strip().strip(" .,:;|-/")
    if not t:
        return False
    if re.fullmatch(r"\d+(?:\.\d+)?", t):
        return True
    return t in NUMBER_WORDS


def is_this_was_value_feedback(text: str) -> bool:
    t = fix_spelling(text).lower()
    return bool(re.search(r"\b(this|it)\s+(was|is)\s+(\d+(?:\.\d+)?|" + "|".join(NUMBER_WORDS.keys()) + r")\b", t))


def mentions_previous_stock(text: str) -> bool:
    t = fix_spelling(text).lower()
    return bool(re.search(
        r"\bprevious\s+stock\b|\bprev\s+stock\b|\bprevstock\b|\bprep\s+stock\b|\bpre\s+stock\b|"
        r"\bopening\s+stock\b|\blast\s+month\s+stock\b",
        t,
    ))


def stock_component_kind_from_text(text: str) -> Optional[str]:
    t = normalize(fix_spelling(text))
    patterns = {
        "front": ["frontstock", "frontstalk", "frontunitstock", "frontunitstalk", "frontunitcount", "frontcount", "frontunit", "frontunits", "ambientfrontcount"],
        "back": ["backstock", "backstalk", "backunitstock", "backunitstalk", "backunitcount", "backcount", "backunit", "backunits"],
        "chest_freezer": ["chestfreezer", "chestfreezers", "chestfreezerstock", "chestfreezerstalk", "freezerstock", "freezerstalk"],
        "cooler_bank": ["coolerbank", "coolerbanks", "coolerbankstock", "coolerbankstalk", "coolerstock", "coolerstalk", "coolerbankscount"],
    }
    for kind, pats in patterns.items():
        if any(p in t for p in pats):
            return kind
    return None


def split_feedback_into_parts(text: str) -> List[str]:
    fixed = fix_spelling(text)
    if not fixed:
        return []

    # Strong separators only. We avoid aggressive trimming so no information is lost.
    parts = [p.strip() for p in re.split(r"\s*\|\s*|\n+|\r+", fixed) if p.strip()]
    if len(parts) > 1:
        return parts
    return [fixed]


def light_grammar_cleanup(text: str) -> str:
    """Small grammar cleanup for the visible Feedback column without removing meaning."""
    t = fix_spelling(text)
    if not t:
        return ""

    # Normalize spacing around punctuation.
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"\s+([,.;:!?])", r"\1", t)
    t = re.sub(r"([,.;:!?])(?=\S)", r"\1 ", t)

    # Common auditor feedback phrases.
    t = re.sub(r"\b(am|a)\s+error\s+from\s+the\s+auditor\b", "An error from the auditor", t, flags=re.IGNORECASE)
    t = re.sub(r"\b(it\s+was\s+)?an\s+auditor\s+error\b", "An auditor error", t, flags=re.IGNORECASE)
    t = re.sub(r"\bstock\s+counted\s+was\b", "stock counted was", t, flags=re.IGNORECASE)
    t = re.sub(r"\b(\d+(?:\.\d+)?)\s+stock\s+correctly\s+counted\b", r"\1 stock was correctly counted", t, flags=re.IGNORECASE)
    t = re.sub(r"\b(\d+(?:\.\d+)?)\s+stocks\s+correctly\s+counted\b", r"\1 stock was correctly counted", t, flags=re.IGNORECASE)

    # Capitalize first letter only; preserve the rest of the auditor wording as much as possible.
    if t:
        t = t[0].upper() + t[1:]
    if t and t[-1] not in ".!?":
        t += "."
    return t


def configured_metric_hint_from_query(query_text: str, metric: str) -> str:
    """Return the exact configured column label mentioned in the query, where possible.

    This is important for Nigeria MRA where a price query can mean either
    'Capture Price' or 'Capture Price Excl Container'. We must update the exact
    metric named in the query and its matching history column.
    """
    qn = normalize(query_text)
    config = PROJECT_CONFIGS.get(PROJECT_NAME, {})

    if metric == "price":
        price_cols = list(config.get("price_cols", []) or [])
        # Longest first so Capture Price Excl Container wins over Capture Price.
        for col in sorted(price_cols, key=lambda x: len(normalize(x)), reverse=True):
            if normalize(col) and normalize(col) in qn:
                return col
        return price_cols[0] if len(price_cols) == 1 else "price"

    if metric in ["stock", "previous_stock"]:
        stock_col = clean_text(config.get("stock_col")) or "Total Stock"
        if normalize(stock_col) and normalize(stock_col) in qn:
            return stock_col
        if "totalstock" in qn:
            return "Total Stock"
        return stock_col

    if metric == "purchase":
        purchase_col = clean_text(config.get("purchase_col")) or "Purchases"
        if normalize(purchase_col) and normalize(purchase_col) in qn:
            return purchase_col
        if "totalpurchase" in qn:
            return "Total Purchase"
        if "purchase" in qn or "purchases" in qn:
            return purchase_col
        return purchase_col

    return ""


def refine_feedback_display(query_text: str, feedback_text: str) -> str:
    """Return the feedback text for the output Feedback column.

    Refining here means: spelling, spacing, capitalization, punctuation, and light
    grammar cleanup. It must NOT shorten the auditor's explanation or remove a
    second correction from the same feedback.

    Only very bare values are expanded using the query context.
    """
    fixed = fix_spelling(feedback_text)
    if not fixed:
        return ""

    q_metric = infer_query_metric(query_text)
    value = extract_first_positive_number(fixed)
    t = clean_text(fixed).strip()
    t_lower = t.lower().strip(" .,:;|-/")

    bare = is_bare_number_feedback(t) or is_this_was_value_feedback(t)
    money_bare = bool(re.fullmatch(r"(?:ugx|kes|ksh|ngn|naira|n|#)?\s*\d+(?:\.\d+)?", t_lower, flags=re.IGNORECASE))

    if value is not None and (bare or money_bare):
        if q_metric == "price":
            hint = configured_metric_hint_from_query(query_text, "price")
            label = hint if hint and hint != "price" else "selling price"
            return f"The {label} was {format_value(value)}."
        if q_metric == "purchase":
            hint = configured_metric_hint_from_query(query_text, "purchase") or "purchase"
            return f"The {hint} was {format_value(value)} units."
        if q_metric == "previous_stock":
            return f"The previous stock was {format_value(value)} units."
        if q_metric in ["stock", "negative_sales"]:
            hint = configured_metric_hint_from_query(query_text, "stock") or "stock"
            return f"The {hint} was {format_value(value)} units."

    return light_grammar_cleanup(t)


def feedback_explicit_metrics(feedback_text: str) -> set:
    """Metrics explicitly mentioned by the feedback text.

    This is used to prevent contradictory corrections. Example:
    - Query asks for Purchases, but feedback says Selling Price was 500.
      The feedback explicitly mentions price, so we must NOT use 500 to update
      Purchases just because the query requested purchases.

    Bare feedback such as "12" or "this was 12" returns an empty set, which means
    the query context may safely decide the metric.
    """
    fixed = fix_spelling(feedback_text)
    f = fixed.lower()
    explicit = set()

    if not f:
        return explicit

    if mentions_previous_stock(fixed):
        explicit.add("previous_stock")

    if stock_component_kind_from_text(fixed):
        explicit.add("stock_component")

    if re.search(r"\b(price|selling\s+price|capture\s+price|retail\s+price|buying\s+price|ugx|shs|ksh|kes|naira|ngn)\b", f):
        explicit.add("price")

    if re.search(r"\b(purchase|purchases|purchased|bought|buy|stock\s+increase)\b", f):
        explicit.add("purchase")

    # Generic stock wording. Avoid marking chest freezer/cooler feedback as stock
    # unless it is clearly a stock/count statement.
    if re.search(r"\b(total\s+stock|current\s+stock|stock\s+counted|stock\s+was|stock\s+is|stock|unit\s+count|counted|available)\b", f):
        explicit.add("stock")

    if re.search(r"\b(cooler|coolers|freezer|freezers|dispenser|dispensers)\b", f):
        explicit.add("cooler")

    if re.search(r"\b(signage|chair|chairs|table|tables|runner|runners|shop\s+sign|pos)\b", f):
        explicit.add("pos")

    return explicit


def has_metric_contradiction(feedback_text: str, requested_metrics: set) -> bool:
    """True when feedback explicitly talks only about metrics not requested.

    If feedback has no explicit metric, it is not a contradiction because values
    like "12" rely on the query context. If there is at least one explicit metric
    that overlaps the query, it is not a contradiction.
    """
    explicit = {canonical_metric(m) for m in feedback_explicit_metrics(feedback_text)}
    requested = {canonical_metric(m) for m in requested_metrics}
    if not explicit:
        return False
    return explicit.isdisjoint(requested)

def rule_interpret_feedback_multi(query_text: str, feedback_text: str) -> List[Dict[str, Any]]:
    """Return one or more correction decisions from a feedback part.

    Important: this does not trim meaning. If a single feedback mentions both purchase and price,
    both decisions can be returned and applied.
    """
    original = clean_text(feedback_text)
    fixed = fix_spelling(original)
    q_metric = infer_query_metric(query_text)
    f_lower = fixed.lower()
    display_feedback = refine_feedback_display(query_text, fixed)
    explicit_metrics = {canonical_metric(m) for m in feedback_explicit_metrics(fixed)}
    bare_value_feedback = is_bare_number_feedback(fixed) or is_this_was_value_feedback(fixed)
    decisions: List[Dict[str, Any]] = []

    def add(metric: str, value: Optional[float], refined: str, hint: str = "", notes: str = "Rule fallback used.", confidence: str = "medium"):
        decisions.append({
            "refined_feedback": refined,
            "answered_metric": metric,
            "correction_value": value,
            "target_column_hint": hint,
            "notes": notes,
            "confidence": confidence,
        })

    value = extract_first_positive_number(fixed)

    if mentions_previous_stock(fixed):
        add(
            "previous_stock",
            value,
            display_feedback,
            "previous_stock",
            "Previous stock mentioned. Back stock is not previous stock.",
            "high" if value is not None else "low",
        )

    component_kind = stock_component_kind_from_text(fixed)
    if component_kind:
        readable = {
            "front": "front stock unit count",
            "back": "back stock unit count",
            "chest_freezer": "chest freezer stock",
            "cooler_bank": "cooler bank stock",
        }.get(component_kind, "stock component")
        add(
            "stock_component",
            value,
            display_feedback,
            component_kind,
            "Current-month stock component. Only the named component should be corrected.",
            "high" if value is not None else "low",
        )

    purchase_value = extract_purchase_number(fixed)
    # If the query asks for purchases and the feedback is just a number (or "this was 12"),
    # still treat it as a purchase correction. However, if the feedback explicitly
    # says price/stock/etc., do NOT use that number to change purchases.
    if (
        purchase_value is None
        and q_metric == "purchase"
        and value is not None
        and (bare_value_feedback or not explicit_metrics or "purchase" in explicit_metrics)
        and not (explicit_metrics - {"purchase"})
    ):
        purchase_value = value
    if purchase_value is not None:
        period = extract_purchase_period(fixed)
        hint = "purchase"
        period_text = ""
        if period.get("day"):
            hint += f"|day={period['day']}"
            period_text = f" on day {period['day']}"
        elif period.get("week"):
            hint += f"|week={period['week']}"
            period_text = f" in week {period['week']}"
        add(
            "purchase",
            purchase_value,
            display_feedback,
            (configured_metric_hint_from_query(query_text, "purchase") + (hint.replace("purchase", "", 1) if hint.startswith("purchase") else "|" + hint)).strip("|"),
            "Purchase correction detected.",
            "high",
        )

    # Price logic: handle explicit price language, money symbols, or bare number when query is price.
    # Do not turn an explicitly stock/purchase feedback value into a price correction.
    price_value = value
    if price_value is not None:
        price_cols_text = " ".join(PROJECT_CONFIGS.get(PROJECT_NAME, {}).get("price_cols", []))
        feedback_mentions_price = bool(
            re.search(r"\bprice\b|selling|capture|retail|buying|ugx|shs|ksh|kes|naira|ngn", f_lower)
            or (price_cols_text and any(normalize(p) in normalize(fixed) for p in price_cols_text.split()))
        )
        price_allowed_from_query_context = (
            q_metric == "price"
            and (bare_value_feedback or not explicit_metrics or "price" in explicit_metrics)
            and not (explicit_metrics - {"price"})
        )
        if feedback_mentions_price or price_allowed_from_query_context:
            add(
                "price",
                price_value,
                display_feedback,
                configured_metric_hint_from_query(query_text, "price"),
                "Price correction detected.",
                "high" if q_metric == "price" else "medium",
            )

    # Generic stock should only be added if no stock component / previous-stock decision already covers it.
    # Do not turn an explicitly purchase/price feedback value into a stock correction.
    if value is not None:
        feedback_mentions_stock = bool(re.search(r"\bstock\b|counted|available|unit count", f_lower))
        already_stock_specific = any(d["answered_metric"] in ["previous_stock", "stock_component"] for d in decisions)
        stock_allowed_from_query_context = (
            q_metric in ["stock", "negative_sales"]
            and (bare_value_feedback or not explicit_metrics or "stock" in explicit_metrics)
            and not (explicit_metrics - {"stock"})
        )
        if not already_stock_specific and (feedback_mentions_stock or stock_allowed_from_query_context):
            add(
                "stock",
                value,
                display_feedback,
                configured_metric_hint_from_query(query_text, "stock"),
                "Stock correction detected.",
                "high" if q_metric in ["stock", "negative_sales"] or feedback_mentions_stock else "medium",
            )

    # Cooler / POS generic numeric corrections for extra sheets.
    if value is not None and not decisions:
        if q_metric == "cooler" or re.search(r"cooler|freezer", f_lower):
            add("cooler", value, display_feedback, "cooler", "Cooler/freezer correction detected.", "medium")
        elif q_metric == "pos" or re.search(r"signage|chair|table|runner|shop sign|pos", f_lower):
            add("pos", value, display_feedback, "pos", "POS correction detected.", "medium")

    if not decisions:
        add(q_metric if q_metric != "other" else "other", None, display_feedback, "", "No clear correction value found.", "low")

    # Deduplicate exact metric+value+hint decisions caused by overlapping rules.
    unique: List[Dict[str, Any]] = []
    seen = set()
    for d in decisions:
        key = (d.get("answered_metric"), d.get("correction_value"), d.get("target_column_hint"))
        if key not in seen:
            unique.append(d)
            seen.add(key)
    return unique


# Optional AI hook left in place. Currently disabled by USE_AI_AGENT=False.
_AI_CLIENT = None
_AI_AVAILABLE = False


def get_ai_client():
    global _AI_CLIENT, _AI_AVAILABLE
    if _AI_CLIENT is not None:
        return _AI_CLIENT
    if not USE_AI_AGENT:
        _AI_AVAILABLE = False
        return None
    if not os.getenv("OPENAI_API_KEY"):
        _AI_AVAILABLE = False
        return None
    try:
        from openai import OpenAI
        _AI_CLIENT = OpenAI()
        _AI_AVAILABLE = True
        return _AI_CLIENT
    except Exception:
        _AI_AVAILABLE = False
        return None


def load_ai_cache(output_dir: str) -> Dict[str, Any]:
    path = Path(output_dir) / AI_CACHE_FILE
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_ai_cache(output_dir: str, cache: Dict[str, Any]) -> None:
    try:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        path = Path(output_dir) / AI_CACHE_FILE
        path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def interpret_feedback_multi(query_text: str, feedback_text: str, cache: Dict[str, Any]) -> List[Dict[str, Any]]:
    # AI is intentionally not used unless USE_AI_AGENT=True.
    return rule_interpret_feedback_multi(query_text, feedback_text)


# ═══════════════════════════════════════════════════════════════
# APPLY CORRECTIONS TO ROW
# ═══════════════════════════════════════════════════════════════


def is_previous_stock_col(col: str) -> bool:
    cn = normalize(col)
    return any(p in cn for p in ["previousstock", "prevstock", "openingstock", "lastmonthstock"])


def is_stock_component_col(col: str) -> bool:
    return stock_component_kind_from_text(col) is not None


def is_price_col(col: str) -> bool:
    if is_metric_excluded_col(col):
        return False
    cn = normalize(col)
    config = PROJECT_CONFIGS.get(PROJECT_NAME, {})
    configured = [normalize(c) for c in config.get("price_cols", []) if c]
    if any(x and x in cn for x in configured):
        return True
    return any(p in cn for p in ["sellingprice", "captureprice", "retailprice", "buyingprice", "price"])


def is_purchase_col(col: str) -> bool:
    if is_metric_excluded_col(col):
        return False
    cn = normalize(col)
    return any(p in cn for p in ["purchase", "purchases", "totalpurchase", "stockincrease"])


def is_stock_col(col: str) -> bool:
    if is_metric_excluded_col(col):
        return False
    cn = normalize(col)
    if is_previous_stock_col(col) or is_stock_component_col(col) or is_price_col(col) or is_purchase_col(col):
        return False
    return any(p in cn for p in ["totalstock", "currentstock", "stock"])


def is_cooler_col(col: str) -> bool:
    cn = normalize(col)
    return any(p in cn for p in ["brandedcooler", "unbrandedcooler", "chestfreezer", "coolerbank", "coolers"])


def is_pos_col(col: str) -> bool:
    cn = normalize(col)
    return any(p in cn for p in ["brandedsignage", "brandedchair", "brandedtable", "brandedtablerunner", "brandedshopsign", "pos"])


def component_hint_matches_col(hint: str, col: str) -> bool:
    hint_kind = stock_component_kind_from_text(hint)
    col_kind = stock_component_kind_from_text(col)
    return bool(hint_kind and col_kind and hint_kind == col_kind)



def column_base_matches_hint(col: str, hint: str) -> bool:
    """True when a column is the exact configured metric or a history version of it.

    Examples:
    - hint 'Capture Price Excl Container' matches 'Capture Price Excl Container (May26)'
    - hint 'Total Stock' matches 'Total Stock Sum (May26)' and 'Total Stock (May26)'
    """
    hn = normalize(hint)
    cn = normalize(col)
    if not hn:
        return False
    if hn in cn:
        return True
    # Accept Sum history variants, e.g. Total Purchase Sum (May26).
    if hn.replace("s", "") and hn.replace("s", "") in cn:
        return True
    return False


def latest_available_month_from_columns(columns: List[str], current_month: Optional[Tuple[int, int]]) -> Optional[Tuple[int, int]]:
    if current_month:
        return current_month
    found: List[Tuple[int, int]] = []
    for col in columns:
        ym = month_year_in_text(col)
        if ym:
            found.append(ym)
    return found[-1] if found else None

def candidate_columns_for_metric(columns: List[str], metric: str, hint: str = "", target_month: Optional[Tuple[int, int]] = None) -> List[str]:
    hint_n = normalize(hint)
    scored: List[Tuple[int, str]] = []

    for col in columns:
        cn = normalize(col)
        month_match = column_mentions_month(col, target_month)
        hint_match = column_base_matches_hint(col, hint)
        score = -1

        if metric == "price":
            if is_price_col(col):
                score = max(score, 30)
                if hint_match:
                    score = max(score, 85)
                if month_match:
                    score = max(score, 90)
                if hint_match and month_match:
                    score = max(score, 120)

        elif metric == "purchase":
            if is_purchase_col(col):
                score = max(score, 30)
                if hint_match:
                    score = max(score, 85)
                if month_match:
                    score = max(score, 90)
                if hint_match and month_match:
                    score = max(score, 120)

        elif metric == "previous_stock":
            if target_month and month_match and is_stock_col(col):
                score = max(score, 100)
                if hint_match:
                    score = max(score, 120)
            elif is_previous_stock_col(col):
                score = max(score, 80)

        elif metric == "stock_component":
            if is_stock_component_col(col):
                score = max(score, 70)
                if component_hint_matches_col(hint, col):
                    score = max(score, 95)
                if month_match:
                    score = max(score, 90)

        elif metric == "stock":
            if is_stock_col(col):
                score = max(score, 30)
                if hint_match:
                    score = max(score, 85)
                if month_match:
                    score = max(score, 90)
                if hint_match and month_match:
                    score = max(score, 120)

        elif metric == "cooler":
            if is_cooler_col(col):
                score = max(score, 40)
                if hint_n and (hint_n in cn or cn in hint_n):
                    score = max(score, 80)
                if month_match:
                    score = max(score, 90)

        elif metric == "pos":
            if is_pos_col(col):
                score = max(score, 40)
                if hint_n and (hint_n in cn or cn in hint_n):
                    score = max(score, 80)
                if month_match:
                    score = max(score, 90)

        if score >= 0:
            # Prefer non-helper columns.
            if re.fullmatch(r"xx(?:\.\d+)?", clean_text(col), flags=re.IGNORECASE):
                score -= 1000
            scored.append((score, col))

    scored.sort(key=lambda x: (-x[0], columns.index(x[1]) if x[1] in columns else 9999))
    return list(dict.fromkeys([col for _, col in scored]))

def pick_best_column(candidates: List[str], row: pd.Series, metric: str, hint: str, feedback_text: str, query_text: str = "") -> Optional[str]:
    if not candidates:
        return None

    combined = normalize(f"{hint} {feedback_text} {query_text}")

    if metric == "stock_component":
        requested_kind = stock_component_kind_from_text(f"{hint} {feedback_text} {query_text}")
        if requested_kind:
            exact_kind_matches = [c for c in candidates if stock_component_kind_from_text(c) == requested_kind]
            if exact_kind_matches:
                return exact_kind_matches[0]
        return None

    # Exact configured metric label wins, especially for NG-MRA Capture Price vs Capture Price Excl Container.
    if hint:
        exact_hint_matches = [c for c in candidates if column_base_matches_hint(c, hint)]
        if exact_hint_matches:
            month_exact = [c for c in exact_hint_matches if column_mentions_any_month(c)]
            if month_exact and any(column_mentions_any_month(c) for c in candidates):
                return month_exact[0]
            return exact_hint_matches[0]

    if metric in ["price", "purchase", "stock", "cooler", "pos"] and len(candidates) > 1:
        explicit = [c for c in candidates if normalize(c) and normalize(c) in combined]
        if explicit:
            return explicit[0]

    return candidates[0]

def purchase_day_number_from_col(col: str) -> Optional[int]:
    cn = normalize(col)
    patterns = [
        r"^pday0?([1-9]|[12][0-9]|3[01])$",
        r"^pd0?([1-9]|[12][0-9]|3[01])$",
        r"purchase(?:day)?0?([1-9]|[12][0-9]|3[01])$",
        r"day0?([1-9]|[12][0-9]|3[01])purchase",
    ]
    for pattern in patterns:
        m = re.search(pattern, cn)
        if m:
            return int(m.group(1))
    return None


def purchase_day_cols(columns: List[str]) -> Dict[int, str]:
    out: Dict[int, str] = {}
    for col in columns:
        day = purchase_day_number_from_col(col)
        if day is not None and day not in out:
            out[day] = col
    return out


def days_for_purchase_week(week: int) -> List[int]:
    if week <= 1:
        return list(range(1, 8))
    if week == 2:
        return list(range(8, 15))
    if week == 3:
        return list(range(15, 22))
    if week == 4:
        return list(range(22, 29))
    return list(range(29, 32))


def parse_purchase_hint(hint: str) -> Dict[str, Optional[int]]:
    data: Dict[str, Optional[int]] = {"week": None, "day": None}
    h = clean_text(hint).lower()
    m = re.search(r"day\s*=\s*(\d{1,2})", h)
    if m:
        data["day"] = int(m.group(1))
    m = re.search(r"week\s*=\s*(\d{1,2})", h)
    if m:
        data["week"] = int(m.group(1))
    return data


def choose_purchase_day_column(row: pd.Series, hint: str, feedback_text: str, query_text: str) -> Optional[str]:
    columns = list(row.index)
    day_cols = purchase_day_cols(columns)
    if not day_cols:
        return None

    period = parse_purchase_hint(hint)
    if not period.get("day") and not period.get("week"):
        text_period = extract_purchase_period(f"{feedback_text} {query_text}")
        period["day"] = text_period.get("day")
        period["week"] = text_period.get("week")

    if period.get("day") and period["day"] in day_cols:
        return day_cols[period["day"]]

    if period.get("week"):
        candidate_days = [d for d in days_for_purchase_week(period["week"]) if d in day_cols]
        if not candidate_days:
            return None
        populated = [d for d in candidate_days if safe_number(row.get(day_cols[d])) not in [None, 0.0]]
        chosen_day = populated[0] if populated else candidate_days[0]
        return day_cols[chosen_day]

    return None


def recalculate_total_purchase_from_days(row: pd.Series, current_month: Optional[Tuple[int, int]]) -> Tuple[pd.Series, str]:
    columns = list(row.index)
    day_cols = purchase_day_cols(columns)
    if not day_cols:
        return row, "No purchase day columns found for recalculating purchases"

    total = 0.0
    for day in sorted(day_cols):
        num = safe_number(row.get(day_cols[day]))
        total += num if num is not None else 0.0

    total_value = format_value(total)
    notes = [f"Recalculated purchases as {total_value}"]

    normal_candidates = [c for c in candidate_columns_for_metric(columns, "purchase") if not column_mentions_any_month(c)]
    if normal_candidates:
        row[normal_candidates[0]] = total_value
        notes.append(f"Updated '{normal_candidates[0]}' to {total_value}")

    if current_month:
        hist_candidates = [c for c in candidate_columns_for_metric(columns, "purchase", target_month=current_month) if column_mentions_month(c, current_month)]
        if hist_candidates and hist_candidates[0] not in normal_candidates[:1]:
            row[hist_candidates[0]] = total_value
            notes.append(f"Updated history purchase column '{hist_candidates[0]}' to {total_value}")

    return row, "; ".join(notes)


def recalculate_total_stock_from_components(row: pd.Series, current_month: Optional[Tuple[int, int]]) -> Tuple[pd.Series, str]:
    columns = list(row.index)
    component_cols = [c for c in columns if is_stock_component_col(c) and not column_mentions_any_month(c)]
    if not component_cols:
        return row, "No stock component columns found for recalculating total stock"

    total = 0.0
    found_numeric = False
    for col in component_cols:
        num = safe_number(row.get(col))
        if num is None:
            num = 0.0
        else:
            found_numeric = True
        total += num

    if not found_numeric:
        return row, "Stock component columns found, but none had numeric values"

    total_value = format_value(total)
    notes = [f"Recalculated total stock as {total_value}"]

    normal_candidates = [c for c in candidate_columns_for_metric(columns, "stock") if not column_mentions_any_month(c)]
    if normal_candidates:
        row[normal_candidates[0]] = total_value
        notes.append(f"Updated '{normal_candidates[0]}' to {total_value}")

    if current_month:
        hist_candidates = [c for c in candidate_columns_for_metric(columns, "stock", target_month=current_month) if column_mentions_month(c, current_month)]
        if hist_candidates and hist_candidates[0] not in normal_candidates[:1]:
            row[hist_candidates[0]] = total_value
            notes.append(f"Updated history stock column '{hist_candidates[0]}' to {total_value}")

    return row, "; ".join(notes)


def apply_value_to_row(row: pd.Series, metric: str, value: Optional[float], hint: str, feedback_text: str, query_text: str = "", current_month: Optional[Tuple[int, int]] = None) -> Tuple[pd.Series, str]:
    if value is None:
        return row, "No correction value to apply."

    columns = list(row.index)
    new_value = format_value(abs(value) if metric == "price" else value)
    notes: List[str] = []
    updated_cols: List[str] = []

    if metric == "purchase":
        purchase_day_target = choose_purchase_day_column(row, hint, feedback_text, query_text)
        if purchase_day_target:
            row[purchase_day_target] = new_value
            notes.append(f"Updated purchase day column '{purchase_day_target}' to {new_value}")
            row, total_note = recalculate_total_purchase_from_days(row, current_month)
            notes.append(total_note)
            return row, "; ".join(notes) + "."

    explicit_month = explicit_month_from_feedback(feedback_text, current_month)
    history_month = explicit_month
    if history_month is None:
        if metric == "previous_stock" and current_month:
            history_month = add_months(current_month[0], current_month[1], -1)
        elif metric in ["price", "stock", "purchase"]:
            # Dynamic latest month: do not hardcode May. Use the latest month available
            # for the exact metric/hint columns, falling back to current_month if needed.
            history_month = latest_metric_month_from_columns(columns, metric, hint, current_month=current_month)
        else:
            history_month = current_month

    # If the row has history columns for the target month, update the exact month column first.
    # This is important for Data Queries and Nigeria outlet sheets where the query names
    # a metric such as Capture Price Excl Container, Total Stock, or Total Purchase and
    # the correction should land in the latest matching month column.
    history_first_target = None
    if metric in ["price", "stock", "purchase"] and history_month:
        history_first = [
            c for c in candidate_columns_for_metric(columns, metric, hint, target_month=history_month)
            if column_mentions_month(c, history_month)
        ]
        history_first_target = pick_best_column(history_first, row, metric, hint, feedback_text, query_text)
        if history_first_target:
            row[history_first_target] = new_value
            updated_cols.append(history_first_target)
            notes.append(f"Updated history column '{history_first_target}' for {MONTH_DISPLAY[history_month[1]]} {history_month[0]} to {new_value}")

    if metric == "previous_stock":
        primary_candidates = [c for c in candidate_columns_for_metric(columns, "previous_stock", hint) if is_previous_stock_col(c)]
    elif metric == "stock_component":
        requested_kind = stock_component_kind_from_text(f"{hint} {feedback_text} {query_text}")
        all_component_candidates = [c for c in candidate_columns_for_metric(columns, "stock_component", hint) if not column_mentions_any_month(c)]
        primary_candidates = [c for c in all_component_candidates if stock_component_kind_from_text(c) == requested_kind] if requested_kind else []
    else:
        # Only update true non-month/current columns here. Do NOT fall back to an old
        # historical month like Dec25 just because there is no current non-month column.
        primary_candidates = [c for c in candidate_columns_for_metric(columns, metric, hint) if not column_mentions_any_month(c)]
        if not primary_candidates and not history_first_target:
            primary_candidates = candidate_columns_for_metric(columns, metric, hint)

    primary_target = pick_best_column(primary_candidates, row, metric, hint, feedback_text, query_text)
    if primary_target and primary_target not in updated_cols:
        row[primary_target] = new_value
        updated_cols.append(primary_target)
        notes.append(f"Updated '{primary_target}' to {new_value}")
    elif not primary_target:
        notes.append(f"No normal/current column found for metric '{metric}'")

    if metric == "stock_component":
        if not primary_target:
            notes.append("Total Stock was not recalculated because no exact stock component column was corrected")
            return row, "; ".join(notes) + "."
        row, total_note = recalculate_total_stock_from_components(row, current_month)
        notes.append(total_note)
        return row, "; ".join(notes) + "."

    history_metric = "stock" if metric == "previous_stock" else metric
    history_target = None

    if history_month:
        history_candidates = candidate_columns_for_metric(columns, history_metric, hint, target_month=history_month)
        history_candidates = [c for c in history_candidates if column_mentions_month(c, history_month)]
        if metric == "previous_stock":
            history_candidates = [c for c in history_candidates if is_stock_col(c)]
        history_target = pick_best_column(history_candidates, row, history_metric, hint, feedback_text, query_text)

    if history_target and history_target not in updated_cols:
        row[history_target] = new_value
        notes.append(f"Updated history column '{history_target}' for {MONTH_DISPLAY[history_month[1]]} {history_month[0]} to {new_value}")
    elif history_month:
        notes.append(f"No matching history month column found for {MONTH_DISPLAY[history_month[1]]} {history_month[0]}")
    else:
        notes.append("No current/history month could be inferred")

    return row, "; ".join(notes) + "."


def target_metric_for_numeric_sheet(sheet_name: str) -> str:
    sn = normalize(sheet_name)
    if "purchase" in sn:
        return "purchase"
    return "stock"


def apply_nigeria_numeric_replacement(row: pd.Series, sheet_name: str, numeric_value: Any, feedback_text: str, current_month: Optional[Tuple[int, int]]) -> Tuple[pd.Series, str]:
    """Apply the last numeric feedback column for Nigeria Outlet Stock/Purchases.

    Important: this numeric feedback is SKU-level. It must update the individual
    monthly column such as 'Total Purchase (May26)' or 'Total Stock (May26)', NOT
    the outlet-level 'Total Purchase Sum (May26)' / 'Total Stock Sum (May26)' column.
    """
    value = safe_number(numeric_value)
    if value is None:
        return row, "No numeric replacement feedback value; existing value kept."

    metric = target_metric_for_numeric_sheet(sheet_name)
    columns = list(row.index)
    target_month = explicit_month_from_feedback(feedback_text, current_month)
    if target_month is None:
        target_month = latest_metric_month_from_columns(columns, metric, "Total Purchase" if metric == "purchase" else "Total Stock", current_month=current_month, prefer_individual=True)

    candidates = candidate_columns_for_metric(columns, metric, hint="Total Purchase" if metric == "purchase" else "Total Stock", target_month=target_month)

    # SKU-level numeric feedback must go to individual columns, never Sum columns.
    individual_candidates = [c for c in candidates if "sum" not in normalize(c)]
    if target_month:
        individual_candidates = [c for c in individual_candidates if column_mentions_month(c, target_month)] or individual_candidates

    target = pick_best_column(individual_candidates, row, metric, "Total Purchase" if metric == "purchase" else "Total Stock", feedback_text, "")
    if not target:
        return row, f"Numeric replacement value {format_value(value)} found, but no individual target {metric} month column was found."

    row[target] = format_value(value)
    month_note = f" for {MONTH_DISPLAY[target_month[1]]} {target_month[0]}" if target_month else ""
    return row, f"Applied Nigeria numeric feedback replacement: updated individual column '{target}'{month_note} to {format_value(value)}."




def is_outlet_id_like_col(col: str) -> bool:
    """True for outlet ID/code columns, not outlet descriptive columns like Outletname."""
    cn = normalize(col)
    if not cn or bad_identifier_column_name(col, "outlet"):
        return False
    config = PROJECT_CONFIGS.get(PROJECT_NAME, {})
    configured = normalize(config.get("outlet_id")) if config.get("outlet_id") else ""
    if configured and (cn == configured or re.fullmatch(rf"{re.escape(configured)}\d+", cn)):
        return True
    return any(p in cn for p in ["outletid", "outletnumber", "outletno", "outnumber", "whoutletid", "projectoutletid"])


def is_sku_id_like_col(col: str) -> bool:
    """True for SKU/product ID/code columns, not SKU descriptions."""
    cn = normalize(col)
    if not cn or bad_identifier_column_name(col, "sku"):
        return False
    config = PROJECT_CONFIGS.get(PROJECT_NAME, {})
    configured = normalize(config.get("sku_id")) if config.get("sku_id") else ""
    if configured and (cn == configured or re.fullmatch(rf"{re.escape(configured)}\d+", cn)):
        return True
    return any(p in cn for p in ["whskuid", "skuid", "sku_id", "prodcode", "productcode"])


def remove_unmatched_duplicate_identity_columns(
    cols: List[str],
    sheet_name: str,
    selected_identity: Optional[Dict[str, Optional[str]]] = None,
) -> List[str]:
    """For Nigeria Outlet Stock/Purchases, keep only the ID columns used for matching.

    Some files contain both a visual/merged outlet ID column and a repeated detail-row
    outlet ID column. The final output should keep the one selected for matching and
    drop the other duplicate ID column. This is based on population/format, not on a
    hardcoded name.
    """
    if not (PROJECT_NAME in NIGERIA_PROJECTS and normalize(sheet_name) in NIGERIA_NUMERIC_FEEDBACK_SHEETS):
        return cols

    selected_identity = selected_identity or {}
    selected_outlet = selected_identity.get("outlet")
    selected_sku = selected_identity.get("sku")

    out: List[str] = []
    for col in cols:
        if selected_outlet and is_outlet_id_like_col(col) and normalize(col) != normalize(selected_outlet):
            continue
        if selected_sku and is_sku_id_like_col(col) and normalize(col) != normalize(selected_sku):
            continue
        out.append(col)
    return out




# ═══════════════════════════════════════════════════════════════
# STRICT REQUESTED-METRIC GATING + LATEST-FEEDBACK HELPERS
# ═══════════════════════════════════════════════════════════════


def is_nigeria_outlet_numeric_sheet(sheet_name: str) -> bool:
    return PROJECT_NAME in NIGERIA_PROJECTS and normalize(sheet_name) in NIGERIA_NUMERIC_FEEDBACK_SHEETS


def canonical_metric(metric: str) -> str:
    m = clean_text(metric).lower()
    if m in {"stock_component", "current_stock_component"}:
        return "stock_component"
    if m in {"previous_stock", "prev_stock", "opening_stock"}:
        return "previous_stock"
    if m in {"negative_sales", "negative sale"}:
        return "negative_sales"
    if m in {"purchase", "purchases", "total_purchase"}:
        return "purchase"
    if m in {"price", "selling_price", "capture_price"}:
        return "price"
    if m in {"stock", "total_stock"}:
        return "stock"
    if m in {"cooler", "freezer"}:
        return "cooler"
    if m in {"pos", "pos_material"}:
        return "pos"
    if m in {"dynamic_attribute", "dynamic", "attribute"}:
        return "dynamic_attribute"
    return m or "other"


def metrics_requested_by_query(query_text: str, sheet_name: str = "") -> set:
    """Return the metrics that the query actually requested.

    This prevents a feedback response about purchases from changing purchase columns
    when the query was asking about Total Stock, and vice versa.

    For Nigeria Outlet Stock / Outlet Purchases sheets, there may be no normal query
    text, so the sheet itself defines the requested metric.
    """
    q = clean_text(query_text).lower()
    qn = normalize(query_text)

    if not q:
        sn = normalize(sheet_name)
        if sn == "outletstock":
            return {"stock"}
        if sn == "outletpurchases":
            return {"purchase"}
        return set()

    requested = set()

    # Flexible sheets such as Cooler Queries and POS Queries are corrected by
    # matching the actual query labels to the actual sheet columns dynamically.
    # Do not depend on a fixed list like branded/unbranded/chest freezer.
    if is_dynamic_attribute_sheet(sheet_name):
        requested.add("dynamic_attribute")

    config = PROJECT_CONFIGS.get(PROJECT_NAME, {})

    # Price: exact configured price columns matter, especially Capture Price vs
    # Capture Price Excl Container. They are both metric=price but the hint chooses
    # the exact target column later.
    for price_col in config.get("price_cols", []) or []:
        if normalize(price_col) and normalize(price_col) in qn:
            requested.add("price")
    if re.search(r"\b(selling|price|capture\s+price|retail\s+price|buying\s+price|profit|margin)\b", q):
        requested.add("price")

    # Purchases.
    purchase_col = clean_text(config.get("purchase_col")) or ""
    if purchase_col and normalize(purchase_col) in qn:
        requested.add("purchase")
    if re.search(r"\b(total\s+purchase|purchase|purchases|purchased|stock\s+increase|bought|buy)\b", q):
        requested.add("purchase")

    # Previous stock before generic stock.
    if re.search(r"\b(previous\s+stock|prev\s+stock|opening\s+stock|last\s+month\s+stock)\b", q):
        requested.add("previous_stock")

    # Stock components.
    if stock_component_kind_from_text(query_text):
        requested.add("stock_component")

    # Total/current stock.
    stock_col = clean_text(config.get("stock_col")) or ""
    if stock_col and normalize(stock_col) in qn:
        requested.add("stock")
    if re.search(r"\b(total\s+stock|current\s+stock|stock|unit\s+count|available|count)\b", q):
        # Do not convert previous-stock-only queries to generic stock.
        if "previous_stock" not in requested or re.search(r"\btotal\s+stock|current\s+stock\b", q):
            requested.add("stock")

    if re.search(r"\b(cooler|freezer)\b", q):
        requested.add("cooler")
    if re.search(r"\b(signage|chair|table|runner|shop\s+sign|pos)\b", q):
        requested.add("pos")

    # Negative sales is special: it is not a free pass to update any random metric,
    # but the old correction workflow often allowed stock/purchase explanations.
    # Keep this limited to stock/purchase/previous_stock only.
    if re.search(r"\bnegative\s+sales|neg\s*sales\b", q):
        requested.update({"stock", "purchase", "previous_stock"})

    return requested


def metric_allowed_by_query(metric: str, requested_metrics: set) -> bool:
    """Strictly check whether a detected feedback correction is allowed."""
    m = canonical_metric(metric)
    if not requested_metrics:
        return False
    if m in requested_metrics:
        return True
    # Only allow stock_component if the query specifically mentions a stock component.
    # Do NOT treat generic Total Stock as permission to change front/back stock.
    return False


def filter_decisions_by_requested_metrics(decisions: List[Dict[str, Any]], requested_metrics: set) -> Tuple[List[Dict[str, Any]], List[str]]:
    allowed: List[Dict[str, Any]] = []
    skipped: List[str] = []
    for d in decisions:
        metric = canonical_metric(d.get("answered_metric", ""))
        value = d.get("correction_value")
        if metric_allowed_by_query(metric, requested_metrics):
            allowed.append(d)
        else:
            if value is not None:
                skipped.append(
                    f"Skipped {metric} correction because query requested {sorted(requested_metrics) or 'no metric'}"
                )
    return allowed, skipped


def metric_group_for_registry(metric: str) -> str:
    """Metric key for cross-sheet priority checks."""
    m = canonical_metric(metric)
    if m == "stock_component":
        return "stock_component"
    if m == "previous_stock":
        return "previous_stock"
    return m


def correction_registry_key(row: pd.Series, identity_cols: Dict[str, Optional[str]], metric: str) -> Tuple[str, str, str]:
    outlet = clean_text(row.get(identity_cols.get("outlet"), "")) if identity_cols.get("outlet") else ""
    sku = clean_text(row.get(identity_cols.get("sku"), "")) if identity_cols.get("sku") else ""
    return (normalize(outlet), normalize(sku), metric_group_for_registry(metric))


def build_latest_feedback_context_for_sheet(sheet_name: str, feedback_files: Dict[int, str]) -> Dict[str, Dict[str, List[Dict[str, Any]]]]:
    """Build latest feedback lookup across all batches for one sheet.

    If the same outlet+SKU(+query) appears in Batch 1, 2 and 3, the latest batch
    wins. This is applied at lookup time while still preserving all rows in output.
    """
    exact: Dict[str, List[Dict[str, Any]]] = {}
    loose: Dict[str, List[Dict[str, Any]]] = {}

    for batch in sorted(BATCH_NUMBERS):
        fb_file = feedback_files.get(batch)
        if not fb_file or not workbook_has_sheet(fb_file, sheet_name):
            continue
        try:
            feedback_df, _ = read_sheet(fb_file, sheet_name)
        except Exception as e:
            print(f"  Could not build feedback lookup for Batch {batch}, sheet '{sheet_name}': {e}")
            continue

        fb_identity_cols = detect_identity_columns(feedback_df)
        fb_cols = detect_feedback_columns_for_sheet(feedback_df, sheet_name)

        for idx, row in feedback_df.iterrows():
            entry = {
                "batch": batch,
                "row_index": idx,
                "row": row,
                "identity_cols": fb_identity_cols,
                "feedback_cols": fb_cols,
                "feedback_file": fb_file,
            }
            k = row_key(row, fb_identity_cols, include_query=True)
            lk = fallback_key(row, fb_identity_cols)
            if k:
                exact.setdefault(k, []).append(entry)
            if lk:
                loose.setdefault(lk, []).append(entry)

    return {"exact": exact, "loose": loose}


def select_latest_feedback_entries(entries: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not entries:
        return []
    latest_batch = max(int(e.get("batch", 0)) for e in entries)
    return [e for e in entries if int(e.get("batch", 0)) == latest_batch]


def lookup_latest_feedback_entries(
    latest_context: Dict[str, Dict[str, List[Dict[str, Any]]]],
    q_row_original: pd.Series,
    q_cols: Dict[str, Optional[str]],
) -> List[Dict[str, Any]]:
    q_key = row_key(q_row_original, q_cols, include_query=True)
    q_loose = fallback_key(q_row_original, q_cols)
    entries = latest_context.get("exact", {}).get(q_key, [])
    if not entries:
        entries = latest_context.get("loose", {}).get(q_loose, [])
    return select_latest_feedback_entries(entries)

# ═══════════════════════════════════════════════════════════════
# OUTPUT COLUMNS / FORMATTING
# ═══════════════════════════════════════════════════════════════



def insert_feedback_before_sku_for_outlet_sheet(cols: List[str], sheet_name: str) -> List[str]:
    """For Nigeria Outlet Stock/Purchases, place Feedback at the end of the
    outlet-level block, immediately before the SKU/detail section.

    This is formatting only. It does not change matching or correction logic.
    Desired visual layout:
        outlet details + outlet sum columns + Feedback | SKU/details/month values
    """
    if not (PROJECT_NAME in NIGERIA_PROJECTS and normalize(sheet_name) in NIGERIA_NUMERIC_FEEDBACK_SHEETS):
        return cols

    cols = [c for c in cols if normalize(c) != normalize(FEEDBACK_COL)]
    tmp_df = pd.DataFrame(columns=cols)
    sku_idx = find_sku_column_index(tmp_df)

    if sku_idx is None:
        cols.append(FEEDBACK_COL)
    else:
        cols.insert(sku_idx, FEEDBACK_COL)
    return cols


def build_final_columns(format_cols: List[str], sheet_name: str, selected_identity: Optional[Dict[str, Optional[str]]] = None) -> List[str]:
    # Remove helper columns and numeric replacement feedback columns from format.
    helper_norms = {normalize(BATCH_COL), normalize(ORIGINAL_FEEDBACK_COL), normalize(AI_NOTES_COL)}
    cols = [c for c in format_cols if normalize(c) not in helper_norms]
    if DROP_XX_HELPER_COLUMNS:
        cols = [c for c in cols if not re.fullmatch(r"xx(?:\.\d+)?", clean_text(c), flags=re.IGNORECASE)]

    # For Nigeria Outlet Stock/Purchases, drop duplicate ID columns that were not
    # selected for matching. Example: if one outlet column is visually merged/blank
    # and another is populated on every SKU row, keep only the populated one.
    cols = remove_unmatched_duplicate_identity_columns(cols, sheet_name, selected_identity)

    # Remove all existing feedback-like columns from the chosen format.
    # We add one clean matched Feedback column ourselves.
    tmp_df = pd.DataFrame(columns=cols)
    fb_existing = feedback_like_columns(tmp_df)
    if fb_existing:
        cols = [c for c in cols if c not in fb_existing]

    query_col = resolve_col(pd.DataFrame(columns=cols), [QUERY_COL, "Query", "Data Queries", "Question", "Title"])

    if PROJECT_NAME in NIGERIA_PROJECTS and normalize(sheet_name) in NIGERIA_NUMERIC_FEEDBACK_SHEETS:
        # Important formatting fix:
        # Feedback must sit at the end of the outlet-level merged block, before SKU rows.
        cols = insert_feedback_before_sku_for_outlet_sheet(cols, sheet_name)
    elif query_col and query_col in cols:
        pos = cols.index(query_col) + 1
        cols.insert(pos, FEEDBACK_COL)
    else:
        cols.append(FEEDBACK_COL)

    return [BATCH_COL] + cols + [ORIGINAL_FEEDBACK_COL, AI_NOTES_COL]

def safe_excel_sheet_name(name: str, used: set) -> str:
    base = re.sub(r"[\[\]\:\*\?\/\\]", "_", clean_text(name))[:31] or "Sheet"
    candidate = base
    i = 1
    while candidate in used:
        suffix = f"_{i}"
        candidate = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(candidate)
    return candidate


def autofit_worksheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    workbook = writer.book
    ws = writer.sheets[sheet_name]

    # Match the expected output colours:
    # - Normal headers are blue.
    # - Feedback headers are green.
    # - No borders are added.
    # - Text must NOT be wrapped in any output sheet.
    header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#4472C4",
        "font_color": "white",
        "text_wrap": False,
    })
    feedback_header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#00B050",
        "font_color": "white",
        "text_wrap": False,
    })
    normal_body_fmt = workbook.add_format({
        "text_wrap": False,
        "valign": "top",
    })
    feedback_body_fmt = workbook.add_format({
        "text_wrap": False,
        "valign": "top",
        "font_color": "#00B050",
    })

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))

    # Apply no-wrap formatting and a fixed normal row height to the whole used range.
    # Fixed height is important because Excel may keep tall row heights from values
    # that previously contained embedded newlines.
    normal_row_height = 15
    if len(df.columns) > 0:
        ws.set_row(0, normal_row_height, header_fmt)
        for row_idx in range(1, len(df) + 1):
            ws.set_row(row_idx, normal_row_height, normal_body_fmt)

    for i, col in enumerate(df.columns):
        col_norm = normalize(col)
        is_feedback_col = col_norm in {normalize(FEEDBACK_COL), "feedback", "feedbacks"}
        ws.write(0, i, col, feedback_header_fmt if is_feedback_col else header_fmt)

        sample = df[col].astype(str).head(200).tolist() if col in df.columns else []
        max_len = max([len(str(col))] + [len(x) for x in sample])
        width = min(max(max_len + 2, 12), 45)
        if col in [FEEDBACK_COL, ORIGINAL_FEEDBACK_COL, AI_NOTES_COL, QUERY_COL, "Queries"]:
            width = 45

        ws.set_column(i, i, width, feedback_body_fmt if is_feedback_col else normal_body_fmt)

def find_sku_column_index(df: pd.DataFrame) -> Optional[int]:
    config = PROJECT_CONFIGS.get(PROJECT_NAME, {})
    candidates = [
        config.get("sku_id"),
        "wh_skuid", "SKU ID", "Sku ID", "SKU_ID", "SKU", "Sku", "Prod Code", "Prodcode", "Product Code", "product code", "sku_id",
    ]
    col = resolve_col_best_nonblank(df, candidates, role="sku")
    if col and col in df.columns:
        return list(df.columns).index(col)
    return None


def find_outlet_column_index(df: pd.DataFrame) -> Optional[int]:
    config = PROJECT_CONFIGS.get(PROJECT_NAME, {})
    candidates = [
        config.get("outlet_id"),
        "Outlet ID", "OutletID", "Outlet Id", "Outlet Number", "Out Number", "OUTNUMBER",
        "outletid", "outlet_id", "wh_outletid", "projectOutletid", "Outlet No", "OutletNo",
    ]
    sku_idx = find_sku_column_index(df)
    detail_scope = None
    if sku_idx is not None:
        detail_scope = df.iloc[:, sku_idx].apply(lambda x: not is_blank(x))
    col = resolve_col_best_nonblank(df, candidates, role="outlet", scope_mask=detail_scope)
    if col and col in df.columns:
        return list(df.columns).index(col)
    return None


def first_nonblank_value(df: pd.DataFrame, start: int, end: int, col_index: int) -> Any:
    """
    Return the first nonblank value inside a block.

    This is used for outlet-level merged display columns so the merged cell
    does not appear blank just because the first row in the block was blank.
    """
    for r in range(start, end + 1):
        value = df.iloc[r, col_index]
        if not is_blank(value):
            return value
    return ""


def forward_filled_series_for_merge(df: pd.DataFrame, col_index: int) -> List[str]:
    """
    Create a forward-filled version of a column for grouping only.

    This solves the issue where Excel/pandas reads visually merged outlet cells
    as one value followed by blanks. We use the last seen nonblank value so the
    blank rows are still treated as belonging to the same outlet block.
    """
    filled: List[str] = []
    last_value = ""

    for i in range(len(df)):
        value = clean_text(df.iloc[i, col_index])
        if value:
            last_value = value
        filled.append(last_value)

    return filled


def build_outlet_block_keys_for_merge(
    df: pd.DataFrame,
    outlet_idx: int,
) -> List[Tuple[str, str]]:
    """
    Build block keys for Outlet Stock / Outlet Purchases formatting.

    The key is:
        Batch + forward-filled Outlet ID

    This prevents blank outlet cells from breaking one outlet block into many
    wrong blocks.
    """
    outlet_values = forward_filled_series_for_merge(df, outlet_idx)

    if BATCH_COL in df.columns:
        batch_idx = list(df.columns).index(BATCH_COL)
        batch_values = forward_filled_series_for_merge(df, batch_idx)
    else:
        batch_values = [""] * len(df)

    return [
        (clean_text(batch_values[i]), clean_text(outlet_values[i]))
        for i in range(len(df))
    ]


def outlet_block_ranges_from_keys(keys: List[Tuple[str, str]]) -> List[Tuple[int, int]]:
    """
    Convert repeated outlet keys into row ranges.

    Example:
        rows 0-12 outlet A
        rows 13-25 outlet B
    """
    if not keys:
        return []

    ranges: List[Tuple[int, int]] = []
    start = 0
    current = keys[0]

    for i in range(1, len(keys)):
        if keys[i] != current:
            ranges.append((start, i - 1))
            start = i
            current = keys[i]

    ranges.append((start, len(keys) - 1))
    return ranges


def merge_repeated_outlet_blocks_xlsxwriter(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
    sheet_name: str,
    logical_sheet_name: str,
) -> None:
    """
    Format Nigeria Outlet Stock / Outlet Purchases as outlet blocks.

    Rules:
    - Only applies to Nigeria Outlet Stock and Outlet Purchases.
    - Outlet-level columns are visually merged down the full outlet block.
    - Outlet sum/history columns are visually merged down the full outlet block.
    - Feedback is visually merged down the full outlet block.
    - SKU/product columns are NOT merged.
    - No borders are added.
    - Blank outlet cells do NOT break the merge block.
    - Feedback keeps the green formatting.
    """
    if not (
        PROJECT_NAME in NIGERIA_PROJECTS
        and normalize(logical_sheet_name) in NIGERIA_NUMERIC_FEEDBACK_SHEETS
    ):
        return

    if df.empty:
        return

    ws = writer.sheets[sheet_name]
    workbook = writer.book

    merge_fmt = workbook.add_format({
        "valign": "vcenter",
        "align": "center",
        "text_wrap": False,
    })
    feedback_merge_fmt = workbook.add_format({
        "valign": "vcenter",
        "align": "center",
        "text_wrap": False,
        "font_color": "#00B050",
    })

    outlet_idx = find_outlet_column_index(df)
    sku_idx = find_sku_column_index(df)

    if outlet_idx is None or sku_idx is None:
        return

    if sku_idx <= 0:
        return

    # Build outlet blocks using forward-filled Batch + Outlet ID.
    # This is the key rule: blank outlet cells belong to the outlet above them.
    block_keys = build_outlet_block_keys_for_merge(df, outlet_idx)
    block_ranges = outlet_block_ranges_from_keys(block_keys)

    # Merge all outlet-level columns before the SKU/detail section.
    merge_cols = list(range(0, sku_idx))

    # Feedback is deliberately placed before the SKU section by build_final_columns().
    # If it is found elsewhere due to an unusual format source, still merge it.
    if FEEDBACK_COL in df.columns:
        feedback_idx = list(df.columns).index(FEEDBACK_COL)
        if feedback_idx not in merge_cols:
            merge_cols.append(feedback_idx)
    else:
        feedback_idx = None

    # Never merge row-level audit helper columns.
    blocked = {normalize(ORIGINAL_FEEDBACK_COL), normalize(AI_NOTES_COL)}
    merge_cols = sorted(set(c for c in merge_cols if normalize(df.columns[c]) not in blocked))

    for start, end in block_ranges:
        if end < start:
            continue

        for c in merge_cols:
            value = first_nonblank_value(df, start, end, c)
            if is_blank(value):
                continue

            fmt = feedback_merge_fmt if normalize(df.columns[c]) == normalize(FEEDBACK_COL) else merge_fmt

            if end > start:
                ws.merge_range(start + 1, c, end + 1, c, value, fmt)
            else:
                ws.write(start + 1, c, value, fmt)

    # Keep row heights normal after merge operations as well.
    for row_idx in range(0, len(df) + 1):
        ws.set_row(row_idx, 15)

def requested_signature_from_query(query_text: str, sheet_name: str = "") -> str:
    """Stable signature used for latest-batch de-duplication."""
    requested = metrics_requested_by_query(query_text, sheet_name)
    if not requested:
        sn = normalize(sheet_name)
        if sn == "outletstock":
            return "stock:TotalStock"
        if sn == "outletpurchases":
            return "purchase:TotalPurchase"
        return normalize(query_text) or normalize(sheet_name)

    parts: List[str] = []
    for metric in sorted(requested):
        if metric == "price":
            parts.append(f"price:{normalize(configured_metric_hint_from_query(query_text, 'price'))}")
        elif metric == "purchase":
            parts.append(f"purchase:{normalize(configured_metric_hint_from_query(query_text, 'purchase') or 'Total Purchase')}")
        elif metric in {"stock", "previous_stock"}:
            parts.append(f"{metric}:{normalize(configured_metric_hint_from_query(query_text, 'stock') or 'Total Stock')}")
        else:
            parts.append(metric)
    return "|".join(parts)


def output_latest_key(row: pd.Series, identity_cols: Dict[str, Optional[str]], sheet_name: str) -> str:
    outlet = clean_text(row.get(identity_cols.get("outlet"), "")) if identity_cols.get("outlet") else ""
    sku = clean_text(row.get(identity_cols.get("sku"), "")) if identity_cols.get("sku") else ""
    query_text = clean_text(row.get(identity_cols.get("query"), "")) if identity_cols.get("query") else ""
    sig = requested_signature_from_query(query_text, sheet_name)
    return normalize(f"{outlet}|{sku}|{sig}")


def outlet_block_latest_key(row: pd.Series, identity_cols: Dict[str, Optional[str]], sheet_name: str) -> str:
    """Latest-batch key for Nigeria Outlet Stock / Outlet Purchases.

    For these two sheets the output is an outlet-product block, not a single
    outlet+SKU correction row. If we deduplicate by outlet+SKU, some valid
    product/SKU rows can disappear. The correct latest-batch rule is:

        same outlet appears in several batches -> keep the latest full outlet block
        inside that latest block -> keep every product/SKU row exactly as-is

    The caller forward-fills the outlet ID before calling this function, so blank
    cells created by Excel merged cells still belong to the outlet above them.
    """
    outlet = clean_text(row.get(identity_cols.get("outlet"), "")) if identity_cols.get("outlet") else ""
    if not outlet:
        return ""
    sn = normalize(sheet_name)
    return normalize(f"{outlet}|{sn}")


def forward_fill_identity_column_for_dedupe(work: pd.DataFrame, col: Optional[str]) -> None:
    """Forward-fill an identity column in-place only for de-duplication logic.

    This does not change the final displayed dataframe. It is used on the temporary
    `work` dataframe inside `deduplicate_latest_output_rows` so that rows under a
    visually merged outlet cell are grouped with the outlet above them.
    """
    if not col or col not in work.columns:
        return

    last_value = ""
    values = []
    for value in work[col].tolist():
        text = clean_text(value)
        if text:
            last_value = text
            values.append(value)
        else:
            values.append(last_value)
    work[col] = values


def deduplicate_latest_output_rows(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """Keep latest batch rows without deleting valid product/SKU rows.

    General sheets:
      - latest wins by outlet + SKU + requested metric/query signature.
      - repeated rows in the latest batch are preserved.

    Nigeria Outlet Stock / Outlet Purchases:
      - latest wins by FULL OUTLET BLOCK, not by outlet + SKU.
      - once the latest outlet block is selected, every product/SKU row inside
        that block is kept exactly as it appears in the feedback/output source.
      - this prevents products being left out from an outlet block.
    """
    if df.empty or BATCH_COL not in df.columns:
        return df

    identity = detect_identity_columns(df)
    if not identity.get("outlet"):
        return df

    work = df.copy()
    work["__row_order__"] = range(len(work))
    work["__batch_num__"] = pd.to_numeric(work[BATCH_COL], errors="coerce").fillna(-1)

    if is_nigeria_outlet_numeric_sheet(sheet_name):
        # This is the key fix: use the outlet block as the latest-batch unit.
        # Forward-fill the outlet ID only in this temporary dataframe so blank
        # rows under a merged outlet cell are still part of the same block.
        forward_fill_identity_column_for_dedupe(work, identity.get("outlet"))
        work["__latest_key__"] = work.apply(lambda r: outlet_block_latest_key(r, identity, sheet_name), axis=1)
    else:
        work["__latest_key__"] = work.apply(lambda r: output_latest_key(r, identity, sheet_name), axis=1)

    # Blank keys cannot be compared safely, so keep them individually.
    work.loc[work["__latest_key__"].eq(""), "__latest_key__"] = work.loc[
        work["__latest_key__"].eq(""), "__row_order__"
    ].map(lambda x: f"blank_{x}")

    max_batch_by_key = work.groupby("__latest_key__")["__batch_num__"].transform("max")
    work = work[work["__batch_num__"].eq(max_batch_by_key)].copy()
    work = work.sort_values("__row_order__")

    return work.drop(columns=["__latest_key__", "__row_order__", "__batch_num__"])

# ═══════════════════════════════════════════════════════════════
# PROCESSING LOGIC
# ═══════════════════════════════════════════════════════════════


def get_format_source_for_sheet(sheet_name: str) -> Dict[str, Any]:
    for k, v in FORMAT_SOURCE_BY_SHEET.items():
        if normalize(k) == normalize(sheet_name):
            return v
    return FORMAT_SOURCE_DEFAULT


def choose_format_file(sheet_name: str, query_files: Dict[int, str], feedback_files: Dict[int, str]) -> Tuple[Optional[str], Optional[int], Optional[str]]:
    fmt = get_format_source_for_sheet(sheet_name)
    source_type = clean_text(fmt.get("source_type", "query")).lower()
    preferred_batch = fmt.get("batch")
    file_map = feedback_files if source_type == "feedback" else query_files

    # First try requested source type + requested batch.
    if preferred_batch in file_map and workbook_has_sheet(file_map[preferred_batch], sheet_name):
        return file_map[preferred_batch], preferred_batch, source_type

    # Then try requested source type in any available batch.
    for batch in BATCH_NUMBERS:
        if batch in file_map and workbook_has_sheet(file_map[batch], sheet_name):
            return file_map[batch], batch, source_type

    # Last fallback: use the other source type.
    other_type = "query" if source_type == "feedback" else "feedback"
    other_map = query_files if other_type == "query" else feedback_files
    for batch in BATCH_NUMBERS:
        if batch in other_map and workbook_has_sheet(other_map[batch], sheet_name):
            return other_map[batch], batch, other_type

    return None, None, None


def process_batch_sheet(
    batch_no: int,
    sheet_name: str,
    query_file: str,
    feedback_file: str,
    format_cols: List[str],
    final_cols: List[str],
    ai_cache: Dict[str, Any],
    latest_feedback_context: Optional[Dict[str, Dict[str, List[Dict[str, Any]]]]] = None,
    correction_registry: Optional[Dict[Tuple[str, str, str], Dict[str, Any]]] = None,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Process one batch/sheet.

    Key rules enforced here:
    - Feedback is matched using the populated outlet/SKU ID columns, not helper columns.
    - If the same outlet+SKU appears in multiple batches, latest feedback wins.
    - Detected feedback corrections are applied ONLY if the query/sheet requested that metric.
    - Data Queries corrections are registered and override Outlet Stock / Outlet Purchases corrections.
    """
    correction_registry = correction_registry if correction_registry is not None else {}
    latest_feedback_context = latest_feedback_context or {"exact": {}, "loose": {}}

    query_df, _ = read_sheet(query_file, sheet_name)

    # For Nigeria Outlet Stock / Outlet Purchases, the product/SKU detail rows in
    # the feedback workbook are the source of truth for the final output layout.
    # Some query sheets do not contain every SKU/detail row that exists in the
    # feedback file. If we drive the output from the query sheet only, those SKU
    # rows are left out. So for these two sheets only, use the feedback sheet as
    # the row source while keeping all correction/matching logic unchanged.
    base_df = query_df
    base_is_feedback_sheet = False
    base_feedback_cols: Dict[str, Optional[str]] = {"text_feedback": None, "numeric_feedback": None}

    if is_nigeria_outlet_numeric_sheet(sheet_name):
        try:
            feedback_base_df, _ = read_sheet(feedback_file, sheet_name)
            if not feedback_base_df.empty:
                base_df = feedback_base_df
                base_is_feedback_sheet = True
                base_feedback_cols = detect_feedback_columns_for_sheet(base_df, sheet_name)
        except Exception as e:
            print(f"  Batch {batch_no}, sheet '{sheet_name}': could not use feedback rows as base, falling back to query rows: {e}")

    q_cols = detect_identity_columns(base_df)
    current_month = infer_current_month(query_file, feedback_file, str(Path(query_file).parent), str(Path(feedback_file).parent), " ".join(format_cols))

    if PROJECT_NAME in NIGERIA_PROJECTS and normalize(sheet_name) in NIGERIA_NUMERIC_FEEDBACK_SHEETS:
        source_label = "feedback" if base_is_feedback_sheet else "query"
        print(
            f"  Batch {batch_no}, sheet '{sheet_name}' output row source: {source_label}; "
            f"matching columns: outlet={q_cols.get('outlet')}, sku={q_cols.get('sku')}"
        )

    aligned_query = align_to_format(base_df, format_cols)
    output_rows: List[pd.Series] = []

    matched_rows = 0
    unmatched_rows = 0
    rows_with_feedback = 0
    skipped_by_metric = 0
    skipped_by_registry = 0

    for q_idx, q_row_original in base_df.iterrows():
        q_row = aligned_query.loc[q_idx].copy()

        if base_is_feedback_sheet:
            # The current feedback row is already the exact SKU/detail row that
            # should appear in output. Do not use a loose outlet+SKU lookup here,
            # because duplicate SKU rows inside the same outlet can be valid and
            # must not all receive the same combined feedback.
            feedback_entries = [{
                "batch": batch_no,
                "row_index": q_idx,
                "row": q_row_original,
                "identity_cols": q_cols,
                "feedback_cols": base_feedback_cols,
                "feedback_file": feedback_file,
            }]
        else:
            feedback_entries = lookup_latest_feedback_entries(latest_feedback_context, q_row_original, q_cols)

        if feedback_entries:
            matched_rows += 1
        else:
            unmatched_rows += 1

        query_text = clean_text(q_row_original.get(q_cols["query"], "")) if q_cols.get("query") else ""
        requested_metrics = metrics_requested_by_query(query_text, sheet_name)

        refined_feedbacks: List[str] = []
        original_feedbacks: List[str] = []
        ai_notes: List[str] = []

        for entry in feedback_entries:
            fb_row = entry["row"]
            fb_cols = entry.get("feedback_cols", {})
            source_batch = int(entry.get("batch", batch_no))

            original_fb = clean_text(fb_row.get(fb_cols.get("text_feedback"), "")) if fb_cols.get("text_feedback") else ""
            numeric_fb = fb_row.get(fb_cols.get("numeric_feedback"), "") if fb_cols.get("numeric_feedback") else ""

            if original_fb:
                original_feedbacks.append(f"Batch {source_batch}: {original_fb}")

                for feedback_part in split_feedback_into_parts(original_fb):
                    refined_display = refine_feedback_display(query_text, feedback_part)
                    if refined_display and refined_display not in refined_feedbacks:
                        refined_feedbacks.append(refined_display)

                    # Cooler Queries / POS Queries use flexible project-specific
                    # labels. Keep all existing DQ/Outlet Stock/Outlet Purchases
                    # logic untouched, but for these sheets update the actual
                    # matching columns found in the sheet.
                    if is_dynamic_attribute_sheet(sheet_name):
                        q_row, dynamic_note, dynamic_count = apply_dynamic_attribute_corrections(
                            q_row,
                            sheet_name,
                            query_text,
                            feedback_part,
                            current_month=current_month,
                        )
                        if dynamic_note:
                            ai_notes.append(dynamic_note)
                        continue

                    decisions = interpret_feedback_multi(query_text, feedback_part, ai_cache)
                    decisions, skipped_notes = filter_decisions_by_requested_metrics(decisions, requested_metrics)
                    for note in skipped_notes:
                        skipped_by_metric += 1
                        ai_notes.append(note)

                    for decision in decisions:
                        metric = canonical_metric(clean_text(decision.get("answered_metric")))
                        value = safe_number(decision.get("correction_value"))
                        hint = clean_text(decision.get("target_column_hint"))
                        notes = clean_text(decision.get("notes"))

                        # Data Queries overrides Outlet Stock/Outlet Purchases.
                        reg_key = correction_registry_key(q_row_original, q_cols, metric)
                        is_data_queries = normalize(sheet_name) == normalize("Data Queries")
                        if (not is_data_queries) and normalize(sheet_name) in NIGERIA_NUMERIC_FEEDBACK_SHEETS:
                            if reg_key in correction_registry:
                                skipped_by_registry += 1
                                ai_notes.append(
                                    f"Skipped {metric} correction because Data Queries already corrected this outlet/SKU/metric."
                                )
                                continue

                        q_row, apply_note = apply_value_to_row(
                            q_row,
                            metric,
                            value,
                            hint,
                            refined_display,
                            query_text=query_text,
                            current_month=current_month,
                        )

                        if is_data_queries and value is not None:
                            correction_registry[reg_key] = {
                                "source_sheet": sheet_name,
                                "source_batch": source_batch,
                                "metric": metric,
                                "value": value,
                                "query_text": query_text,
                            }

                        ai_notes.append(f"{metric}: {apply_note} {notes}".strip())

            # Nigeria special numeric feedback rule for Outlet Stock / Outlet Purchases.
            # This is allowed only by the sheet metric, and still respects Data Queries override.
            # If the text feedback explicitly contradicts the sheet/query metric, the numeric
            # feedback must also be skipped. Skipped means skipped: no value is changed.
            if is_nigeria_outlet_numeric_sheet(sheet_name) and fb_cols.get("numeric_feedback"):
                metric = target_metric_for_numeric_sheet(sheet_name)
                if not metric_allowed_by_query(metric, requested_metrics):
                    skipped_by_metric += 1
                    ai_notes.append(
                        f"Skipped Nigeria numeric {metric} replacement because query/sheet requested {sorted(requested_metrics) or 'no metric'}."
                    )
                elif original_fb and has_metric_contradiction(original_fb, {metric}):
                    skipped_by_metric += 1
                    ai_notes.append(
                        f"Skipped Nigeria numeric {metric} replacement because feedback text explicitly refers to a different metric."
                    )
                else:
                    reg_key = correction_registry_key(q_row_original, q_cols, metric)
                    if reg_key in correction_registry:
                        skipped_by_registry += 1
                        ai_notes.append(
                            f"Skipped Nigeria numeric {metric} replacement because Data Queries already corrected this outlet/SKU/metric."
                        )
                    else:
                        q_row, replace_note = apply_nigeria_numeric_replacement(
                            q_row,
                            sheet_name,
                            numeric_fb,
                            original_fb,
                            current_month,
                        )
                        ai_notes.append(replace_note)

        final_row = pd.Series(index=final_cols, dtype=object)
        final_row[BATCH_COL] = batch_no

        for col in format_cols:
            if col in final_row.index:
                final_row[col] = q_row.get(col, "")

        final_row[FEEDBACK_COL] = " ".join([x for x in refined_feedbacks if x]).strip()
        final_row[ORIGINAL_FEEDBACK_COL] = " | ".join([x for x in original_feedbacks if x]).strip()
        final_row[AI_NOTES_COL] = " | ".join([x for x in ai_notes if x]).strip()

        if final_row[FEEDBACK_COL]:
            rows_with_feedback += 1

        output_rows.append(final_row)

    stats = {
        "Batch": batch_no,
        "Sheet": sheet_name,
        "Query Rows": len(query_df),
        "Matched Query Rows": matched_rows,
        "Unmatched Query Rows": unmatched_rows,
        "Rows With Refined Feedback": rows_with_feedback,
        "Skipped Corrections - Metric Mismatch": skipped_by_metric,
        "Skipped Corrections - Data Queries Override": skipped_by_registry,
    }

    return pd.DataFrame(output_rows, columns=final_cols), stats


def merge_feedback_batches() -> str:
    t0 = time.time()
    print("\n" + "=" * 78)
    print(f"UNIFIED FEEDBACK BATCH MERGER — {PROJECT_NAME}")
    print("=" * 78)

    if PROJECT_NAME not in PROJECT_CONFIGS:
        raise ValueError(f"PROJECT_NAME='{PROJECT_NAME}' is not in PROJECT_CONFIGS")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    feedback_files: Dict[int, str] = {}
    query_files: Dict[int, str] = {}

    for batch in BATCH_NUMBERS:
        try:
            feedback_files[batch] = find_batch_excel_file(FEEDBACK_FOLDER, batch, "feedback")
            print(f"Batch {batch} feedback: {feedback_files[batch]}")
        except Exception as e:
            print(f"Batch {batch} feedback missing/skipped: {e}")

        try:
            query_files[batch] = find_batch_excel_file(QUERY_FOLDER, batch, "query")
            print(f"Batch {batch} query   : {query_files[batch]}")
        except Exception as e:
            print(f"Batch {batch} query missing/skipped: {e}")

    if not query_files:
        raise FileNotFoundError("No query files found for any configured batch.")
    if not feedback_files:
        raise FileNotFoundError("No feedback files found for any configured batch.")

    available_sheets = discover_available_sheets(query_files, feedback_files)
    if not available_sheets:
        raise ValueError("No processable sheets were discovered.")

    print("\nDiscovered sheets:")
    for s in available_sheets:
        print(f"  - {s}")

    ai_cache = load_ai_cache(str(output_dir))
    output_sheets: Dict[str, pd.DataFrame] = {}
    summary_rows: List[Dict[str, Any]] = []

    # Shared correction registry used so Data Queries corrections can override
    # Outlet Stock / Outlet Purchases corrections for the same outlet + SKU + metric.
    correction_registry: Dict[Tuple[str, str, str], Dict[str, Any]] = {}

    for sheet_name in available_sheets:
        print(f"\nProcessing sheet: {sheet_name}")

        format_file, format_batch, format_source_type = choose_format_file(sheet_name, query_files, feedback_files)
        if not format_file:
            print(f"  Skipped '{sheet_name}' because no format source was available.")
            continue

        format_df, _ = read_sheet(format_file, sheet_name)
        format_cols = [str(c).strip() for c in format_df.columns]
        format_identity_cols = detect_identity_columns(format_df)
        final_cols = build_final_columns(format_cols, sheet_name, selected_identity=format_identity_cols)

        all_batches_for_sheet: List[pd.DataFrame] = []
        latest_feedback_context = build_latest_feedback_context_for_sheet(sheet_name, feedback_files)

        for batch in BATCH_NUMBERS:
            q_file = query_files.get(batch)
            fb_file = feedback_files.get(batch)

            if not q_file or not fb_file:
                print(f"  Batch {batch}: skipped, query or feedback file missing.")
                summary_rows.append({
                    "Sheet": sheet_name,
                    "Batch": batch,
                    "Status": "Skipped - query or feedback file missing",
                    "Format Source": f"{format_source_type} Batch {format_batch}",
                })
                continue

            if not workbook_has_sheet(q_file, sheet_name):
                print(f"  Batch {batch}: skipped, query sheet missing.")
                summary_rows.append({
                    "Sheet": sheet_name,
                    "Batch": batch,
                    "Status": "Skipped - query sheet missing",
                    "Format Source": f"{format_source_type} Batch {format_batch}",
                })
                continue

            if not workbook_has_sheet(fb_file, sheet_name):
                print(f"  Batch {batch}: skipped, feedback sheet missing.")
                summary_rows.append({
                    "Sheet": sheet_name,
                    "Batch": batch,
                    "Status": "Skipped - feedback sheet missing",
                    "Format Source": f"{format_source_type} Batch {format_batch}",
                })
                continue

            try:
                batch_df, stats = process_batch_sheet(
                    batch_no=batch,
                    sheet_name=sheet_name,
                    query_file=q_file,
                    feedback_file=fb_file,
                    format_cols=format_cols,
                    final_cols=final_cols,
                    ai_cache=ai_cache,
                    latest_feedback_context=latest_feedback_context,
                    correction_registry=correction_registry,
                )
                all_batches_for_sheet.append(batch_df)
                save_ai_cache(str(output_dir), ai_cache)

                stats["Status"] = "Processed"
                stats["Format Source"] = f"{format_source_type} Batch {format_batch}"
                summary_rows.append(stats)

                print(f"  Batch {batch}: processed {len(batch_df)} rows")
            except Exception as e:
                print(f"  Batch {batch}: ERROR - {e}")
                summary_rows.append({
                    "Sheet": sheet_name,
                    "Batch": batch,
                    "Status": f"Error - {e}",
                    "Format Source": f"{format_source_type} Batch {format_batch}",
                })

        if all_batches_for_sheet:
            combined_sheet_df = pd.concat(all_batches_for_sheet, ignore_index=True)
            before_dedup = len(combined_sheet_df)
            combined_sheet_df = deduplicate_latest_output_rows(combined_sheet_df, sheet_name)
            after_dedup = len(combined_sheet_df)
            if after_dedup != before_dedup:
                print(f"  Latest-batch de-duplication for '{sheet_name}': kept {after_dedup} of {before_dedup} rows")
            output_sheets[sheet_name] = combined_sheet_df
        else:
            print(f"  No rows processed for sheet '{sheet_name}'.")

    if not output_sheets:
        raise ValueError("No output sheets were produced.")

    output_file = str(output_dir / OUTPUT_FILE_NAME)
    summary_df = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        used_sheet_names = set()
        for sheet_name, df in output_sheets.items():
            safe_name = safe_excel_sheet_name(sheet_name, used_sheet_names)

            # Final display cleanup only: remove embedded CR/LF so output cells
            # do not look wrapped/multiline in Excel. This does not change any
            # correction/matching logic above.
            df_to_write = normalize_output_dataframe_text(df)

            df_to_write.to_excel(writer, sheet_name=safe_name, index=False)
            autofit_worksheet(writer, df_to_write, safe_name)
            merge_repeated_outlet_blocks_xlsxwriter(writer, df_to_write, safe_name, sheet_name)

        summary_name = safe_excel_sheet_name("Merge Summary", used_sheet_names)
        summary_to_write = normalize_output_dataframe_text(summary_df)
        summary_to_write.to_excel(writer, sheet_name=summary_name, index=False)
        autofit_worksheet(writer, summary_to_write, summary_name)

    elapsed = time.time() - t0
    print("\n" + "=" * 78)
    print(f"DONE: {output_file}")
    print(f"Sheets written: {len(output_sheets)} + Merge Summary")
    print(f"Elapsed: {elapsed:.1f}s")
    print("=" * 78)

    return output_file


if __name__ == "__main__":
    merge_feedback_batches()

