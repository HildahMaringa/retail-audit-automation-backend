"""
Unified Cooler + POS Queries

Outputs:
- If project has Coolers only:
  1. <input>-Cooler Queries.xlsx
  2. <input>-Combined Queries.xlsx

- If project has Coolers + POS:
  1. <input>-Cooler Queries.xlsx
  2. <input>-POS Queries.xlsx
  3. <input>-Combined Queries.xlsx

Supported projects:
- NG-MRA
- Kenya-MRA
- KO-Tanzania
- KO-Uganda
- TZ-MRA

Cooler behavior:
- All cooler projects now follow the Nigeria MRA style:
  - numeric cooler history sheets
  - Combined sheet
  - current month compared against history
  - last 2 months consistency suppression
  - numeric tolerance of +/- 1
  - none/no/0 treated as no cooler
  - Sum of Coolerbanks + FridgeStockWithoutCooler where available
"""

import os
import re
import difflib
from collections import defaultdict, Counter
from functools import reduce
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


# ═══════════════════════════════════════════════════════════════
# CONFIG — change these each run
# ═══════════════════════════════════════════════════════════════

PROJECT_NAME = "NG-MRA"

INPUT_FILE = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\NG-MRA\NG-MRA-Coolers\Data-Files\2026\May2026\MRA_Nigeria_Cooler_Data_May2026-Batch 3.xlsx"
OUTPUT_DIR = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\NG-MRA\NG-MRA-Coolers\Outputs\2026\May2026"


# ═══════════════════════════════════════════════════════════════
# PROJECT CONFIGS
# ═══════════════════════════════════════════════════════════════

PROJECT_CONFIGS: Dict[str, Dict[str, Any]] = {
    "NG-MRA": {
        "outlet_col": ["outletid", "OutletID", "Outlet ID", "wh_outletid", "projectOutletid"],
        "cooler_metrics": [
            {
                "label": "Number of Branded Coolers",
                "yes": ["370_Are there any branded coolers fridges in the outlet?", "Are there any branded coolers fridges in the outlet?"],
                "count_parts": [
                    "372_How many branded coolers fridges are in the outlet that is Working?",
                    "373_How many branded coolers fridges are in the outlet that is Not Working?",
                ],
            },
            {
                "label": "Number of Unbranded Coolers",
                "yes": ["374_Are there unbranded coolers fridges in the outlet?", "Are there unbranded coolers fridges in the outlet?"],
                "count_parts": [
                    "375_How many unbranded coolers fridges are in the outlet that is Working?",
                    "376_How many unbranded coolers fridges are in the outlet that is Not Working?",
                ],
            },
            {
                "label": "Number of Chest Freezers",
                "yes": ["377_Are there any chest freezers in the outlet?", "Are there any chest freezers in the outlet?"],
                "count": ["378_How many chest freezers are there?"],
            },
        ],
        "cooler_presence_columns": [
            ["370_Are there any branded coolers fridges in the outlet?", "Are there any branded coolers fridges in the outlet?"],
            ["374_Are there unbranded coolers fridges in the outlet?", "Are there unbranded coolers fridges in the outlet?"],
            ["377_Are there any chest freezers in the outlet?", "Are there any chest freezers in the outlet?"],
        ],
        "sum_coolerbanks": ["379_Sum of Coolerbanks", "Sum of Coolerbanks", "coolerbank_sum"],
        "pos_questions": {
            "is there any wall branding other branded signage in this outlet?": "branded signage",
            "are there branded chairs andor tables inside of the outlet?": "branded chairs or tables",
            "are there any branded table runners in this outlet?": "branded table runners",
            "is there a permanent branded shop sign outside of the outlet?": "branded shop sign",
        },
        "pos_followups": {
            "is there any wall branding other branded signage in this outlet?": ("which brands are on the wall branding other branded signage?", "branded signage is missing"),
            "are there branded chairs andor tables inside of the outlet?": ("which brands are on the chairs andor tables?", "branded chairs or tables is missing"),
            "are there any branded table runners in this outlet?": ("which brands are on the table runners?", "branded table runners is missing"),
            "is there a permanent branded shop sign outside of the outlet?": ("which brand is on the shop sign?", "branded shop sign is missing"),
        },
    },

    "Kenya-MRA": {
        "outlet_col": ["outletid", "OutletID", "Outlet ID", "wh_outletid", "projectOutletid"],
        "cooler_metrics": [
            {
                "label": "Number of Unbranded Coolers",
                "yes": ["1003_Are there unbranded coolers fridges in the outlet?", "Are there unbranded coolers fridges in the outlet?"],
                "count": ["1004_How many unbranded coolers fridges are in the outlet?", "How many unbranded coolers fridges are in the outlet?"],
            },
            {
                "label": "Number of Branded Coolers",
                "yes": ["1005_Are there any branded coolers fridges in the outlet?", "Are there any branded coolers fridges in the outlet?"],
                "count": ["1006_How many branded coolers fridges are in the outlet?", "How many branded coolers fridges are in the outlet?"],
            },
        ],
        "cooler_presence_columns": [
            ["1003_Are there unbranded coolers fridges in the outlet?", "Are there unbranded coolers fridges in the outlet?"],
            ["1005_Are there any branded coolers fridges in the outlet?", "Are there any branded coolers fridges in the outlet?"],
        ],
        "sum_coolerbanks": ["Sum of Coolerbanks", "coolerbank_sum"],
        "pos_questions": {
            "is there a permanent branded shop sign outside of the outlet?": "branded shop sign",
            "is there any wall branding other branded signage in this outlet?": "branded signage",
            "are there branded chairs andor tables inside of the outlet?": "branded chairs or tables",
            "are there any abs posters inside of the outlet?": "abs posters",
            "are there any branded table runners in this outlet?": "branded table runners",
        },
        "pos_followups": {},
    },

    "KO-Tanzania": {
        "outlet_col": ["outletid", "OutletID", "Outlet ID"],
        "cooler_metrics": [
            {
                "label": "Number of Owner Coolers",
                "yes": ["are there any owner coolersfridges in the outlet?"],
                "count": ["how many owner coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of Coke Coolers",
                "yes": ["are there any coke branded coolersfridges in the outlet?"],
                "count": ["how many coke branded coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of Pepsi Coolers",
                "yes": ["are there any pepsi branded coolersfridges in the outlet?"],
                "count": ["how many pepsi branded coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of Competitor Coolers",
                "yes": ["are there any other competitor branded coolersfridges in the outlet?"],
                "count": ["how many other competitor branded coolersfridges are in the outlet?"],
            },
        ],
        "cooler_presence_columns": [
            ["are there any coolersfridges in the outlet?"],
            ["are there any owner coolersfridges in the outlet?"],
            ["are there any coke branded coolersfridges in the outlet?"],
            ["are there any pepsi branded coolersfridges in the outlet?"],
            ["are there any other competitor branded coolersfridges in the outlet?"],
        ],
        "sum_coolerbanks": ["Sum of Coolerbanks", "coolerbank_sum"],
        "pos_questions": {},
        "pos_followups": {},
    },

    "KO-Uganda": {
        "outlet_col": ["outletid", "OutletID", "Outlet ID"],
        "cooler_metrics": [
            {
                "label": "Number of Owner Coolers",
                "count": ["how many owner coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of Coke Coolers",
                "count": ["how many coke branded coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of Pepsi Coolers",
                "count": ["how many pepsi branded coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of Competitor Coolers",
                "count": ["how many other competitor branded coolersfridges are in the outlet?"],
            },
        ],
        "cooler_presence_columns": [
            ["Are there any coolersfridges are in the outlet?", "Are there any coolersfridges in the outlet?"],
        ],
        "sum_coolerbanks": ["Sum of Coolerbanks", "coolerbank_sum"],
        "pos_questions": {},
        "pos_followups": {},
    },

    "TZ-MRA": {
        "outlet_col": ["wh_outletid", "outletid", "OutletID", "Outlet ID"],
        "cooler_metrics": [
            {
                "label": "Number of OWNER coolers",
                "yes": ["1179_Are there any OWNER coolersfridges in the outlet?"],
                "count": ["1180_How many OWNER coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of TBL coolers",
                "yes": ["1181_Are there any TBL branded coolersfridges in the outlet?"],
                "count": ["1182_How many TBL branded coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of SBL coolers",
                "yes": ["1188_Are there any SBL branded coolersfridges in the outlet?"],
                "count": ["1189_How many SBL branded coolersfridges are in the outlet?"],
            },
            {
                "label": "Number of Competitor coolers",
                "yes": ["1195_Are there any other Competitor branded coolersfridges in the outlet?"],
                "count": ["1196_How many other Competitor branded coolersfridges are in the outlet?"],
            },
        ],
        "cooler_presence_columns": [
            ["1179_Are there any OWNER coolersfridges in the outlet?"],
            ["1181_Are there any TBL branded coolersfridges in the outlet?"],
            ["1188_Are there any SBL branded coolersfridges in the outlet?"],
            ["1195_Are there any other Competitor branded coolersfridges in the outlet?"],
        ],
        "sum_coolerbanks": ["Sum of Coolerbanks", "coolerbank_sum"],
        "pos_questions": {},
        "pos_followups": {},
    },
}


# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════

def _derive_output_paths(input_file: str, output_dir: str, has_pos: bool) -> Dict[str, str]:
    out_dir = output_dir or os.path.dirname(input_file) or "."
    base = os.path.splitext(os.path.basename(input_file))[0]
    paths = {
        "cooler": os.path.join(out_dir, f"{base}-Cooler Queries.xlsx"),
        "combined": os.path.join(out_dir, f"{base}-Combined Queries.xlsx"),
    }
    if has_pos:
        paths["pos"] = os.path.join(out_dir, f"{base}-POS Queries.xlsx")
    return paths


def _normalize_col_name(name: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def _clean_question_label(label: Any) -> str:
    return re.sub(r"^[\d\-\~\_: ]+", "", str(label)).strip().lower()


def _clean_cooler_label(label: Any) -> str:
    return re.sub(r"^[^a-zA-Z]+", "", str(label)).strip().lower()


def _find_col(columns: List[Any], candidates: List[str], fuzzy: bool = True) -> Optional[str]:
    cols = list(columns)

    for cand in candidates:
        cn = _normalize_col_name(cand)
        for col in cols:
            if _normalize_col_name(col) == cn:
                return col

    for cand in candidates:
        cn = _normalize_col_name(cand)
        for col in cols:
            coln = _normalize_col_name(col)
            if cn and (cn in coln or coln in cn):
                return col

    if fuzzy:
        cleaned_cols = [_clean_cooler_label(c) for c in cols]
        for cand in candidates:
            target = _clean_cooler_label(cand)
            match = difflib.get_close_matches(target, cleaned_cols, n=1, cutoff=0.72)
            if match:
                return cols[cleaned_cols.index(match[0])]

    return None


def _truthy_yes(value: Any) -> bool:
    return str(value).strip().lower() in ["yes", "y", "1", "true"]


def _false_like(value: Any) -> bool:
    s = str(value).strip().lower()
    if s in ["", "nan", "none", "no", "0", "0.0", "false", "n"]:
        return True
    try:
        return float(s) == 0
    except Exception:
        return False


def _to_number_or_none(value: Any) -> Optional[int]:
    s = str(value).strip().lower()
    if _false_like(s) or s in ["yes", "no"]:
        return None
    try:
        f = float(s)
        if pd.isna(f) or f == 0:
            return None
        return int(round(f))
    except Exception:
        return None


def _to_display_value(value: Any) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if s.lower() in ["nan", ""]:
        return ""
    try:
        f = float(s)
        if pd.isna(f):
            return ""
        if f == 0:
            return "none"
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return s


def _detect_outlier_nigeria_style(row: pd.Series, months: List[str]) -> str:
    if len(months) < 3:
        return ""

    recent_months = months[-3:]
    recent_vals = [str(row.get(m, "")).strip().lower() for m in recent_months]
    if any(v in ["", "nan"] for v in recent_vals):
        return ""

    current_raw = str(row.get(months[-1], "")).strip().lower()
    if current_raw in ["", "nan"]:
        return ""

    history_vals = [
        str(row.get(m, "")).strip().lower()
        for m in months[:-1]
        if str(row.get(m, "")).strip().lower() not in ["", "nan"]
    ]
    if not history_vals:
        return ""

    last_2_vals = [
        str(row.get(m, "")).strip().lower()
        for m in months[-2:]
        if str(row.get(m, "")).strip().lower() not in ["", "nan"]
    ]
    if len(last_2_vals) == 2 and len(set(last_2_vals)) == 1:
        return "FALSE"

    majority_raw = Counter(history_vals).most_common(1)[0][0]
    cur_i = _to_number_or_none(current_raw)
    maj_i = _to_number_or_none(majority_raw)

    if maj_i == 1 and _false_like(current_raw):
        return "FALSE"
    if _false_like(majority_raw) and cur_i == 1:
        return "FALSE"
    if cur_i is not None and maj_i is not None and abs(cur_i - maj_i) <= 1:
        return "FALSE"

    return "TRUE" if current_raw != majority_raw else "FALSE"


def _detect_pos_outlier(row: pd.Series, months: List[str]) -> str:
    current_month = months[-1]
    recent_months = months[-3:]

    if any(str(row.get(m, "")).strip().lower() not in ["yes", "no"] for m in recent_months):
        return ""

    current = str(row.get(current_month, "")).strip().lower()
    history = pd.Series([row.get(m) for m in months[:-1]]).dropna().astype(str).str.strip().str.lower()
    history = history[history.isin(["yes", "no"])]

    if history.empty:
        return "FALSE"

    yes_count = (history == "yes").sum()
    no_count = (history == "no").sum()
    majority = "yes" if yes_count > no_count else "no"
    outlier = "TRUE" if current != majority or yes_count == no_count else "FALSE"

    recent_vals = pd.Series([row.get(m) for m in months[-2:]]).dropna().astype(str).str.strip().str.lower()
    recent_vals = recent_vals[recent_vals.isin(["yes", "no"])]
    if len(recent_vals) == 2 and len(set(recent_vals)) == 1:
        return "FALSE"

    return outlier


def _read_workbook(input_file: str) -> Tuple[pd.ExcelFile, List[str]]:
    xlsx = pd.ExcelFile(input_file)
    return xlsx, xlsx.sheet_names


def _load_monthly_dfs(input_file: str, months: List[str]) -> Dict[str, pd.DataFrame]:
    dfs = {}
    for month in months:
        df = pd.read_excel(input_file, sheet_name=month)
        df.columns = df.columns.astype(str).str.strip()
        dfs[month] = df
    return dfs


def _resolve_outlet_col(df: pd.DataFrame, config: Dict[str, Any]) -> str:
    col = _find_col(df.columns, config.get("outlet_col", ["outletid"]))
    if not col:
        raise ValueError("Could not resolve outlet id column for this project.")
    return col


def _build_metadata_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "batch": _find_col(df.columns, ["Batch", "Batch1", "batchid", "Batch No", "Batch Number"]),
        "outlet_id": _find_col(df.columns, ["outletid", "OutletID", "Outlet ID", "projectOutletid", "wh_outletid", "OUTLET ID"]),
        "outlet_name": _find_col(df.columns, ["Outletname", "Outlet Name", "OUTLET NAME", "Outlet"]),
        "region": _find_col(df.columns, ["PRegion", "Region", "GreaterRegion", "Greater Region"]),
        "territory": _find_col(df.columns, ["PTerritory", "Territory"]),
        "district": _find_col(df.columns, ["DistrictName", "District Name", "District"]),
        "trade_channel": _find_col(df.columns, ["TradeChannel", "Trade Channel", "Channel"]),
        "town": _find_col(df.columns, ["Town"]),
        "auditor": _find_col(df.columns, ["Auditor Name", "Auditor", "AUDITOR", "FW Name", "Fieldworker"]),
        "coolerbanks": _find_col(df.columns, ["Sum of Coolerbanks", "coolerbank_sum", "Coolerbanks Sum", "379_Sum of Coolerbanks"]),
        "status": _find_col(df.columns, ["statusflag", "Status", "Status Flag"]),
        "queries": _find_col(df.columns, ["Queries", "Query"]),
    }


def _get_months_4(months: List[str]) -> List[str]:
    return months[-4:] if len(months) >= 4 else months[:]


def _format_combined_sheet(raw_df: pd.DataFrame, months: List[str], module_type: str) -> pd.DataFrame:
    if raw_df is None or raw_df.empty:
        return pd.DataFrame()

    df = raw_df.copy()
    months_4 = _get_months_4(months)
    meta = _build_metadata_columns(df)

    if meta["outlet_id"] is None:
        return df

    out = pd.DataFrame()
    ordered = [
        ("Batch", "batch"),
        ("Outlet ID", "outlet_id"),
        ("Outlet Name", "outlet_name"),
        ("Region", "region"),
        ("Territory", "territory"),
        ("District Name", "district"),
        ("Trade Channel", "trade_channel"),
        ("Town", "town"),
        ("Auditor", "auditor"),
        ("Sum of Coolerbanks", "coolerbanks"),
        ("Status", "status"),
        ("Queries", "queries"),
    ]

    for output_name, key in ordered:
        col = meta.get(key)
        out[output_name] = df[col] if col else ""

    used = set([v for v in meta.values() if v])

    history_cols = []
    for col in df.columns:
        if col in used:
            continue
        col_s = str(col)
        keep = False
        for m in months_4:
            if module_type == "cooler" and f"({m})" in col_s:
                keep = True
            if module_type == "pos" and col_s.endswith(f"_{m}"):
                keep = True
        if "outlier" in col_s.lower():
            keep = True
        if keep:
            history_cols.append(col)

    for col in history_cols:
        out[col] = df[col]

    return out


def _write_df_sheet(writer, df: pd.DataFrame, sheet_name: str):
    safe = str(sheet_name)[:31]
    for ch in ["/", "\\", "?", "*", "[", "]", ":"]:
        safe = safe.replace(ch, " ")
    if df is None or df.empty:
        pd.DataFrame({"Message": [f"No records for {sheet_name}."]}).to_excel(writer, sheet_name=safe, index=False)
    else:
        df.to_excel(writer, sheet_name=safe, index=False)


# ═══════════════════════════════════════════════════════════════
# COOLER ENGINE
# ═══════════════════════════════════════════════════════════════

def _metric_month_value(row: pd.Series, metric: Dict[str, Any], col_map: Dict[str, Optional[str]]) -> str:
    yes_col = col_map.get("yes")
    count_col = col_map.get("count")
    count_parts = col_map.get("count_parts") or []

    if yes_col:
        yes_val = str(row.get(yes_col, "")).strip().lower()
        if _truthy_yes(yes_val):
            if count_parts:
                total = 0
                any_seen = False
                for c in count_parts:
                    v = pd.to_numeric(row.get(c), errors="coerce")
                    if pd.notna(v):
                        total += v
                        any_seen = True
                return str(int(total)) if any_seen else ""
            if count_col:
                return _to_display_value(row.get(count_col))
            return "yes"

        if _false_like(yes_val):
            return "none"

        return ""

    if count_col:
        val = row.get(count_col)
        if _false_like(val):
            return "none"
        return _to_display_value(val)

    return ""


def _resolve_metric_columns(df: pd.DataFrame, metric: Dict[str, Any]) -> Dict[str, Any]:
    resolved = {}
    if metric.get("yes"):
        resolved["yes"] = _find_col(df.columns, metric["yes"])
    if metric.get("count"):
        resolved["count"] = _find_col(df.columns, metric["count"])
    if metric.get("count_parts"):
        resolved["count_parts"] = [_find_col(df.columns, [c]) for c in metric["count_parts"]]
        resolved["count_parts"] = [c for c in resolved["count_parts"] if c]
    return resolved


def _run_fridge_stock_without_cooler(
    current_df: pd.DataFrame,
    config: Dict[str, Any],
    outlet_col: str,
    queries_dict: Dict[str, set],
) -> Tuple[List[Dict[str, Any]], set]:
    mismatch_rows = []
    mismatch_ids = set()

    sum_col = _find_col(current_df.columns, config.get("sum_coolerbanks", ["Sum of Coolerbanks"]))
    if not sum_col:
        return mismatch_rows, mismatch_ids

    presence_groups = config.get("cooler_presence_columns", [])
    resolved_presence = []
    for candidates in presence_groups:
        col = _find_col(current_df.columns, candidates)
        if col:
            resolved_presence.append(col)

    if not resolved_presence:
        return mismatch_rows, mismatch_ids

    for _, row in current_df.iterrows():
        oid = str(row.get(outlet_col, "")).strip()
        if not oid:
            continue

        sum_val = pd.to_numeric(row.get(sum_col), errors="coerce")
        if pd.isna(sum_val) or sum_val <= 0:
            continue

        declared_no_coolers = all(_false_like(row.get(col)) for col in resolved_presence)
        if declared_no_coolers:
            mismatch_ids.add(oid)
            rec = {
                "OutletID": oid,
                "Sum of Coolerbanks": sum_val,
            }
            for col in resolved_presence:
                rec[col] = row.get(col)
            mismatch_rows.append(rec)
            queries_dict[oid].add("FridgeStockWithoutCooler")

    return mismatch_rows, mismatch_ids


def run_cooler_queries(project_name: str, input_file: str, output_file: str) -> Tuple[pd.DataFrame, List[str]]:
    config = PROJECT_CONFIGS[project_name]
    xlsx, months = _read_workbook(input_file)
    monthly_dfs = _load_monthly_dfs(input_file, months)

    outlet_col = _resolve_outlet_col(monthly_dfs[months[-1]], config)
    sheet_data: Dict[str, Dict[str, Dict[str, str]]] = {
        metric["label"]: defaultdict(dict) for metric in config.get("cooler_metrics", [])
    }

    for month in months:
        df = monthly_dfs[month]
        month_outlet_col = _find_col(df.columns, config.get("outlet_col", ["outletid"])) or outlet_col

        for metric in config.get("cooler_metrics", []):
            label = metric["label"]
            resolved_metric = _resolve_metric_columns(df, metric)

            for _, row in df.iterrows():
                oid = str(row.get(month_outlet_col, "")).strip()
                if not oid:
                    continue
                sheet_data[label][oid][month] = _metric_month_value(row, metric, resolved_metric)

    queries_dict: Dict[str, set] = defaultdict(set)

    current_df = monthly_dfs[months[-1]].copy()
    current_outlet_col = _find_col(current_df.columns, config.get("outlet_col", ["outletid"])) or outlet_col
    mismatch_rows, mismatch_ids = _run_fridge_stock_without_cooler(
        current_df, config, current_outlet_col, queries_dict
    )

    combined_raw = pd.DataFrame()

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        all_outlet_ids = set(mismatch_ids)

        for label, data in sheet_data.items():
            out = pd.DataFrame.from_dict(data, orient="index").reset_index().rename(columns={"index": "OutletID"})
            for m in months:
                if m not in out.columns:
                    out[m] = ""
            out = out[["OutletID"] + months]

            outlier_col = f"{label} Outlier"
            out[outlier_col] = out.apply(lambda row: _detect_outlier_nigeria_style(row, months), axis=1)
            out.to_excel(writer, sheet_name=label[:31], index=False)

            flagged = out[out[outlier_col].astype(str).str.upper() == "TRUE"].copy()
            if not flagged.empty:
                for oid in flagged["OutletID"].astype(str):
                    queries_dict[oid].add(label)
                all_outlet_ids.update(flagged["OutletID"].astype(str).tolist())

        if mismatch_rows:
            pd.DataFrame(mismatch_rows).to_excel(writer, sheet_name="Mismatches", index=False)

        if all_outlet_ids:
            last_month_df = monthly_dfs[months[-1]].copy()
            last_outlet_col = _find_col(last_month_df.columns, config.get("outlet_col", ["outletid"])) or outlet_col
            last_month_df["OutletID"] = last_month_df[last_outlet_col].astype(str).str.strip()

            meta_cols = [
                c for c in last_month_df.columns
                if (
                    not re.search(r"cooler|freezer|fridge", c, re.IGNORECASE)
                    or _normalize_col_name(c) in ["sumofcoolerbanks", "coolerbanksum", "379sumofcoolerbanks"]
                )
            ]
            if "OutletID" not in meta_cols:
                meta_cols.append("OutletID")

            metadata = last_month_df[meta_cols].copy()
            metadata = metadata[metadata["OutletID"].astype(str).isin(all_outlet_ids)]

            cooler_frames = []
            for label, data in sheet_data.items():
                out = pd.DataFrame.from_dict(data, orient="index").reset_index().rename(columns={"index": "OutletID"})
                for m in months:
                    if m not in out.columns:
                        out[m] = ""
                out = out[["OutletID"] + months]

                outlier_col = f"{label} Outlier"
                out[outlier_col] = out.apply(lambda row: _detect_outlier_nigeria_style(row, months), axis=1)

                renamed = out.rename(columns={m: f"{label} ({m})" for m in months})
                renamed[f"{label} (Outlier)"] = out[outlier_col]
                renamed.drop(columns=[outlier_col], inplace=True)
                cooler_frames.append(renamed[renamed["OutletID"].astype(str).isin(all_outlet_ids)])

            cooler_merged = reduce(lambda left, right: pd.merge(left, right, on="OutletID", how="outer"), cooler_frames)
            final = pd.merge(metadata, cooler_merged, on="OutletID", how="left")
            final["Queries"] = final["OutletID"].map(lambda x: ", ".join(sorted(queries_dict.get(str(x), []))))
            final = final[final["Queries"].astype(str).str.strip() != ""]
            combined_raw = final.copy()
            final.to_excel(writer, sheet_name="Combined", index=False)
        else:
            pd.DataFrame(columns=["No outliers found"]).to_excel(writer, sheet_name="Combined", index=False)

    return combined_raw, months


# ═══════════════════════════════════════════════════════════════
# POS ENGINE
# ═══════════════════════════════════════════════════════════════

def run_pos_queries(project_name: str, input_file: str, output_file: str) -> Tuple[pd.DataFrame, List[str]]:
    config = PROJECT_CONFIGS[project_name]
    question_sheet_map = config.get("pos_questions", {})
    followup_check_map = config.get("pos_followups", {})

    if not question_sheet_map:
        return pd.DataFrame(), []

    xlsx, months = _read_workbook(input_file)
    question_data = defaultdict(lambda: defaultdict(dict))
    outlier_flags = defaultdict(set)

    for month in months:
        df = pd.read_excel(input_file, sheet_name=month)
        df.columns = df.columns.astype(str).str.strip()
        outlet_col = _find_col(df.columns, config.get("outlet_col", ["outletid"])) or "outletid"
        cleaned_cols = [_clean_question_label(c) for c in df.columns]

        for i, label in enumerate(cleaned_cols):
            if label in question_sheet_map:
                question_col = df.columns[i]
                for _, row in df.iterrows():
                    oid = str(row.get(outlet_col, "")).strip()
                    if not oid:
                        continue
                    val = str(row.get(question_col, "")).strip().lower()
                    question_data[label][oid][month] = val if val in ["yes", "no"] else ""

    combined_dfs = []
    combined_raw = pd.DataFrame()

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        for question_text, outlet_dict in question_data.items():
            df_out = pd.DataFrame.from_dict(outlet_dict, orient="index")
            df_out.index.name = "outletid"
            df_out = df_out.reset_index()

            for m in months:
                if m not in df_out.columns:
                    df_out[m] = ""

            df_out = df_out[["outletid"] + months]
            sheet_name = question_sheet_map[question_text][:31]
            outlier_col = question_sheet_map[question_text] + " Outliers"

            df_out[outlier_col] = df_out.apply(lambda row: _detect_pos_outlier(row, months), axis=1)
            for oid, val in zip(df_out["outletid"], df_out[outlier_col]):
                if val == "TRUE":
                    outlier_flags[str(oid)].add(sheet_name)

            df_out.to_excel(writer, sheet_name=sheet_name, index=False)

            prefixed = df_out.copy()
            prefixed.columns = ["outletid"] + [f"{sheet_name}_{m}" for m in months] + [f"{sheet_name}_Outliers"]
            combined_dfs.append(prefixed)

        combined_final = combined_dfs[0] if combined_dfs else pd.DataFrame(columns=["outletid"])
        for df in combined_dfs[1:]:
            combined_final = combined_final.merge(df, on="outletid", how="outer")

        last_df = pd.read_excel(input_file, sheet_name=months[-1])
        last_df.columns = last_df.columns.astype(str).str.strip()
        outlet_col = _find_col(last_df.columns, config.get("outlet_col", ["outletid"])) or "outletid"
        last_df["outletid"] = last_df[outlet_col].astype(str).str.strip()

        meta_cols = [
            c for c in last_df.columns
            if (
                not re.search(r"cooler|freezer|fridge", c, re.IGNORECASE)
                or _normalize_col_name(c) in ["sumofcoolerbanks", "coolerbanksum", "379sumofcoolerbanks"]
            )
        ]
        metadata = last_df[meta_cols].copy()

        if followup_check_map:
            current_df = last_df.copy()
            cleaned_columns = {_clean_question_label(col): col for col in current_df.columns}

            for yesno_raw, (followup_raw, query_text) in followup_check_map.items():
                yesno_clean = _clean_question_label(yesno_raw)
                followup_clean = _clean_question_label(followup_raw)

                yesno_col = next((v for k, v in cleaned_columns.items() if yesno_clean in k), None)
                followup_col = next((v for k, v in cleaned_columns.items() if followup_clean in k), None)

                if yesno_col and followup_col:
                    for _, row in current_df.iterrows():
                        outlet = str(row.get("outletid", "")).strip()
                        yes_val = str(row.get(yesno_col, "")).strip().lower()
                        followup_val = str(row.get(followup_col, "")).strip()

                        if outlet and yes_val == "yes" and (followup_val == "" or followup_val.lower() in ["none", "nan"]):
                            outlier_flags[outlet].add(query_text)

        flagged_outlets = set(outlier_flags.keys())
        existing_outlets = set(combined_final["outletid"].astype(str)) if "outletid" in combined_final.columns else set()

        new_only = flagged_outlets - existing_outlets
        if new_only:
            additional_rows = last_df[last_df["outletid"].isin(new_only)][["outletid"]].copy()
            combined_final = pd.concat([combined_final, additional_rows], ignore_index=True)

        combined_final = pd.merge(metadata, combined_final, on="outletid", how="right")
        combined_final["Queries"] = combined_final["outletid"].map(lambda x: ", ".join(sorted(outlier_flags.get(str(x), []))))
        combined_final = combined_final[combined_final["Queries"].astype(str).str.strip() != ""]
        combined_raw = combined_final.copy()
        combined_final.to_excel(writer, sheet_name="Combined", index=False)

    return combined_raw, months


# ═══════════════════════════════════════════════════════════════
# MAIN RUNNER
# ═══════════════════════════════════════════════════════════════

def run_unified_cooler_pos_queries(project_name: str, input_file: str, output_dir: str):
    if project_name not in PROJECT_CONFIGS:
        raise ValueError(f"Unknown PROJECT_NAME '{project_name}'. Available: {list(PROJECT_CONFIGS.keys())}")

    config = PROJECT_CONFIGS[project_name]
    has_pos = bool(config.get("pos_questions"))
    paths = _derive_output_paths(input_file, output_dir, has_pos=has_pos)
    os.makedirs(output_dir, exist_ok=True)

    print("\n" + "=" * 70)
    print(f"  UNIFIED COOLER/POS QUERY ENGINE — {project_name}")
    print("=" * 70)
    print(f"Input:  {input_file}")
    print(f"Output: {output_dir}")

    print("\nRunning Cooler Queries...")
    cooler_raw, cooler_months = run_cooler_queries(project_name, input_file, paths["cooler"])
    cooler_sheet = _format_combined_sheet(cooler_raw, cooler_months, "cooler")

    pos_sheet = pd.DataFrame()
    if has_pos:
        print("\nRunning POS Queries...")
        pos_raw, pos_months = run_pos_queries(project_name, input_file, paths["pos"])
        pos_sheet = _format_combined_sheet(pos_raw, pos_months, "pos")
    else:
        print("\nPOS Queries skipped: no POS config for this project.")

    print("\nWriting Combined Queries workbook...")
    with pd.ExcelWriter(paths["combined"], engine="xlsxwriter") as writer:
        _write_df_sheet(writer, cooler_sheet, "Cooler Queries")
        if has_pos:
            _write_df_sheet(writer, pos_sheet, "POS Queries")

    print("\nSaved outputs:")
    print(f"  Cooler Queries:   {paths['cooler']}")
    if has_pos:
        print(f"  POS Queries:      {paths['pos']}")
    print(f"  Combined Queries: {paths['combined']}")


if __name__ == "__main__":
    run_unified_cooler_pos_queries(PROJECT_NAME, INPUT_FILE, OUTPUT_DIR)
