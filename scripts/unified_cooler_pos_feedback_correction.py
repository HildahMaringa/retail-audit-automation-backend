"""
Unified Cooler + POS Feedback Correction Tool

This is a fresh unified correction script for Cooler/POS feedback.

It supports:
1. One feedback column layouts:
   - Feedback
   - Comments
   - Field Comments

2. Multiple feedback columns layouts:
   - Owner cooler feedback
   - Coke coolers feedback
   - Pepsi cooler feedback
   - Competitor cooler feedback
   - repeated "Feedback" columns after POS/cooler history blocks

3. One output workbook:
   - current month data sheet only
   - Correction Summary sheet

4. Output tracking columns:
   - Cooler Feedback
   - Cooler Correction Status
   - POS Feedback
   - POS Correction Status

Important:
- If the feedback file has multiple cooler feedback columns, they are combined into one
  "Cooler Feedback" cell separated by commas.
- If the feedback file has multiple POS feedback columns, they are combined into one
  "POS Feedback" cell separated by commas.
- The correction parser is not restricted to one project layout. It uses the query text,
  feedback text, and nearby feedback-column context to infer what should be corrected.
"""

import os
import re
import time
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


# ═══════════════════════════════════════════════════════════════
# CONFIG — change these each run
# ═══════════════════════════════════════════════════════════════

PROJECT_NAME = "NG-MRA"

DATA_PATH = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\NG-MRA\NG-MRA-Coolers\Data-Files\2026\May2026\MRA_Nigeria_Cooler_Data_May2026-Batch 2.xlsx"
FEEDBACK_PATH = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\NG-MRA\NG-MRA Queries-Feedbacks\2026\May2026\NG-MRA-Feedbacks-May2026-Batch 2.xlsx"
OUTPUT_DIR = r"C:\Users\ID0373122\OneDrive - Kantar\Desktop\ALL PROJECTS\NG-MRA\NG-MRA Queries-Feedbacks\2026\May2026-1"

COOLER_FEEDBACK_COL = "Cooler Feedback"
COOLER_STATUS_COL = "Cooler Correction Status"
POS_FEEDBACK_COL = "POS Feedback"
POS_STATUS_COL = "POS Correction Status"

PREFERRED_FEEDBACK_SHEETS = [
    "Cooler Queries",
    "POS Queries"
]


# ═══════════════════════════════════════════════════════════════
# PROJECT CONFIGS
# ═══════════════════════════════════════════════════════════════

PROJECT_CONFIGS: Dict[str, Dict[str, Any]] = {
    "NG-MRA": {
        "outlet_cols": ["outletid", "OutletID", "Outlet ID", "wh_outletid", "projectOutletid"],
        "coolers": {
            "branded": {
                "tokens": ["branded cooler", "branded coolers", "branded coolers fridges"],
                "exclude_tokens": ["unbranded"],
                "yes_cols": ["370_Are there any branded coolers fridges in the outlet?", "Are there any branded coolers fridges in the outlet?"],
                "count_cols": [],
                "working_cols": ["372_How many branded coolers fridges are in the outlet that is Working?"],
                "not_working_cols": ["373_How many branded coolers fridges are in the outlet that is Not Working?"],
            },
            "unbranded": {
                "tokens": ["unbranded cooler", "unbranded coolers", "unbranded coolers fridges"],
                "exclude_tokens": [],
                "yes_cols": ["374_Are there unbranded coolers fridges in the outlet?", "Are there unbranded coolers fridges in the outlet?"],
                "count_cols": [],
                "working_cols": ["375_How many unbranded coolers fridges are in the outlet that is Working?"],
                "not_working_cols": ["376_How many unbranded coolers fridges are in the outlet that is Not Working?"],
            },
            "chest": {
                "tokens": ["chest freezer", "chest freezers"],
                "exclude_tokens": [],
                "yes_cols": ["377_Are there any chest freezers in the outlet?", "Are there any chest freezers in the outlet?"],
                "count_cols": ["378_How many chest freezers are there?"],
                "working_cols": [],
                "not_working_cols": [],
            },
        },
        "pos": {
            "branded signage": {
                "tokens": ["wall branding", "branded signage", "signage"],
                "yes_cols": ["364_Is there any wall branding other branded signage in this outlet?", "wall branding", "branded signage"],
            },
            "branded chairs or tables": {
                "tokens": ["branded chairs", "branded tables", "chairs", "tables"],
                "yes_cols": ["366_Are there branded chairs andor tables inside of the outlet?", "branded chairs", "branded tables"],
            },
            "branded table runners": {
                "tokens": ["table runner", "table runners"],
                "yes_cols": ["368_Are there any branded table runners in this outlet?", "table runners"],
            },
            "branded shop sign": {
                "tokens": ["shop sign", "branded shop sign"],
                "yes_cols": ["395_Is there a permanent branded shop sign outside of the outlet?", "shop sign"],
            },
        },
    },

    "Kenya-MRA": {
        "outlet_cols": ["outletid", "OutletID", "Outlet ID", "wh_outletid", "projectOutletid"],
        "coolers": {
            "branded": {
                "tokens": ["branded cooler", "branded coolers"],
                "exclude_tokens": ["unbranded"],
                "yes_cols": ["1005_Are there any branded coolers fridges in the outlet?", "Are there any branded coolers fridges in the outlet?"],
                "count_cols": ["1006_How many branded coolers fridges are in the outlet?", "How many branded coolers fridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "unbranded": {
                "tokens": ["unbranded cooler", "unbranded coolers"],
                "exclude_tokens": [],
                "yes_cols": ["1003_Are there unbranded coolers fridges in the outlet?", "Are there unbranded coolers fridges in the outlet?"],
                "count_cols": ["1004_How many unbranded coolers fridges are in the outlet?", "How many unbranded coolers fridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
        },
        "pos": {
            "branded shop sign": {
                "tokens": ["shop sign", "branded shop sign"],
                "yes_cols": ["1001_Is there a permanent branded shop sign outside of the outlet?", "branded shop sign"],
            },
            "branded signage": {
                "tokens": ["wall branding", "branded signage", "signage"],
                "yes_cols": ["1014_Is there any wall branding/ other branded signage in this outlet?", "wall branding", "branded signage"],
            },
            "branded chairs or tables": {
                "tokens": ["branded chairs", "branded tables", "chairs", "tables"],
                "yes_cols": ["1008_Are there branded chairs and/or tables inside of the outlet?", "branded chairs", "branded tables"],
            },
            "abs posters": {
                "tokens": ["abs poster", "abs posters"],
                "yes_cols": ["1010_Are there any ABS posters inside of the outlet?", "abs posters"],
            },
            "branded table runners": {
                "tokens": ["table runner", "table runners"],
                "yes_cols": ["1012_Are there any branded table runners in this outlet?", "table runners"],
            },
        },
    },

    "KO-Uganda": {
        "outlet_cols": ["outletid", "OutletID", "Outlet ID"],
        "coolers": {
            "owner": {
                "tokens": ["owner cooler", "owner coolers", "own cooler", "own coolers"],
                "exclude_tokens": [],
                "yes_cols": ["Are there any coolersfridges are in the outlet?", "Are there any coolersfridges in the outlet?"],
                "count_cols": ["Number of Owner Coolers", "how many owner coolersfridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "coke": {
                "tokens": ["coke cooler", "coke coolers"],
                "exclude_tokens": [],
                "yes_cols": ["Are there any coolersfridges are in the outlet?", "Are there any coolersfridges in the outlet?"],
                "count_cols": ["Number of Coke Coolers", "how many coke branded coolersfridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "pepsi": {
                "tokens": ["pepsi cooler", "pepsi coolers", "pespi cooler", "pespi coolers"],
                "exclude_tokens": [],
                "yes_cols": ["Are there any coolersfridges are in the outlet?", "Are there any coolersfridges in the outlet?"],
                "count_cols": ["Number of Pepsi Coolers", "how many pepsi branded coolersfridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "competitor": {
                "tokens": ["competitor cooler", "competitor coolers"],
                "exclude_tokens": [],
                "yes_cols": ["Are there any coolersfridges are in the outlet?", "Are there any coolersfridges in the outlet?"],
                "count_cols": ["Number of Competitor Coolers", "how many other competitor branded coolersfridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
        },
        "pos": {},
    },

    "KO-Tanzania": {
        "outlet_cols": ["outletid", "OutletID", "Outlet ID", "wh_outletid", "projectOutletid"],
        "coolers": {
            "owner": {
                "tokens": ["owner cooler", "owner coolers"],
                "exclude_tokens": [],
                "yes_cols": ["Owner Coolers", "Are there any owner coolersfridges in the outlet?"],
                "count_cols": ["Number of Owner Coolers", "How many owner coolersfridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "coke": {
                "tokens": ["coke cooler", "coke coolers"],
                "exclude_tokens": [],
                "yes_cols": ["Coke Coolers", "Are there any coke branded coolersfridges in the outlet?"],
                "count_cols": ["Number of Coke Coolers", "How many coke branded coolersfridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "pepsi": {
                "tokens": ["pepsi cooler", "pepsi coolers", "pespi cooler", "pespi coolers"],
                "exclude_tokens": [],
                "yes_cols": ["Pepsi Coolers", "Are there any pepsi branded coolersfridges in the outlet?"],
                "count_cols": ["Number of Pepsi Coolers", "How many pepsi branded coolersfridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "competitor": {
                "tokens": ["competitor cooler", "competitor coolers"],
                "exclude_tokens": [],
                "yes_cols": ["Competitor Coolers", "Are there any other competitor branded coolersfridges in the outlet?"],
                "count_cols": ["Number of Competitor Coolers", "How many other competitor branded coolersfridges are in the outlet?"],
                "working_cols": [],
                "not_working_cols": [],
            },
        },
        "pos": {},
    },

    "TZ-MRA": {
        "outlet_cols": ["wh_outletid", "outletid", "OutletID", "Outlet ID"],
        "coolers": {
            "owner": {
                "tokens": ["owner cooler", "owner coolers"],
                "exclude_tokens": [],
                "yes_cols": ["1179_Are there any OWNER coolersfridges in the outlet?", "Owner Coolersfridges"],
                "count_cols": ["1180_How many OWNER coolersfridges are in the outlet?", "Number Of Owner Coolersfridges"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "tbl": {
                "tokens": ["tbl cooler", "tbl coolers"],
                "exclude_tokens": [],
                "yes_cols": ["1181_Are there any TBL branded coolersfridges in the outlet?", "Tbl Branded Coolersfridges"],
                "count_cols": ["1182_How many TBL branded coolersfridges are in the outlet?", "Number Of Tbl Branded Coolersfridges"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "sbl": {
                "tokens": ["sbl cooler", "sbl coolers"],
                "exclude_tokens": [],
                "yes_cols": ["1188_Are there any SBL branded coolersfridges in the outlet?", "Sbl Branded Coolersfridges"],
                "count_cols": ["1189_How many SBL branded coolersfridges are in the outlet?", "Number Of Sbl Branded Coolersfridges"],
                "working_cols": [],
                "not_working_cols": [],
            },
            "competitor": {
                "tokens": ["competitor cooler", "competitor coolers"],
                "exclude_tokens": [],
                "yes_cols": ["1195_Are there any other Competitor branded coolersfridges in the outlet?", "Competitor Branded Coolersfridges"],
                "count_cols": ["1196_How many other Competitor branded coolersfridges are in the outlet?", "Number Of Competitor Branded Coolersfridges"],
                "working_cols": [],
                "not_working_cols": [],
            },
        },
        "pos": {},
    },
}


# ═══════════════════════════════════════════════════════════════
# BASIC HELPERS
# ═══════════════════════════════════════════════════════════════

def normalize(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def resolve_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if not candidates:
        return None

    cols = list(df.columns)
    norm_map = {normalize(c): c for c in cols}

    for cand in candidates:
        if cand in df.columns:
            return cand
        cn = normalize(cand)
        if cn in norm_map:
            return norm_map[cn]

    for cand in candidates:
        cn = normalize(cand)
        for col in cols:
            n = normalize(col)
            if cn and (cn in n or n in cn):
                return col

    return None


def load_workbook_sheets(path: str) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    xl = pd.ExcelFile(path)
    sheets = {}
    for sheet in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]
        sheets[sheet] = df
    return sheets, xl.sheet_names


def read_feedback_sheet(path: str) -> Tuple[pd.DataFrame, str]:
    xl = pd.ExcelFile(path)
    selected = xl.sheet_names[0]

    for pref in PREFERRED_FEEDBACK_SHEETS:
        for sheet in xl.sheet_names:
            if normalize(sheet) == normalize(pref):
                selected = sheet
                break

    df = pd.read_excel(path, sheet_name=selected)
    df.columns = [str(c).strip() for c in df.columns]
    return df, selected


def ensure_tracking_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in [COOLER_FEEDBACK_COL, COOLER_STATUS_COL, POS_FEEDBACK_COL, POS_STATUS_COL]:
        if col not in df.columns:
            df[col] = ""
    return df


def append_comma(existing: Any, new_text: str) -> str:
    old = clean_text(existing)
    new = clean_text(new_text)
    if not new:
        return old
    if not old:
        return new

    old_parts = [p.strip() for p in old.split(",") if p.strip()]
    if new in old_parts or new in old:
        return old
    return old + ", " + new


def make_outlet_map(df: pd.DataFrame, outlet_col: str) -> Dict[str, List[int]]:
    out = {}
    for idx, row in df.iterrows():
        outlet = clean_text(row.get(outlet_col))
        if outlet:
            out.setdefault(outlet, []).append(idx)
    return out


def extract_first_number(text: str) -> Optional[int]:
    m = re.search(r"(-?\d+(?:\.\d+)?)", str(text).replace(",", ""))
    if not m:
        return None
    try:
        return int(round(float(m.group(1))))
    except Exception:
        return None


def is_negative_feedback(text: str) -> bool:
    t = text.lower()
    return (
        bool(re.search(r"\b(no|none|zero)\b", t))
        or any(p in t for p in [
            "removed",
            "disposed",
            "taken away",
            "sold off",
            "worn out",
            "not available",
            "not there",
            "has no",
            "have no",
            "no longer",
        ])
    )


def infer_yes_no(text: str) -> Optional[str]:
    t = text.lower()

    if is_negative_feedback(t):
        return "No"

    positive_phrases = [
        "available",
        "introduced",
        "mounted",
        "acquired",
        "has",
        "have",
        "there is",
        "there are",
        "present",
        "branded with",
        "put",
        "added",
        "rectified",
    ]
    if any(p in t for p in positive_phrases):
        return "Yes"

    if re.search(r"\byes\b", t):
        return "Yes"
    if re.search(r"\bno\b", t):
        return "No"

    return None


# ═══════════════════════════════════════════════════════════════
# FEEDBACK DISCOVERY
# ═══════════════════════════════════════════════════════════════

def is_feedback_like_col(col: Any) -> bool:
    n = normalize(col)
    return (
        "feedback" in n
        or "comment" in n
        or "comments" in n
        or "fieldcomment" in n
        or "fieldcomments" in n
    )


def discover_feedback_columns(feedback_df: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    Finds all feedback/comment columns.
    For repeated feedback columns, nearby left columns are used as context.
    """
    cols = list(feedback_df.columns)
    discovered = []

    for idx, col in enumerate(cols):
        if not is_feedback_like_col(col):
            continue

        # The question block usually sits to the left of the feedback/comment column.
        left_context = cols[max(0, idx - 8):idx]
        right_context = cols[idx + 1:min(len(cols), idx + 3)]
        context = " ".join([str(x) for x in left_context + [col] + right_context])

        discovered.append({
            "col": col,
            "idx": idx,
            "context": context,
        })

    return discovered


def row_queries(row: pd.Series, feedback_df: pd.DataFrame) -> str:
    q_col = resolve_col(feedback_df, ["Queries", "Query"])
    return clean_text(row.get(q_col)) if q_col else ""


# ═══════════════════════════════════════════════════════════════
# ACTION PARSING
# ═══════════════════════════════════════════════════════════════

def token_allowed(text: str, token: str, exclude_tokens: List[str]) -> bool:
    t = text.lower()
    tok = token.lower()

    if tok not in t:
        return False

    for ex in exclude_tokens:
        if ex.lower() in t:
            return False

    return True


def detect_cooler_type(text: str, config: Dict[str, Any]) -> Optional[str]:
    t = text.lower()
    for cooler_type, c in config.get("coolers", {}).items():
        excludes = c.get("exclude_tokens", [])
        for token in c.get("tokens", []):
            if token_allowed(t, token, excludes):
                return cooler_type
    return None


def detect_pos_type(text: str, config: Dict[str, Any]) -> Optional[str]:
    t = text.lower()
    for pos_type, p in config.get("pos", {}).items():
        for token in p.get("tokens", []):
            if token.lower() in t:
                return pos_type
    return None


def parse_cooler_actions(query_text: str, feedback_text: str, context_text: str, config: Dict[str, Any]) -> List[Dict[str, Any]]:
    joined = f"{query_text} {context_text} {feedback_text}".lower()
    fb = feedback_text.lower()
    actions = []

    for cooler_type, c in config.get("coolers", {}).items():
        token_hit = any(
            token_allowed(joined, token, c.get("exclude_tokens", []))
            for token in c.get("tokens", [])
        )
        if not token_hit:
            continue

        if is_negative_feedback(fb):
            value = 0
        else:
            value = extract_first_number(feedback_text)

        if value is not None:
            actions.append({
                "module": "cooler",
                "type": cooler_type,
                "value": value,
            })

    # Fallback: query/context identifies item, feedback is just a number.
    if not actions:
        cooler_type = detect_cooler_type(f"{query_text} {context_text}", config)
        value = extract_first_number(feedback_text)
        if cooler_type and value is not None:
            actions.append({
                "module": "cooler",
                "type": cooler_type,
                "value": value,
            })

    return actions


def parse_pos_actions(query_text: str, feedback_text: str, context_text: str, config: Dict[str, Any]) -> List[Dict[str, Any]]:
    joined = f"{query_text} {context_text} {feedback_text}".lower()
    actions = []

    value = infer_yes_no(feedback_text)
    if value is None:
        return actions

    for pos_type, p in config.get("pos", {}).items():
        if any(token.lower() in joined for token in p.get("tokens", [])):
            actions.append({
                "module": "pos",
                "type": pos_type,
                "value": value,
            })

    return actions


# ═══════════════════════════════════════════════════════════════
# COLUMN RESOLUTION + APPLY
# ═══════════════════════════════════════════════════════════════

def resolve_data_columns(data_df: pd.DataFrame, config: Dict[str, Any]) -> Dict[str, Any]:
    resolved = {
        "outlet": resolve_col(data_df, config.get("outlet_cols", ["outletid"])),
        "coolers": {},
        "pos": {},
    }

    for cooler_type, c in config.get("coolers", {}).items():
        resolved["coolers"][cooler_type] = {
            "yes": resolve_col(data_df, c.get("yes_cols", [])),
            "count": resolve_col(data_df, c.get("count_cols", [])),
            "working": resolve_col(data_df, c.get("working_cols", [])),
            "not_working": resolve_col(data_df, c.get("not_working_cols", [])),
        }

    for pos_type, p in config.get("pos", {}).items():
        resolved["pos"][pos_type] = {
            "yes": resolve_col(data_df, p.get("yes_cols", [])),
        }

    return resolved


def apply_cooler_action(df: pd.DataFrame, row_idx: int, action: Dict[str, Any], resolved: Dict[str, Any]) -> bool:
    ctype = action["type"]
    value = action["value"]

    cols = resolved["coolers"].get(ctype, {})
    yes_col = cols.get("yes")
    count_col = cols.get("count")
    working_col = cols.get("working")
    not_working_col = cols.get("not_working")

    changed = False

    if yes_col and yes_col in df.columns:
        df.at[row_idx, yes_col] = "Yes" if value > 0 else "No"
        changed = True

    # For NG-MRA style working/not-working split, put corrected count in working
    # and set not-working to 0.
    if working_col and working_col in df.columns:
        df.at[row_idx, working_col] = value
        changed = True

    if not_working_col and not_working_col in df.columns:
        df.at[row_idx, not_working_col] = 0
        changed = True

    if count_col and count_col in df.columns:
        df.at[row_idx, count_col] = value
        changed = True

    return changed


def apply_pos_action(df: pd.DataFrame, row_idx: int, action: Dict[str, Any], resolved: Dict[str, Any]) -> bool:
    ptype = action["type"]
    value = action["value"]

    cols = resolved["pos"].get(ptype, {})
    yes_col = cols.get("yes")

    if yes_col and yes_col in df.columns:
        df.at[row_idx, yes_col] = value
        return True

    return False


# ═══════════════════════════════════════════════════════════════
# MAIN ENGINE
# ═══════════════════════════════════════════════════════════════

def apply_unified_feedback_corrections(
    project_name: str,
    data_path: str,
    feedback_path: str,
    output_dir: str,
) -> str:
    if project_name not in PROJECT_CONFIGS:
        raise ValueError(f"Unknown PROJECT_NAME: {project_name}. Available: {list(PROJECT_CONFIGS.keys())}")

    start = time.time()
    config = PROJECT_CONFIGS[project_name]

    print("\n" + "=" * 70)
    print(f"UNIFIED COOLER/POS FEEDBACK CORRECTION — {project_name}")
    print("=" * 70)

    data_sheets, sheet_names = load_workbook_sheets(data_path)
    current_sheet = sheet_names[-1]
    df_current = ensure_tracking_cols(data_sheets[current_sheet])

    feedback_df, feedback_sheet = read_feedback_sheet(feedback_path)

    resolved = resolve_data_columns(df_current, config)
    if not resolved["outlet"]:
        raise ValueError("Could not resolve outlet column in data file.")

    fb_outlet_col = resolve_col(feedback_df, config.get("outlet_cols", ["outletid"]))
    if not fb_outlet_col:
        raise ValueError("Could not resolve outlet column in feedback file.")

    feedback_cols = discover_feedback_columns(feedback_df)
    if not feedback_cols:
        raise ValueError("No Feedback / Comment / Field Comments columns found in feedback file.")

    row_map = make_outlet_map(df_current, resolved["outlet"])

    print(f"Data current sheet: {current_sheet}")
    print(f"Feedback sheet used: {feedback_sheet}")
    print(f"Feedback/comment columns found: {[x['col'] for x in feedback_cols]}")

    summary_rows = []
    corrected_rows = 0
    not_corrected_rows = 0

    for fb_idx, fb_row in feedback_df.iterrows():
        outlet = clean_text(fb_row.get(fb_outlet_col))
        if not outlet:
            continue

        target_rows = row_map.get(outlet, [])
        if not target_rows:
            continue

        query_text = row_queries(fb_row, feedback_df)

        cooler_feedbacks = []
        pos_feedbacks = []
        actions = []

        for meta in feedback_cols:
            fb_col = meta["col"]
            feedback_text = clean_text(fb_row.get(fb_col))
            if not feedback_text:
                continue

            context_text = meta["context"]

            cooler_actions = parse_cooler_actions(query_text, feedback_text, context_text, config)
            pos_actions = parse_pos_actions(query_text, feedback_text, context_text, config)

            if cooler_actions:
                cooler_feedbacks.append(feedback_text)
                actions.extend(cooler_actions)

            if pos_actions:
                pos_feedbacks.append(feedback_text)
                actions.extend(pos_actions)

        # If feedback exists but no action was understood, track it as Not Corrected
        if not actions:
            raw_feedbacks = [
                clean_text(fb_row.get(x["col"]))
                for x in feedback_cols
                if clean_text(fb_row.get(x["col"]))
            ]

            if raw_feedbacks:
                combined_feedback = ", ".join(raw_feedbacks)
                likely_pos = detect_pos_type(f"{query_text} {combined_feedback}", config)

                for r in target_rows:
                    if likely_pos:
                        df_current.at[r, POS_FEEDBACK_COL] = append_comma(df_current.at[r, POS_FEEDBACK_COL], combined_feedback)
                        df_current.at[r, POS_STATUS_COL] = append_comma(df_current.at[r, POS_STATUS_COL], "Not Corrected")
                    else:
                        df_current.at[r, COOLER_FEEDBACK_COL] = append_comma(df_current.at[r, COOLER_FEEDBACK_COL], combined_feedback)
                        df_current.at[r, COOLER_STATUS_COL] = append_comma(df_current.at[r, COOLER_STATUS_COL], "Not Corrected")

                not_corrected_rows += 1
                summary_rows.append({
                    "Feedback Row": fb_idx + 2,
                    "Outlet": outlet,
                    "Queries": query_text,
                    "Feedback": combined_feedback,
                    "Action": "",
                    "Status": "Not Corrected",
                    "Reason": "Feedback found, but no clear correction action parsed",
                })

            continue

        row_applied = False

        for r in target_rows:
            cooler_applied = False
            pos_applied = False

            for action in actions:
                if action["module"] == "cooler":
                    ok = apply_cooler_action(df_current, r, action, resolved)
                    cooler_applied = cooler_applied or ok
                else:
                    ok = apply_pos_action(df_current, r, action, resolved)
                    pos_applied = pos_applied or ok

                row_applied = row_applied or ok

                summary_rows.append({
                    "Feedback Row": fb_idx + 2,
                    "Outlet": outlet,
                    "Queries": query_text,
                    "Feedback": ", ".join(cooler_feedbacks + pos_feedbacks),
                    "Action": f"{action['module']}:{action['type']}={action['value']}",
                    "Status": "Corrected" if ok else "Not Corrected",
                    "Reason": "Applied parsed feedback" if ok else "Could not resolve target data column",
                })

            if cooler_feedbacks:
                df_current.at[r, COOLER_FEEDBACK_COL] = append_comma(
                    df_current.at[r, COOLER_FEEDBACK_COL],
                    ", ".join(cooler_feedbacks),
                )
                df_current.at[r, COOLER_STATUS_COL] = append_comma(
                    df_current.at[r, COOLER_STATUS_COL],
                    "Corrected" if cooler_applied else "Not Corrected",
                )

            if pos_feedbacks:
                df_current.at[r, POS_FEEDBACK_COL] = append_comma(
                    df_current.at[r, POS_FEEDBACK_COL],
                    ", ".join(pos_feedbacks),
                )
                df_current.at[r, POS_STATUS_COL] = append_comma(
                    df_current.at[r, POS_STATUS_COL],
                    "Corrected" if pos_applied else "Not Corrected",
                )

        if row_applied:
            corrected_rows += 1
        else:
            not_corrected_rows += 1

    os.makedirs(output_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(data_path))[0]
    output_file = os.path.join(output_dir, f"{base}-Unified-Cooler-POS-Feedback-Corrected.xlsx")

    summary_df = pd.DataFrame(summary_rows)
    if summary_df.empty:
        summary_df = pd.DataFrame({"Message": ["No feedback corrections processed."]})

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        df_current.to_excel(writer, sheet_name=str(current_sheet)[:31], index=False)
        summary_df.to_excel(writer, sheet_name="Correction Summary", index=False)

        for worksheet in writer.sheets.values():
            worksheet.freeze_panes(1, 0)
            worksheet.set_column(0, 200, 18)

    elapsed = time.time() - start

    print("\nChanged / tracked sheet:")
    print(f"  - {current_sheet}")
    print(f"\nCorrected feedback records: {corrected_rows}")
    print(f"Not corrected / needs review: {not_corrected_rows}")
    print(f"Runtime: {elapsed:.1f}s")
    print(f"Saved corrected file:\n  {output_file}")

    return output_file


if __name__ == "__main__":
    apply_unified_feedback_corrections(
        PROJECT_NAME,
        DATA_PATH,
        FEEDBACK_PATH,
        OUTPUT_DIR,
    )
