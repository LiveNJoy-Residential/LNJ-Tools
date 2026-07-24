"""
Shared utility helpers for the Resident Activity Audit Tool.
Standalone — no dependency on audit_bot.py.
"""

import io
import os
import re
import pandas as pd


# ---------------------------------------------------------------------------
# Value cleaners
# ---------------------------------------------------------------------------

def clean_currency(val) -> float:
    """Strip $, commas, spaces → float. Returns 0.0 on failure."""
    if pd.isna(val) or str(val).strip() in ("", "nan", "--"):
        return 0.0
    cleaned = re.sub(r'[$,"\s]', "", str(val))
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def clean_unit(val) -> str:
    """Normalize unit number: strip leading zeros, handle '101 - Name' format."""
    s = str(val).strip() if not pd.isna(val) else ""
    if not s or s.lower() == "nan":
        return "UNKNOWN"
    if " - " in s:
        s = s.split(" - ")[0].strip()
    return s.lstrip("0") or "0"


def parse_date(val):
    """Best-effort date parser; returns pd.Timestamp or None."""
    if pd.isna(val) or str(val).strip() in ("", "nan"):
        return None
    try:
        return pd.to_datetime(val, infer_datetime_format=True)
    except Exception:
        return None


def clean_name(val) -> str:
    """Strip ResMan status markers (* = NTV, ** = MTM) from resident names."""
    return re.sub(r"\*+", "", str(val)).strip()


# ---------------------------------------------------------------------------
# Property name resolution
# ---------------------------------------------------------------------------

def derive_property(filename: str) -> str:
    """Map any ResMan export filename to the standard full property name."""
    CODE_MAP = {
        "CAI": "Crossings at Irving",
        "POT": "Parks on Taylor",
        "HP":  "Highland Park",
        "LP":  "La Prada",
        "VG":  "Village Green",
        "VPA": "Valencia Plaza",
        "VP":  "Valencia Plaza",
        "WST": "Western Station",
    }
    KEYWORD_MAP = {
        "crossing": "Crossings at Irving",
        "irving":   "Crossings at Irving",
        "taylor":   "Parks on Taylor",
        "highland": "Highland Park",
        "prada":    "La Prada",
        "village":  "Village Green",
        "valencia": "Valencia Plaza",
        "western":  "Western Station",
    }
    first_word = filename.split(" ")[0].replace(",", "").upper()
    if first_word in CODE_MAP:
        return CODE_MAP[first_word]
    fname_lower = filename.lower()
    for keyword, prop in KEYWORD_MAP.items():
        if keyword in fname_lower:
            return prop
    return filename.split(" ")[0]


# ---------------------------------------------------------------------------
# File I/O helpers
# ---------------------------------------------------------------------------

def _csv_files(folder: str) -> list:
    if not os.path.exists(folder):
        print(f"  [WARN] Folder missing: {folder}")
        return []
    files = [f for f in os.listdir(folder) if f.lower().endswith(".csv")]
    if not files:
        print(f"  [INFO] No CSVs found in: {folder}")
    return files


def _read_csv(fpath_or_buffer, **kwargs) -> pd.DataFrame:
    """Accept a file path or file-like object. Try multiple encodings."""
    if hasattr(fpath_or_buffer, "read"):
        if hasattr(fpath_or_buffer, "seek"):
            fpath_or_buffer.seek(0)
        raw = fpath_or_buffer.read()
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=enc, **kwargs)
            except UnicodeDecodeError:
                continue
        raise ValueError("Could not decode uploaded file with utf-8-sig, cp1252, or latin-1.")
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return pd.read_csv(fpath_or_buffer, encoding=enc, **kwargs)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Could not decode {fpath_or_buffer}.")
