"""
utils.py
========
Shared utility functions used across all three pipeline modules.
This module must be imported first by all other modules to avoid code duplication.

Responsibilities:
    - UTF-8 console configuration (Windows-safe)
    - ID sanitization (prevent scientific notation / decimal corruption)
    - ASIN extraction via regex
    - Keyword / expression text cleaning
    - Raw BulkSheet Excel loader (dtype=str enforcement)
    - Safe SKU → filename string conversion
"""

import os
import re
import sys
import glob

import pandas as pd

# ── Console Setup ─────────────────────────────────────────────────────────────
# Reconfigure stdout/stderr to UTF-8 so emoji and Vietnamese characters display
# correctly in Windows terminals (PowerShell, cmd, Windows Terminal).
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


# ── Compiled Regexes ──────────────────────────────────────────────────────────
# Matches a standard Amazon ASIN: letter B followed by 9 uppercase alphanumeric chars.
# Case-insensitive so it catches lowercase input as well.
_ASIN_RE = re.compile(r"B[A-Z0-9]{9}", re.IGNORECASE)

# Matches one or more consecutive underscores (used for cleanup after concatenation).
_MULTI_UNDERSCORE_RE = re.compile(r"_+")


# ── ID / Value Cleaning ───────────────────────────────────────────────────────

def clean_id(val) -> str:
    """
    Sanitize a numeric-looking ID field (Campaign ID, Portfolio ID, etc.) that
    Amazon's Excel export often corrupts into scientific notation (e.g. '1.23e+17')
    or decimal string (e.g. '123456789012345.0').

    Always returns a plain integer string, or '' for empty/NaN values.

    Args:
        val: Raw cell value (could be float, int, str, NaN).

    Returns:
        A clean string representation of the ID, or '' if empty/invalid.
    """
    # Handle None / NaN gracefully
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return ""

    # Strip trailing '.0' (pandas float-to-str artifact)
    if s.endswith(".0"):
        s = s[:-2]

    # Convert scientific notation to integer string
    if "e+" in s.lower() or "e-" in s.lower():
        try:
            s = str(int(float(s)))
        except ValueError:
            pass  # leave as-is if conversion fails

    return s


def safe_str(val) -> str:
    """
    Return a clean string from any value, stripping leading/trailing whitespace.
    Returns '' for None / NaN.

    Args:
        val: Any raw cell value.

    Returns:
        Stripped string or ''.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


# ── ASIN Extraction ───────────────────────────────────────────────────────────

def extract_asin(text: str) -> str | None:
    """
    Find the first Amazon ASIN in a string using the standard B + 9-char pattern.

    Args:
        text: Any string that might contain an ASIN (SKU, campaign name, etc.)

    Returns:
        The matched ASIN in UPPERCASE, or None if not found.
    """
    if not text or safe_str(text) == "":
        return None
    match = _ASIN_RE.search(str(text).upper())
    return match.group(0) if match else None


# ── Text Cleaning ─────────────────────────────────────────────────────────────

def clean_keyword(kw: str) -> str:
    """
    Sanitize a keyword string for use inside a campaign name:
    - Strip special characters (keeping letters, digits, spaces, hyphens)
    - Collapse multiple spaces → single space
    - Lowercase the result

    Args:
        kw: Raw keyword or targeting expression string.

    Returns:
        Clean lowercase string safe for filename / campaign name embedding.
    """
    kw = re.sub(r"[^a-zA-Z0-9 \-]", "", str(kw))
    kw = re.sub(r"\s+", " ", kw).strip()
    return kw.lower()


def no_double_underscore(name: str) -> str:
    """
    Collapse consecutive underscores into a single one and strip leading/trailing
    underscores. Prevents names like 'ASIN__SP00___KW_Nguyen'.

    Args:
        name: Campaign name string that may have redundant underscores.

    Returns:
        Cleaned name with only single underscores.
    """
    return _MULTI_UNDERSCORE_RE.sub("_", name).strip("_")


def safe_filename(sku: str) -> str:
    """
    Convert a SKU string into a filesystem-safe filename component by replacing
    any character that is NOT alphanumeric, hyphen, or dot with an underscore.

    Args:
        sku: The raw SKU string.

    Returns:
        A safe string suitable for use in a filename.
    """
    return re.sub(r"[^\w\-\.]", "_", sku)


# ── BulkSheet Loader ──────────────────────────────────────────────────────────

# Columns that must be treated as strings to prevent Amazon ID corruption.
ID_COLUMNS = [
    "Campaign ID",
    "Portfolio ID",
    "Ad Group ID",
    "Ad ID",
    "Keyword ID",
    "Product Targeting ID",
]


def load_bulksheet(path: str) -> pd.DataFrame:
    """
    Load an Amazon BulkSheet Excel file into a clean Pandas DataFrame.

    Key safety measures:
    1. dtype=str enforced globally to prevent ANY numeric conversion.
    2. Targets the 'Sponsored Products Campaigns' sheet by name, falling back
       to the first sheet if the named sheet is not found.
    3. Column names are stripped of leading/trailing whitespace.
    4. All ID columns are explicitly re-cleaned via clean_id() as a second-pass
       safeguard (handles cases where openpyxl still reads floats before str cast).

    Args:
        path: Absolute path to the BulkSheetExport_*.xlsx file.

    Returns:
        A clean DataFrame with all values as strings and safe IDs.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file cannot be parsed as an Excel workbook.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"[ERROR] File not found: {path}")

    print(f"[LOADER] Opening: {os.path.basename(path)}")

    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
    except Exception as exc:
        raise ValueError(f"[ERROR] Cannot open Excel file: {exc}") from exc

    # Prefer the exact Amazon sheet name; fall back gracefully
    sheet_name = None
    for s in xl.sheet_names:
        if "sponsored product" in s.lower():
            sheet_name = s
            break

    if sheet_name is None:
        sheet_name = xl.sheet_names[0]
        print(
            f"[LOADER] WARNING – 'Sponsored Products Campaigns' sheet not found. "
            f"Using first sheet: '{sheet_name}'"
        )
    else:
        print(f"[LOADER] Found target sheet: '{sheet_name}'")

    # Read EVERYTHING as str to prevent float coercion on large IDs
    df = pd.read_excel(xl, sheet_name=sheet_name, dtype=str)

    # Normalise column headers
    df.columns = [str(c).strip() for c in df.columns]

    # Second-pass ID cleanup (belt-and-suspenders)
    for col in ID_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(clean_id)

    # Drop completely blank rows
    df = df.dropna(how="all")
    df = df[df.apply(lambda r: r.astype(str).str.strip().str.len().sum() > 0, axis=1)]

    print(f"[LOADER] Loaded {len(df):,} rows from sheet '{sheet_name}'.")
    return df


# ── BulkSheet File Discovery ──────────────────────────────────────────────────

def find_latest_bulksheet(base_dir: str) -> str:
    """
    Locate the most recently modified BulkSheetExport_*.xlsx file under
    <base_dir>/data/input/ or <base_dir>/data/, excluding any file that is
    itself a derived/processed output (e.g. *_filtered.xlsx, *_renamed.xlsx).

    Args:
        base_dir: Root directory of the project (where module scripts live).

    Returns:
        Absolute path to the most recent raw BulkSheet file.

    Raises:
        FileNotFoundError: If no matching file is discovered.
    """
    data_dir = os.path.join(base_dir, "data")
    patterns = [
        os.path.join(data_dir, "input", "BulkSheetExport_*.xlsx"),
        os.path.join(data_dir, "BulkSheetExport_*.xlsx"),
    ]
    exclusions = ["_renamed", "_filtered", "_unrenamed"]

    candidates: list[str] = []
    for pattern in patterns:
        candidates.extend(glob.glob(pattern))

    candidates = [
        f for f in candidates
        if not any(x in os.path.basename(f).lower() for x in exclusions)
    ]

    if not candidates:
        raise FileNotFoundError(
            f"[ERROR] No raw BulkSheetExport_*.xlsx found under {data_dir}. "
            "Place the Amazon export file in data/input/ and re-run."
        )

    # Return the newest file by modification time
    return sorted(candidates, key=os.path.getmtime, reverse=True)[0]
