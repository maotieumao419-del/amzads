"""
merge_upload_files.py
======================
POST-PIPELINE UTILITY – Master Upload File Merger

Purpose:
    After the 3-module pipeline generates individual per-SKU update files
    (Update_CampaignName_<SKU>.xlsx), this utility merges all of them into
    a single, upload-ready Master Amazon Bulksheet file.

Input:
    All files matching Update_CampaignName_*.xlsx in the scan directory.
    Default scan directory: data/upload_files/
    (This is the exact directory that Module 3 writes to.)

Output:
    data/upload_files/Master_Amazon_Bulksheet_Update_Upload.xlsx
    data/upload_files/Master_Amazon_Bulksheet_Update_Upload.csv  (utf-8-sig)

Data Integrity Guarantees:
    • All ID columns (Campaign ID, Portfolio ID, etc.) are forced to dtype=str
      on read, then passed through clean_id() as a second-pass safeguard.
    • No scientific notation (e+) or trailing decimals (.0) can survive.
    • Columns are aligned to the official Amazon template column order before save.

Usage:
    python merge_upload_files.py
    python merge_upload_files.py --input  data/upload_files/
    python merge_upload_files.py --input  data/upload_files/ --output data/final/
    python merge_upload_files.py --no-csv    (skip .csv output)
"""

import argparse
import glob
import os
import sys

import pandas as pd

# ── Import shared utilities ────────────────────────────────────────────────────
# clean_id() and ID_COLUMNS are defined once in utils.py and reused here,
# keeping ID-sanitisation logic strictly in one place.
from utils import clean_id, ID_COLUMNS


# ── Console Setup (belt-and-suspenders; utils.py also does this) ──────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Default input:  the directory Module 3 writes to
_DEFAULT_INPUT_DIR = os.path.join(BASE_DIR, "data", "upload_files")

# Default output: same directory (can be overridden via --output)
_DEFAULT_OUTPUT_DIR = _DEFAULT_INPUT_DIR

# Pattern that identifies per-SKU update files written by Module 3
_FILE_PATTERN = "Update_CampaignName_*.xlsx"

# Name of the output master file
_OUTPUT_XLSX = "Master_Amazon_Bulksheet_Update_Upload.xlsx"
_OUTPUT_CSV  = "Master_Amazon_Bulksheet_Update_Upload.csv"

# Amazon sheet name that upload validators look for
_SHEET_NAME = "Sponsored Products Campaigns"


# ── Official Amazon Template Column Order ─────────────────────────────────────
# Must match the exact header row in AmazonAdvertisingBulksheetSellerTemplate.
# Module 3 already writes in this order; we re-enforce it here for safety.
TEMPLATE_COLUMNS: list[str] = [
    "Product",
    "Entity",
    "Operation",
    "Campaign ID",
    "Ad Group ID",
    "Portfolio ID",
    "Ad ID",
    "Keyword ID",
    "Product Targeting ID",
    "Campaign Name",
    "Ad Group Name",
    "Start Date",
    "End Date",
    "Targeting Type",
    "State",
    "Daily Budget",
    "SKU",
    "Ad Group Default Bid",
    "Bid",
    "Keyword Text",
    "Native Language Keyword",
    "Native Language Locale",
    "Match Type",
    "Bidding Strategy",
    "Placement",
    "Percentage",
    "Product Targeting Expression",
    "Audience ID",
    "Shopper Cohort Percentage",
    "Shopper Cohort Type",
]


# ── Argument Parsing ──────────────────────────────────────────────────────────

def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Merge per-SKU Amazon Bulksheet update files into one Master Upload file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--input", "-i",
        metavar="DIR",
        default=_DEFAULT_INPUT_DIR,
        help=(
            f"Directory to scan for {_FILE_PATTERN} files. "
            f"Default: {_DEFAULT_INPUT_DIR}"
        ),
    )
    parser.add_argument(
        "--output", "-o",
        metavar="DIR",
        default=None,
        help=(
            "Directory to write the master output files. "
            "Default: same as --input."
        ),
    )
    parser.add_argument(
        "--no-csv",
        action="store_true",
        default=False,
        help="Skip generating the .csv companion file.",
    )
    parser.add_argument(
        "--store",
        default="Store",
        help="Store name for output file naming.",
    )
    parser.add_argument(
        "--task",
        default="SetupCampaignNames",
        help="Task name for output file naming.",
    )
    return parser.parse_args()


# ── Core Merge Function ───────────────────────────────────────────────────────

def merge_upload_files(
    input_dir: str,
    output_dir: str,
    write_csv: bool = True,
    store_name: str = "Store",
    task_name: str = "SetupCampaignNames",
) -> str | None:
    from datetime import datetime
    date_str = datetime.now().strftime("%d%m%y")
    output_xlsx = f"{store_name}_{task_name}_{date_str}.xlsx"
    output_csv  = f"{store_name}_{task_name}_{date_str}.csv"
    """
    Scan *input_dir* for per-SKU update files, merge them into a single
    master DataFrame, and write the result as an Amazon-compatible XLSX
    (and optionally a CSV).

    Args:
        input_dir:  Directory containing Update_CampaignName_*.xlsx files.
        output_dir: Directory where the master output files are written.
        write_csv:  If True, also saves a utf-8-sig encoded .csv alongside
                    the .xlsx (recommended for Windows Excel compatibility).

    Returns:
        Absolute path to the saved .xlsx, or None on failure.
    """
    print()
    print("=" * 70)
    print("  Amazon Ads – Master Upload File Merger")
    print("=" * 70)
    print(f"  Scan directory : {input_dir}")
    print(f"  Output directory: {output_dir}")
    print("=" * 70)
    print()

    # ── Step 1: Discover all per-SKU update files ─────────────────────────────
    scan_pattern = os.path.join(input_dir, _FILE_PATTERN)
    candidate_files = sorted(glob.glob(scan_pattern))

    # Exclude the master output file itself in case this script is re-run
    master_out_path = os.path.join(output_dir, output_xlsx)
    candidate_files = [
        f for f in candidate_files
        if os.path.abspath(f) != os.path.abspath(master_out_path)
    ]

    if not candidate_files:
        print(
            f"[MERGE] ERROR – No files matching '{_FILE_PATTERN}' found in:\n"
            f"        {input_dir}\n\n"
            "  Make sure you have run run_pipeline.py first so that Module 3\n"
            "  writes the individual Update_CampaignName_<SKU>.xlsx files.\n"
        )
        return None

    print(f"[MERGE] Found {len(candidate_files)} file(s) to merge:\n")

    # ── Step 2: Read and accumulate rows from every file ─────────────────────
    # We build a list of DataFrames then concatenate once at the end — this is
    # significantly faster than growing a single DataFrame row-by-row.
    frames: list[pd.DataFrame] = []
    total_rows_read = 0
    files_failed    = 0

    for idx, filepath in enumerate(candidate_files, start=1):
        filename = os.path.basename(filepath)

        try:
            # ── Read with dtype=str to prevent ANY numeric coercion ───────────
            # We target the Sponsored Products Campaigns sheet if it exists;
            # otherwise fall back to the first (and usually only) sheet.
            xls       = pd.ExcelFile(filepath, engine="openpyxl")
            sheet     = (
                "Sponsored Products Campaigns"
                if "Sponsored Products Campaigns" in xls.sheet_names
                else xls.sheet_names[0]
            )
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str)

            # ── Normalise column headers ───────────────────────────────────────
            df.columns = [str(c).strip() for c in df.columns]

            # ── Drop completely blank rows ────────────────────────────────────
            df = df.dropna(how="all")
            df = df[
                df.apply(
                    lambda r: r.astype(str).str.strip().str.len().sum() > 0,
                    axis=1,
                )
            ]

            if df.empty:
                print(
                    f"  [{idx:>3}/{len(candidate_files)}] SKIP  {filename} "
                    "– file contains no data rows."
                )
                continue

            # ── Second-pass ID cleanup (belt-and-suspenders) ──────────────────
            # Even with dtype=str, openpyxl occasionally converts large integers
            # to floats before Pandas applies the dtype cast.  clean_id() from
            # utils.py strips any remaining '.0' or 'e+' artefacts.
            for col in ID_COLUMNS:
                if col in df.columns:
                    df[col] = df[col].apply(clean_id)

            # ── Only keep Campaign entity rows (sanity guard) ─────────────────
            # Each individual file should already contain only Campaign rows,
            # but we verify here in case a file was manually edited.
            if "Entity" in df.columns:
                before = len(df)
                df = df[df["Entity"].str.strip().str.lower() == "campaign"]
                skipped = before - len(df)
                if skipped:
                    print(
                        f"  [{idx:>3}/{len(candidate_files)}]        {filename} "
                        f"– dropped {skipped} non-Campaign row(s)."
                    )

            if df.empty:
                print(
                    f"  [{idx:>3}/{len(candidate_files)}] SKIP  {filename} "
                    "– no Campaign entity rows found."
                )
                continue

            frames.append(df)
            total_rows_read += len(df)

            print(
                f"  [{idx:>3}/{len(candidate_files)}] OK    {filename:<55} "
                f"{len(df):>4} campaign row(s)"
            )

        except Exception as exc:
            print(
                f"  [{idx:>3}/{len(candidate_files)}] ERROR {filename}: {exc}"
            )
            files_failed += 1
            continue

    # ── Step 3: Guard – nothing was merged ────────────────────────────────────
    if not frames:
        print(
            "\n[MERGE] ERROR – No valid Campaign rows were collected. "
            "Master file not created.\n"
        )
        return None

    # ── Step 4: Concatenate all frames ────────────────────────────────────────
    master_df = pd.concat(frames, ignore_index=True)
    print(f"\n[MERGE] Total campaign rows aggregated: {len(master_df):,}")

    # ── Step 5: Align columns to the official Amazon template order ───────────
    # Build the output column list:
    #   • Start with TEMPLATE_COLUMNS that actually exist in the merged data.
    #   • Append any extra columns that were present in the source files but
    #     are not in the template (preserves unexpected data; avoids silent loss).
    present_template_cols = [c for c in TEMPLATE_COLUMNS if c in master_df.columns]
    extra_cols = [c for c in master_df.columns if c not in set(TEMPLATE_COLUMNS)]

    final_col_order = present_template_cols + extra_cols

    # Reindex to the final order; missing template columns become empty strings
    master_df = master_df.reindex(columns=final_col_order, fill_value="")

    # ── Step 6: Replace NaN remnants with empty string ────────────────────────
    # Ensures no 'nan' text appears in the uploaded file.
    master_df = master_df.fillna("")

    # ── Step 7: Write the Master XLSX ─────────────────────────────────────────
    os.makedirs(output_dir, exist_ok=True)
    xlsx_path = os.path.join(output_dir, output_xlsx)

    try:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            master_df.to_excel(
                writer,
                sheet_name=_SHEET_NAME,
                index=False,
            )
        print(f"[MERGE] XLSX saved → {xlsx_path}")
    except Exception as exc:
        print(f"[MERGE] ERROR – Could not write XLSX: {exc}")
        return None

    # ── Step 8: Write the companion CSV (optional) ────────────────────────────
    if write_csv:
        csv_path = os.path.join(output_dir, output_csv)
        try:
            # utf-8-sig BOM ensures Windows Excel opens it without garbled text
            master_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"[MERGE] CSV  saved → {csv_path}")
        except Exception as exc:
            print(f"[MERGE] WARNING – Could not write CSV: {exc}")

    # ── Step 9: Summary ───────────────────────────────────────────────────────
    print()
    print("─" * 70)
    print(f"  Files scanned   : {len(candidate_files)}")
    print(f"  Files merged    : {len(frames)}")
    print(f"  Files failed    : {files_failed}")
    print(f"  Total rows      : {len(master_df):,} campaign update(s)")
    print(f"  Output XLSX     : {output_xlsx}")
    if write_csv:
        print(f"  Output CSV      : {output_csv}")
    print("─" * 70)
    print()

    return xlsx_path


# ── Entry Point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    args = _parse_args()

    # Resolve absolute paths
    input_dir  = os.path.abspath(args.input)
    output_dir = os.path.abspath(args.output) if args.output else input_dir

    # Validate input directory existence
    if not os.path.isdir(input_dir):
        print(
            f"[MERGE] ERROR – Input directory does not exist: {input_dir}\n"
            "  Create it or pass a valid path with --input."
        )
        sys.exit(1)

    result = merge_upload_files(
        input_dir  = input_dir,
        output_dir = output_dir,
        write_csv  = not args.no_csv,
        store_name = args.store,
        task_name  = args.task,
    )

    sys.exit(0 if result else 1)
