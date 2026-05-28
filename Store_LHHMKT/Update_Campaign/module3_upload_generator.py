"""
module3_upload_generator.py
============================
MODULE 3 – Amazon Bulksheet Update File Generator

Pipeline Step:
    rename_map (Campaign ID → New Name)  +  source DataFrame per SKU
    →  Filter to Campaign-entity rows only
    →  Populate official Amazon template columns
    →  Save as Update_CampaignName_{SKU}.xlsx

Amazon Template Column Set (30 columns, official order):
    Product, Entity, Operation, Campaign ID, Ad Group ID, Portfolio ID,
    Ad ID, Keyword ID, Product Targeting ID, Campaign Name, Ad Group Name,
    Start Date, End Date, Targeting Type, State, Daily Budget, SKU,
    Ad Group Default Bid, Bid, Keyword Text, Native Language Keyword,
    Native Language Locale, Match Type, Bidding Strategy, Placement,
    Percentage, Product Targeting Expression, Audience ID,
    Shopper Cohort Percentage, Shopper Cohort Type

Critical Rules:
    • ONLY rows where Entity == 'Campaign' are written to the upload file.
    • Campaign ID  must be the CLEAN original ID string (no decimal / sci-notation).
    • Portfolio ID must be the CLEAN original string — unlinking a campaign from
      its portfolio by leaving this blank is a critical data-integrity error.
    • All other non-specified columns are left blank to avoid overwriting
      fields the user did not intend to change.

Output filename:  Update_CampaignName_{safe_sku}.xlsx
"""

import os
import sys

import openpyxl
import pandas as pd

from utils import clean_id, safe_filename, safe_str


# ── Official Amazon Template Column Order ─────────────────────────────────────
# Must match the header row in AmazonAdvertisingBulksheetSellerTemplate exactly.
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

# Static field values written for every Update row
_PRODUCT_VALUE   = "Sponsored Products"
_ENTITY_VALUE    = "Campaign"
_OPERATION_VALUE = "Update"
_STATE_VALUE     = "enabled"

# Output filename prefix
_OUTPUT_PREFIX   = "Update_CampaignName_"

# Entity string for filtering (lowercase comparison)
_ENTITY_CAMPAIGN = "campaign"


# ── Public API ────────────────────────────────────────────────────────────────

def generate_upload_file(
    df: pd.DataFrame,
    rename_map: dict[str, str],
    sku: str,
    output_dir: str,
    template_path: str | None = None,
) -> str | None:
    """
    Build and write the final Amazon Bulksheet Update file for a single SKU.

    Only rows where Entity == 'Campaign' are written. Each row is populated
    with the standardised campaign name and preserved Portfolio ID from the
    original cleaned data.

    Args:
        df:            The per-SKU DataFrame (all entity rows).
        rename_map:    { campaign_id_str → new_campaign_name_str }
                       Produced by Module 2.
        sku:           SKU string used to construct the output filename.
        output_dir:    Directory where the upload file will be saved.
        template_path: Optional path to the Amazon template XLSX.  When supplied,
                       the output is written into a copy of that template to
                       preserve any hidden sheets / formatting Amazon expects.
                       If None or the file doesn't exist, a plain XLSX is written.

    Returns:
        Absolute path of the file written, or None on failure.
    """
    os.makedirs(output_dir, exist_ok=True)

    # ── Guard: rename_map must not be empty ───────────────────────────────────
    if not rename_map:
        print(
            f"[MODULE 3] SKIP – No rename entries for SKU '{sku}'. "
            "No upload file generated."
        )
        return None

    # ── Step 1: Filter source df to Campaign-entity rows only ─────────────────
    if "Entity" not in df.columns:
        print(f"[MODULE 3] ERROR – 'Entity' column missing. Skipping SKU '{sku}'.")
        return None

    campaign_rows = df[
        df["Entity"].str.strip().str.lower() == _ENTITY_CAMPAIGN
    ].copy()

    if campaign_rows.empty:
        print(f"[MODULE 3] WARNING – No Campaign rows found in data for SKU '{sku}'.")
        return None

    # ── Step 2: Build the list of update row dicts ────────────────────────────
    rows_out: list[dict] = []

    for _, row in campaign_rows.iterrows():
        cid          = clean_id(row.get("Campaign ID", ""))
        portfolio_id = clean_id(row.get("Portfolio ID", ""))

        # Only write campaigns that have a resolved name in the rename map
        if cid not in rename_map:
            print(
                f"[MODULE 3]   SKIP – Campaign ID {cid} not in rename_map "
                f"(may have been skipped by Module 2). "
                "Row omitted from upload file."
            )
            continue

        new_name = rename_map[cid]

        # Initialise all template columns as empty strings
        row_out: dict[str, str] = {col: "" for col in TEMPLATE_COLUMNS}

        # Populate the required update fields
        row_out["Product"]       = _PRODUCT_VALUE
        row_out["Entity"]        = _ENTITY_VALUE
        row_out["Operation"]     = _OPERATION_VALUE
        row_out["Campaign ID"]   = cid           # Original clean ID – CRITICAL
        row_out["Portfolio ID"]  = portfolio_id  # Original clean ID – CRITICAL
        row_out["Campaign Name"] = new_name      # New standardised name
        row_out["State"]         = _STATE_VALUE

        rows_out.append(row_out)

    if not rows_out:
        print(
            f"[MODULE 3] WARNING – Zero eligible campaigns to write "
            f"for SKU '{sku}'. No file created."
        )
        return None

    # ── Step 3: Determine output file path ────────────────────────────────────
    safe_sku   = safe_filename(sku)
    out_fname  = f"{_OUTPUT_PREFIX}{safe_sku}.xlsx"
    out_path   = os.path.join(output_dir, out_fname)

    # ── Step 4: Write the file ────────────────────────────────────────────────
    # Strategy A: Copy the official Amazon template and fill the Campaigns sheet.
    # Strategy B: Plain openpyxl workbook (fallback when template unavailable).
    written = _write_with_template(rows_out, out_path, template_path)
    if not written:
        _write_plain(rows_out, out_path)

    print(
        f"[MODULE 3] [UPLOAD] SKU '{sku}' → {out_fname}  "
        f"({len(rows_out)} campaign row(s))"
    )
    return out_path


# ── Internal File Writers ─────────────────────────────────────────────────────

def _write_with_template(
    rows: list[dict],
    out_path: str,
    template_path: str | None,
) -> bool:
    """
    Write the update rows into a copy of the Amazon seller template XLSX.
    This preserves hidden sheets, named ranges, and cell formatting that
    Amazon's upload validation may check.

    Args:
        rows:          List of row dicts (keys = TEMPLATE_COLUMNS).
        out_path:      Destination file path.
        template_path: Path to AmazonAdvertisingBulksheetSellerTemplate.xlsx.

    Returns:
        True if successfully written via template, False otherwise.
    """
    if not template_path or not os.path.exists(template_path):
        return False  # Caller will fall back to plain writer

    try:
        wb = openpyxl.load_workbook(template_path)

        # Locate the Sponsored Products Campaigns sheet
        ws = None
        for candidate in ("Sponsored Products Campaigns", "Sponsored Product Campaigns"):
            if candidate in wb.sheetnames:
                ws = wb[candidate]
                break
        if ws is None:
            ws = wb.active

        # Clear any pre-existing data rows (keep header at row 1)
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)

        # Write each row cell-by-cell as a plain string
        for r_idx, row_dict in enumerate(rows, start=2):
            for c_idx, col_name in enumerate(TEMPLATE_COLUMNS, start=1):
                cell_val = row_dict.get(col_name, "")
                ws.cell(row=r_idx, column=c_idx, value=str(cell_val) if cell_val else "")

        # Remove all sheets that are not the campaigns sheet to keep file clean
        for sheet_name in list(wb.sheetnames):
            if sheet_name != ws.title:
                del wb[sheet_name]

        wb.save(out_path)
        return True

    except Exception as exc:
        print(
            f"[MODULE 3]   WARN – Template write failed for '{out_path}': {exc}. "
            "Falling back to plain writer."
        )
        return False


def _write_plain(rows: list[dict], out_path: str) -> None:
    """
    Write the update rows as a plain openpyxl workbook (no template dependency).
    Used as the fallback when the Amazon template is not available.

    Args:
        rows:     List of row dicts (keys = TEMPLATE_COLUMNS).
        out_path: Destination file path.
    """
    try:
        df_out = pd.DataFrame(rows, columns=TEMPLATE_COLUMNS)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df_out.to_excel(
                writer,
                sheet_name="Sponsored Products Campaigns",
                index=False,
            )
    except Exception as exc:
        print(f"[MODULE 3] ERROR – Failed to write '{out_path}': {exc}")


# ── Standalone Entry Point ────────────────────────────────────────────────────

if __name__ == "__main__":
    """
    Run Module 3 in isolation for testing.
    Usage:
        python module3_upload_generator.py <sku_file.xlsx> [template.xlsx]

    Combines Module 2 naming inline so this module can be tested end-to-end.
    """
    if len(sys.argv) < 2:
        print("Usage: python module3_upload_generator.py <sku_file.xlsx> [template.xlsx]")
        sys.exit(1)

    from utils import load_bulksheet, safe_filename
    from module2_naming import generate_rename_map

    sku_file      = sys.argv[1]
    template_file = sys.argv[2] if len(sys.argv) >= 3 else None
    sku_name      = os.path.splitext(os.path.basename(sku_file))[0].replace("SKU_", "")

    print("=" * 70)
    print("  MODULE 3 – Amazon Bulksheet Update Generator")
    print("=" * 70)
    print(f"[MODULE 3] Source:   {sku_file}")
    print(f"[MODULE 3] Template: {template_file or '(none – plain writer)'}\n")

    df_sku   = load_bulksheet(sku_file)
    rmap     = generate_rename_map(df_sku, sku_name)
    out_dir  = os.path.join(os.path.dirname(os.path.abspath(sku_file)), "upload_files")

    generate_upload_file(df_sku, rmap, sku_name, out_dir, template_file)
