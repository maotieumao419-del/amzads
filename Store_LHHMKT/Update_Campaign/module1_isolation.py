"""
module1_isolation.py
=====================
MODULE 1 – SKU-Based Isolation & Complex Campaign Detection

Pipeline Step:
    Raw BulkSheet DataFrame (from utils.load_bulksheet)
    →  Complex campaign detection & quarantine
    →  Standard campaigns grouped by SKU
    →  Individual per-SKU Excel files (full row sets)
    →  One quarantine file: Complex_Campaigns_Review.xlsx

Isolation Rules:
    COMPLEX  = campaign has MULTIPLE unique SKUs   OR
               has BOTH Keyword rows AND Product Targeting rows  (mixed targeting)
    STANDARD = single SKU, single targeting type

Outputs (written to <output_dir>):
    • SKU_<safe_sku>.xlsx          – one file per unique standard SKU
    • Complex_Campaigns_Review.xlsx – all flagged complex campaigns
"""

import os
import sys

import pandas as pd

from utils import (
    clean_id,
    safe_str,
    safe_filename,
)


# ── Constants ─────────────────────────────────────────────────────────────────

# Entity values we inspect when classifying a campaign's targeting type.
_ENTITY_CAMPAIGN          = "campaign"
_ENTITY_PRODUCT_AD        = "product ad"
_ENTITY_KEYWORD           = "keyword"
_ENTITY_PRODUCT_TARGETING = "product targeting"

# Output file naming
_COMPLEX_FILENAME = "Complex_Campaigns_Review.xlsx"
_SKU_PREFIX       = "SKU_"


# ── Core Function ─────────────────────────────────────────────────────────────

def isolate_campaigns(
    df: pd.DataFrame,
    output_dir: str,
) -> tuple[dict[str, list[str]], list[str]]:
    """
    Analyse every campaign in *df*, separate complex ones from standard ones,
    and write per-SKU Excel files for standard campaigns.

    Args:
        df:         Cleaned BulkSheet DataFrame (all values as str).
        output_dir: Directory where output Excel files will be saved.

    Returns:
        A tuple of:
            standard_sku_map : dict[sku_str → list[campaign_id_str]]
                Mapping of each SKU to the Campaign IDs it owns.
            complex_campaign_ids : list[str]
                Campaign IDs that were quarantined as complex.
    """
    os.makedirs(output_dir, exist_ok=True)

    # ── Verify required columns exist ─────────────────────────────────────────
    required_cols = {"Campaign ID", "Entity"}
    missing = required_cols - set(df.columns)
    if missing:
        print(f"[MODULE 1] ERROR – Missing required columns: {missing}")
        sys.exit(1)

    # Determine the Campaign Name column (Amazon uses several variants)
    name_col = _resolve_name_col(df)

    # ── Step 1: Group every DataFrame row by its Campaign ID ──────────────────
    # by_campaign: { campaign_id (str) → list of row dicts }
    by_campaign: dict[str, list[dict]] = {}
    for _, row in df.iterrows():
        cid = clean_id(row.get("Campaign ID"))
        if cid:
            by_campaign.setdefault(cid, []).append(row.to_dict())

    print(f"[MODULE 1] Total unique Campaign IDs found: {len(by_campaign):,}")

    # ── Step 2: Classify every campaign ───────────────────────────────────────
    # Collect Campaign IDs into two buckets.
    complex_cids:  list[str] = []    # Multi-SKU or mixed-targeting → quarantine
    standard_cids: list[str] = []    # Clean single-SKU / single-targeting type

    # We'll also build the SKU→CID mapping for standard campaigns here.
    standard_sku_map: dict[str, list[str]] = {}   # sku → [cid, ...]
    complex_reasons:  dict[str, str] = {}          # cid → human-readable reason

    for cid, rows in by_campaign.items():
        reason, sku = _classify_campaign(cid, rows)

        if reason:
            # Complex – quarantine
            complex_cids.append(cid)
            complex_reasons[cid] = reason
        else:
            # Standard – add to per-SKU bucket
            standard_cids.append(cid)
            standard_sku_map.setdefault(sku, []).append(cid)

    # ── Step 3: Console warnings for complex campaigns ─────────────────────────
    if complex_cids:
        print()
        print("=" * 70)
        print(
            f"  [MODULE 1] ⚠️  COMPLEX CAMPAIGNS DETECTED – {len(complex_cids)} campaign(s)"
        )
        print("  These campaigns require MANUAL REVIEW and have been quarantined.")
        print("=" * 70)
        for cid in complex_cids:
            reason = complex_reasons[cid]
            camp_name = _get_campaign_name(by_campaign[cid], name_col)
            print(f"  • Campaign ID: {cid} | Name: {camp_name}")
            print(f"    Reason: {reason}")
        print("=" * 70)
        print()

    # ── Step 4: Write the Complex_Campaigns_Review.xlsx quarantine file ────────
    if complex_cids:
        _write_complex_file(df, complex_cids, output_dir)

    # ── Step 5: Write one Excel file per standard SKU (full row sets) ──────────
    print(f"[MODULE 1] Writing per-SKU files for {len(standard_sku_map)} unique SKU(s)…")
    for sku, cids in standard_sku_map.items():
        _write_sku_file(df, sku, cids, output_dir)

    print(
        f"\n[MODULE 1] Done. "
        f"{len(standard_sku_map)} standard SKU file(s) written, "
        f"{len(complex_cids)} complex campaign(s) quarantined."
    )

    return standard_sku_map, complex_cids


# ── Classification Helper ─────────────────────────────────────────────────────

def _classify_campaign(
    cid: str,
    rows: list[dict],
) -> tuple[str, str]:
    """
    Inspect all rows belonging to *cid* and decide if the campaign is complex.

    Returns:
        (reason, sku)
            • reason is '' for standard campaigns; a non-empty string otherwise.
            • sku is the single SKU found (empty string for complex campaigns).
    """
    # Collect entity types within this campaign
    product_ad_rows = [
        r for r in rows
        if safe_str(r.get("Entity")).lower() == _ENTITY_PRODUCT_AD
    ]
    keyword_rows = [
        r for r in rows
        if safe_str(r.get("Entity")).lower() == _ENTITY_KEYWORD
    ]
    pt_rows = [
        r for r in rows
        if safe_str(r.get("Entity")).lower() == _ENTITY_PRODUCT_TARGETING
    ]

    # Collect all unique non-empty SKUs
    skus = {
        safe_str(r.get("SKU"))
        for r in product_ad_rows
        if safe_str(r.get("SKU"))
    }

    # Edge case 1: Multiple SKUs → complex
    if len(skus) > 1:
        reason = (
            f"Multiple SKUs detected ({len(skus)}): "
            + ", ".join(sorted(skus))
        )
        return reason, ""

    # Edge case 2: Mixed targeting (both Keywords AND Product Targeting) → complex
    has_kw = len(keyword_rows) > 0
    has_pt = len(pt_rows)  > 0
    if has_kw and has_pt:
        reason = (
            f"Mixed targeting: {len(keyword_rows)} Keyword row(s) "
            f"AND {len(pt_rows)} Product Targeting row(s)"
        )
        return reason, ""

    # Standard: resolve the single SKU
    sku = next(iter(skus)) if skus else "UNKNOWN_SKU"
    return "", sku


# ── File Writers ──────────────────────────────────────────────────────────────

def _write_complex_file(
    df: pd.DataFrame,
    complex_cids: list[str],
    output_dir: str,
) -> None:
    """
    Extract all rows belonging to the complex campaign IDs and save them to
    Complex_Campaigns_Review.xlsx.

    Args:
        df:           Full source DataFrame.
        complex_cids: List of Campaign IDs that are quarantined.
        output_dir:   Destination directory.
    """
    cid_set = set(complex_cids)
    mask = df["Campaign ID"].apply(lambda v: clean_id(v) in cid_set)
    df_complex = df[mask].copy()

    out_path = os.path.join(output_dir, _COMPLEX_FILENAME)
    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df_complex.to_excel(
                writer,
                sheet_name="Sponsored Products Campaigns",
                index=False,
            )
        print(
            f"[MODULE 1] [QUARANTINE] Saved {len(df_complex):,} rows "
            f"({len(complex_cids)} campaigns) → {_COMPLEX_FILENAME}"
        )
    except Exception as exc:
        print(f"[MODULE 1] ERROR – Could not write {_COMPLEX_FILENAME}: {exc}")


def _write_sku_file(
    df: pd.DataFrame,
    sku: str,
    campaign_ids: list[str],
    output_dir: str,
) -> None:
    """
    Extract ALL rows (Campaign, Ad Group, Product Ad, Keyword, Product Targeting)
    for the given list of Campaign IDs and write them to a per-SKU Excel file.

    This preserves the full campaign architecture, making the file auditable.

    Args:
        df:           Full source DataFrame.
        sku:          SKU string (used for the filename).
        campaign_ids: List of Campaign IDs belonging to this SKU.
        output_dir:   Destination directory.
    """
    cid_set = set(campaign_ids)
    mask = df["Campaign ID"].apply(lambda v: clean_id(v) in cid_set)
    df_sku = df[mask].copy()

    filename = f"{_SKU_PREFIX}{safe_filename(sku)}.xlsx"
    out_path = os.path.join(output_dir, filename)

    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df_sku.to_excel(
                writer,
                sheet_name="Sponsored Products Campaigns",
                index=False,
            )
        print(
            f"[MODULE 1]   [SKU] '{sku}' → {filename}  "
            f"({len(campaign_ids)} campaign(s), {len(df_sku):,} rows)"
        )
    except Exception as exc:
        print(f"[MODULE 1] ERROR – Could not write SKU file for '{sku}': {exc}")


# ── Column Resolution Helpers ─────────────────────────────────────────────────

def _resolve_name_col(df: pd.DataFrame) -> str:
    """
    Find the Campaign Name column, checking multiple Amazon variant names.
    Returns the column name found, or raises ValueError if none is present.
    """
    candidates = [
        "Campaign Name (Informational only)",
        "Campaign Name",
        "Name",
        "Campaign",
    ]
    for col in candidates:
        if col in df.columns:
            return col
    raise ValueError(
        f"[MODULE 1] ERROR – No Campaign Name column found. "
        f"Columns present: {list(df.columns)}"
    )


def _get_campaign_name(rows: list[dict], name_col: str) -> str:
    """
    Return the campaign name from the 'Campaign' entity row of a campaign group.
    Falls back to the first row's name if no Campaign entity row is found.
    """
    camp_row = next(
        (r for r in rows if safe_str(r.get("Entity")).lower() == _ENTITY_CAMPAIGN),
        None,
    )
    if camp_row:
        return safe_str(camp_row.get(name_col) or camp_row.get("Campaign Name", ""))
    return safe_str(rows[0].get(name_col, "")) if rows else ""


# ── Standalone Entry Point ────────────────────────────────────────────────────

if __name__ == "__main__":
    """
    Run Module 1 in isolation.
    Reads the latest BulkSheetExport_*.xlsx from data/input/,
    writes per-SKU files and the quarantine file to data/sku_files/.
    """
    from utils import find_latest_bulksheet, load_bulksheet

    BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
    OUTPUT_DIR = os.path.join(BASE_DIR, "data", "sku_files")

    print("=" * 70)
    print("  MODULE 1 – SKU Isolation & Complex Campaign Detection")
    print("=" * 70)

    bulk_path = find_latest_bulksheet(BASE_DIR)
    print(f"[MODULE 1] Source file: {bulk_path}\n")

    df_raw = load_bulksheet(bulk_path)
    isolate_campaigns(df_raw, OUTPUT_DIR)
