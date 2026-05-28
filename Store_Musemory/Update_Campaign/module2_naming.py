"""
module2_naming.py
==================
MODULE 2 – Campaign Naming SOP Implementation

Pipeline Step:
    Per-SKU DataFrames (output of Module 1)
    →  Resolve the canonical ASIN for each SKU file
    →  Apply strict naming SOP to every campaign
    →  De-duplicate generated names with _v2_Nguyen versioning
    →  Return a rename map: { campaign_id → new_campaign_name }

Naming SOP Architecture:
    AUTO      →  {ASIN}_SP02_AUTO_Nguyen
    Keyword   →  {ASIN}_SP00_{KT_Suffix}_Nguyen          (if KT pattern in old name)
              →  {ASIN}_SP00_{keyword}_{MATCH_TYPE}_Nguyen (data-driven fallback)
    PT        →  {ASIN}_SP03_{PT_Suffix}_Nguyen           (if PT pattern in old name)
              →  {ASIN}_SP03_PT_{Competitor_ASIN}_Nguyen  (competitor ASIN)
              →  {ASIN}_SP03_PT_{cleaned_expr}_Nguyen     (expression fallback)

CRITICAL ASIN Resolution Order (per SKU file):
    1. 'ASIN (Informational only)' or 'ASIN' column on Product Ad row
    2. Regex extraction from the SKU string itself
    3. Regex extraction from the original Campaign Name string
    4. SKIP campaign (warn user) if all sources fail
"""

import re
import sys
import os

import pandas as pd

from utils import (
    clean_id,
    safe_str,
    extract_asin,
    clean_keyword,
    no_double_underscore,
)


# ── Entity Constants ──────────────────────────────────────────────────────────
_ENTITY_CAMPAIGN          = "campaign"
_ENTITY_PRODUCT_AD        = "product ad"
_ENTITY_KEYWORD           = "keyword"
_ENTITY_PRODUCT_TARGETING = "product targeting"

# Auto-campaign target expressions used in SP02 AUTO campaigns
_AUTO_EXPRESSIONS = {"close-match", "loose-match", "substitutes", "complements"}

# Regex: captures a KT-prefixed suffix (e.g. KT01, KT_BROAD, KTbirthday-mom)
_KT_PATTERN = re.compile(r"KT([A-Za-z0-9_\-]+)", re.IGNORECASE)

# Regex: captures a PT-prefixed suffix (e.g. PT01, PT_ASIN, PTcompetitor)
_PT_PATTERN = re.compile(r"PT([A-Za-z0-9_\-]+)", re.IGNORECASE)

# Regex: strips any trailing 'nguyen' (with optional leading separator) from a suffix
_TRAILING_NGUYEN = re.compile(r"[_\- ]*nguyen$", re.IGNORECASE)


# ── Public API ────────────────────────────────────────────────────────────────

def generate_rename_map(
    df: pd.DataFrame,
    sku: str,
) -> dict[str, str]:
    """
    Given a DataFrame containing ALL rows for a single SKU's campaigns,
    produce a dict mapping each Campaign ID to its new standardised name.

    The ASIN used as the name prefix is resolved once for the entire SKU file
    (see _resolve_asin_for_sku()) to guarantee 100% SKU-ASIN alignment.

    Args:
        df:  DataFrame for one SKU (output of Module 1 per-SKU file).
             Must contain 'Campaign ID', 'Entity', 'Campaign Name', etc.
        sku: The SKU string (used as fallback ASIN source and for console messages).

    Returns:
        A dict: { campaign_id_str → new_campaign_name_str }
        Campaigns that cannot be processed (no ASIN, missing structure) are omitted.
    """
    # ── Validate required columns ──────────────────────────────────────────────
    for col in ("Campaign ID", "Entity"):
        if col not in df.columns:
            print(f"[MODULE 2] ERROR – Missing column '{col}'. Skipping SKU '{sku}'.")
            return {}

    name_col = _resolve_name_col(df)

    # ── Step 1: Resolve ASIN for this entire SKU file ─────────────────────────
    asin = _resolve_asin_for_sku(df, sku, name_col)

    if not asin:
        print(
            f"[MODULE 2] WARNING – Could not resolve ASIN for SKU '{sku}'. "
            "All campaigns in this file will be skipped from naming."
        )
        return {}

    print(f"[MODULE 2] SKU '{sku}' → ASIN resolved as: {asin}")

    # ── Step 2: Group rows by Campaign ID ─────────────────────────────────────
    by_campaign: dict[str, list[dict]] = {}
    for _, row in df.iterrows():
        cid = clean_id(row.get("Campaign ID"))
        if cid:
            by_campaign.setdefault(cid, []).append(row.to_dict())

    # ── Step 3: Generate a new name for every campaign ────────────────────────
    # global_names tracks all names produced so far (within this SKU scope)
    # to enable de-duplication with versioning.
    global_names: dict[str, str] = {}   # generated_name → campaign_id
    rename_map:   dict[str, str] = {}   # campaign_id    → generated_name

    for cid, rows in by_campaign.items():
        new_name = _name_single_campaign(cid, rows, asin, name_col)

        if not new_name:
            print(
                f"[MODULE 2]   SKIP – Campaign ID {cid}: "
                "could not determine campaign type (no keywords/PT/AUTO structure)."
            )
            continue

        # ── De-duplication: append _v2_Nguyen, _v3_Nguyen … if name is taken ─
        final_name = _deduplicate_name(new_name, cid, global_names)
        global_names[final_name] = cid
        rename_map[cid] = final_name

    print(
        f"[MODULE 2]   {len(rename_map)}/{len(by_campaign)} campaigns "
        f"named successfully for SKU '{sku}'."
    )
    return rename_map


# ── ASIN Resolution ───────────────────────────────────────────────────────────

def _resolve_asin_for_sku(
    df: pd.DataFrame,
    sku: str,
    name_col: str,
) -> str:
    """
    Resolve the single canonical ASIN for the entire SKU file using a strict
    priority order:
        1. 'ASIN (Informational only)' or 'ASIN' column on any Product Ad row
        2. Regex on the SKU string
        3. Regex on any Campaign Name string in the file
        4. Return '' (caller handles the empty case)

    Args:
        df:       The per-SKU DataFrame.
        sku:      The raw SKU string.
        name_col: Resolved campaign name column name.

    Returns:
        ASIN string in UPPERCASE, or '' if unresolvable.
    """
    # Priority 1 – Product Ad row ASIN columns
    pa_rows = df[df["Entity"].str.strip().str.lower() == _ENTITY_PRODUCT_AD]
    for asin_col in ("ASIN (Informational only)", "ASIN"):
        if asin_col in df.columns:
            for val in pa_rows[asin_col].dropna():
                asin = extract_asin(str(val))
                if asin:
                    return asin.upper()

    # Priority 2 – Extract ASIN directly from the SKU string
    asin = extract_asin(sku)
    if asin:
        return asin.upper()

    # Priority 3 – Extract ASIN from any Campaign Name in this file
    if name_col in df.columns:
        for name_val in df[name_col].dropna():
            asin = extract_asin(str(name_val))
            if asin:
                return asin.upper()

    return ""  # All sources exhausted


# ── Single Campaign Naming ────────────────────────────────────────────────────

def _name_single_campaign(
    cid: str,
    rows: list[dict],
    asin: str,
    name_col: str,
) -> str:
    """
    Determine and construct the standardized name for one campaign.

    Classification priority:
        1. If original name contains 'auto' → SP02 AUTO
        2. Keyword rows present             → SP00 Keyword campaign
        3. Product Targeting rows present   → SP03 Product Targeting campaign
        4. None of the above               → return '' (will be skipped)

    Args:
        cid:      Campaign ID (for logging).
        rows:     All rows belonging to this campaign.
        asin:     The resolved ASIN prefix for this SKU file.
        name_col: The campaign name column key.

    Returns:
        The generated campaign name string, or '' if classification fails.
    """
    # Find the campaign header row (Entity = 'campaign')
    camp_row = next(
        (r for r in rows if safe_str(r.get("Entity")).lower() == _ENTITY_CAMPAIGN),
        None,
    )
    orig_name = safe_str(camp_row.get(name_col, "") if camp_row else "")

    # Collect targeting sub-rows
    kw_rows = [r for r in rows if safe_str(r.get("Entity")).lower() == _ENTITY_KEYWORD]
    pt_rows = [r for r in rows if safe_str(r.get("Entity")).lower() == _ENTITY_PRODUCT_TARGETING]

    # ── Branch 1: AUTO Campaign ───────────────────────────────────────────────
    if "auto" in orig_name.lower():
        return no_double_underscore(f"{asin}_SP02_AUTO_Nguyen")

    # ── Branch 2: Keyword Campaign (SP00) ─────────────────────────────────────
    if kw_rows:
        return _build_sp00_name(asin, orig_name, kw_rows)

    # ── Branch 3: Product Targeting Campaign (SP03) ───────────────────────────
    if pt_rows:
        return _build_sp03_name(asin, orig_name, pt_rows)

    # ── Unclassified: only AUTO expressions in PT → treat as AUTO ─────────────
    # Edge case: campaign has product targeting rows that are all auto expressions
    # but wasn't caught by the 'auto' in orig_name check above.
    if not kw_rows and not pt_rows:
        # No targeting structure at all
        return ""

    return ""  # Fallthrough – should not reach here normally


def _build_sp00_name(asin: str, orig_name: str, kw_rows: list[dict]) -> str:
    """
    Construct the SP00 Keyword campaign name.

    Logic:
        • If the original name has a KT-suffix pattern → reuse it (strip trailing Nguyen).
        • Otherwise → use first keyword text + match type from actual data.

    Args:
        asin:     Resolved ASIN prefix.
        orig_name: The original campaign name string.
        kw_rows:  All Keyword entity rows for this campaign.

    Returns:
        Fully formed SP00 campaign name string.
    """
    # Check for KT pattern in original name
    kt_match = _KT_PATTERN.search(orig_name)
    if kt_match:
        # Extract the portion after 'KT' and clean trailing 'nguyen'
        kt_body = kt_match.group(1)
        kt_body = _TRAILING_NGUYEN.sub("", kt_body).strip("_- ")
        # Reconstruct full KT suffix including the 'KT' prefix
        kt_suffix = f"KT{kt_body}"
        return no_double_underscore(f"{asin}_SP00_{kt_suffix}_Nguyen")

    # Fallback: derive name from actual keyword data
    first_kw = kw_rows[0]
    kw_text   = safe_str(first_kw.get("Keyword Text", ""))
    match_type = safe_str(first_kw.get("Match Type", "EXACT")).upper() or "EXACT"

    if kw_text:
        cleaned = clean_keyword(kw_text)
        if cleaned:
            return no_double_underscore(f"{asin}_SP00_{cleaned}_{match_type}_Nguyen")

    # Last resort: generic unknown keyword name
    return no_double_underscore(f"{asin}_SP00_Unknown_{match_type}_Nguyen")


def _build_sp03_name(asin: str, orig_name: str, pt_rows: list[dict]) -> str:
    """
    Construct the SP03 Product Targeting campaign name.

    Logic:
        • If the original name has a PT-suffix pattern → reuse it (strip trailing Nguyen).
        • Otherwise → try to extract a competitor ASIN from the targeting expression.
        • Fallback → clean the expression text directly.

    Args:
        asin:     Resolved ASIN prefix.
        orig_name: The original campaign name string.
        pt_rows:  All Product Targeting entity rows for this campaign.

    Returns:
        Fully formed SP03 campaign name string.
    """
    # Check for PT pattern in original name
    pt_match = _PT_PATTERN.search(orig_name)
    if pt_match:
        pt_body = pt_match.group(1)
        pt_body = _TRAILING_NGUYEN.sub("", pt_body).strip("_- ")
        pt_suffix = f"PT{pt_body}"
        return no_double_underscore(f"{asin}_SP03_{pt_suffix}_Nguyen")

    # Fallback: use first Product Targeting expression
    first_pt = pt_rows[0]
    expr = safe_str(first_pt.get("Product Targeting Expression", ""))

    # Skip auto-target expressions (those belong to SP02)
    if expr.lower() in _AUTO_EXPRESSIONS:
        return no_double_underscore(f"{asin}_SP02_AUTO_Nguyen")

    # Try to extract competitor ASIN from the expression
    competitor_asin = extract_asin(expr)
    if competitor_asin:
        return no_double_underscore(f"{asin}_SP03_PT_{competitor_asin}_Nguyen")

    # Clean the expression text and use it directly
    if expr:
        cleaned_expr = clean_keyword(expr)
        if cleaned_expr:
            return no_double_underscore(f"{asin}_SP03_PT_{cleaned_expr}_Nguyen")

    # Absolute fallback
    return no_double_underscore(f"{asin}_SP03_PT_Unknown_Nguyen")


# ── De-duplication ────────────────────────────────────────────────────────────

def _deduplicate_name(
    base_name: str,
    cid: str,
    global_names: dict[str, str],
) -> str:
    """
    Ensure the generated campaign name is unique within the current processing scope.
    If *base_name* is already taken by a DIFFERENT campaign, append _v2_Nguyen,
    _v3_Nguyen, etc. until a free slot is found.

    Versioning format:
        {ASIN}_SP00_{kw}_EXACT_Nguyen        (original)
        {ASIN}_SP00_{kw}_EXACT_v2_Nguyen     (first duplicate)
        {ASIN}_SP00_{kw}_EXACT_v3_Nguyen     (second duplicate)
        ...

    Args:
        base_name:    The initially generated name (may conflict).
        cid:          The campaign ID that owns this name.
        global_names: Existing { name → campaign_id } registry.

    Returns:
        A unique campaign name string.
    """
    # If the name doesn't exist yet, or it was already assigned to THIS campaign
    if base_name not in global_names or global_names[base_name] == cid:
        return base_name

    # Strip existing '_Nguyen' suffix to inject version number before it
    stem = base_name
    if stem.upper().endswith("_NGUYEN"):
        stem = stem[:-7]   # remove '_Nguyen' (7 chars)

    counter = 2
    while True:
        candidate = f"{stem}_v{counter}_Nguyen"
        if candidate not in global_names or global_names[candidate] == cid:
            return candidate
        counter += 1


# ── Column Resolution Helper ──────────────────────────────────────────────────

def _resolve_name_col(df: pd.DataFrame) -> str:
    """
    Locate the Campaign Name column, checking multiple Amazon column-name variants.
    Precedence: informational-only variant > standard name > fallbacks.
    """
    for col in (
        "Campaign Name (Informational only)",
        "Campaign Name",
        "Name",
        "Campaign",
    ):
        if col in df.columns:
            return col
    raise ValueError(
        f"[MODULE 2] No Campaign Name column found. "
        f"Columns: {list(df.columns)}"
    )


# ── Standalone Entry Point ────────────────────────────────────────────────────

if __name__ == "__main__":
    """
    Run Module 2 in isolation against a single per-SKU Excel file.
    Usage:  python module2_naming.py <path_to_sku_file.xlsx>
    """
    if len(sys.argv) < 2:
        print("Usage: python module2_naming.py <path_to_sku_file.xlsx>")
        sys.exit(1)

    from utils import load_bulksheet

    sku_file = sys.argv[1]
    sku_name = os.path.splitext(os.path.basename(sku_file))[0].replace("SKU_", "")

    print("=" * 70)
    print("  MODULE 2 – Campaign Naming SOP Implementation")
    print("=" * 70)
    print(f"[MODULE 2] Processing file: {sku_file}")
    print(f"[MODULE 2] Inferred SKU: '{sku_name}'\n")

    df_sku = load_bulksheet(sku_file)
    rmap   = generate_rename_map(df_sku, sku_name)

    print("\n[MODULE 2] Rename Map:")
    for campaign_id, new_name in rmap.items():
        print(f"  {campaign_id}  →  {new_name}")
