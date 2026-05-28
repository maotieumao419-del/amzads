"""
module2_naming.py
==================
MODULE 2 – Campaign Naming SOP Implementation (SKU-prefix variant)

Pipeline Step:
    Per-SKU DataFrames (output of Module 1)
    →  Use the SKU directly as the campaign name prefix (instead of ASIN)
    →  Apply strict naming SOP to every campaign
    →  De-duplicate generated names with _v2_Nguyen versioning
    →  Return a rename map: { campaign_id → new_campaign_name }

Naming SOP Architecture:
    AUTO      →  {SKU}_SP02_AUTO_Nguyen
    Keyword   →  {SKU}_SP00_{KT_Suffix}_Nguyen          (if KT pattern in old name)
              →  {SKU}_SP00_{keyword}_{MATCH_TYPE}_Nguyen (data-driven fallback)
    PT        →  {SKU}_SP03_{PT_Suffix}_Nguyen           (if PT pattern in old name)
              →  {SKU}_SP03_PT_{Competitor_ASIN}_Nguyen  (competitor ASIN)
              →  {SKU}_SP03_PT_{cleaned_expr}_Nguyen     (expression fallback)

NOTE (vs. original D_setup_name):
    The ONLY difference from the original module2_naming.py is that the
    campaign name prefix uses the RAW SKU string directly instead of the
    resolved ASIN.  All other logic (KT/PT suffix reuse, de-duplication,
    AUTO detection, etc.) is identical.
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
    safe_filename,
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


# ── SKU Sanitisation ──────────────────────────────────────────────────────────

def _sanitise_sku_for_name(sku: str) -> str:
    """
    Convert a raw SKU string into a safe campaign-name component.

    Rules:
        • Replace any character that is NOT alphanumeric or hyphen with '_'
        • Collapse consecutive underscores to a single one
        • Strip leading/trailing underscores

    This mirrors safe_filename() but is specific to campaign-name embedding
    (dots are also replaced to avoid confusion with ASIN patterns).

    Args:
        sku: The raw SKU string from BulkSheet data.

    Returns:
        A clean, upper-cased string safe for use in a campaign name.
    """
    cleaned = re.sub(r"[^\w\-]", "_", sku)          # keep word-chars + hyphens
    cleaned = re.sub(r"_+", "_", cleaned).strip("_") # collapse consecutive _
    return cleaned.upper()


# ── Public API ────────────────────────────────────────────────────────────────

def generate_rename_map(
    df: pd.DataFrame,
    sku: str,
) -> dict[str, str]:
    """
    Given a DataFrame containing ALL rows for a single SKU's campaigns,
    produce a dict mapping each Campaign ID to its new standardised name.

    The SKU is used directly as the campaign name prefix (this is the key
    difference from the original D_setup_name version which used the ASIN).

    Args:
        df:  DataFrame for one SKU (output of Module 1 per-SKU file).
             Must contain 'Campaign ID', 'Entity', 'Campaign Name', etc.
        sku: The SKU string – used as the name prefix AND for console messages.

    Returns:
        A dict: { campaign_id_str → new_campaign_name_str }
        Campaigns that cannot be classified are omitted.
    """
    # ── Validate required columns ──────────────────────────────────────────────
    for col in ("Campaign ID", "Entity"):
        if col not in df.columns:
            print(f"[MODULE 2] ERROR – Missing column '{col}'. Skipping SKU '{sku}'.")
            return {}

    name_col = _resolve_name_col(df)

    # ── Build the sanitised SKU prefix ────────────────────────────────────────
    sku_prefix = _sanitise_sku_for_name(sku)
    if not sku_prefix:
        print(
            f"[MODULE 2] WARNING – SKU '{sku}' produced an empty prefix after "
            "sanitisation. All campaigns will be skipped."
        )
        return {}

    print(f"[MODULE 2] SKU '{sku}' → prefix used in campaign names: '{sku_prefix}'")

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
        new_name = _name_single_campaign(cid, rows, sku_prefix, name_col)

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


# ── Single Campaign Naming ────────────────────────────────────────────────────

def _name_single_campaign(
    cid: str,
    rows: list[dict],
    sku_prefix: str,
    name_col: str,
) -> str:
    """
    Determine and construct the standardized name for one campaign.

    Uses *sku_prefix* (the sanitised SKU) as the leading component of the
    name instead of an ASIN – this is the sole behavioural change vs. the
    original D_setup_name module.

    Classification priority:
        1. If original name contains 'auto' → SP02 AUTO
        2. Keyword rows present             → SP00 Keyword campaign
        3. Product Targeting rows present   → SP03 Product Targeting campaign
        4. None of the above               → return '' (will be skipped)

    Args:
        cid:        Campaign ID (for logging).
        rows:       All rows belonging to this campaign.
        sku_prefix: The sanitised SKU string used as the name prefix.
        name_col:   The campaign name column key.

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
        return no_double_underscore(f"{sku_prefix}_SP02_AUTO_Nguyen")

    # ── Branch 2: Keyword Campaign (SP00) ─────────────────────────────────────
    if kw_rows:
        return _build_sp00_name(sku_prefix, orig_name, kw_rows)

    # ── Branch 3: Product Targeting Campaign (SP03) ───────────────────────────
    if pt_rows:
        return _build_sp03_name(sku_prefix, orig_name, pt_rows)

    # ── Unclassified: no targeting structure at all ───────────────────────────
    return ""


def _build_sp00_name(sku_prefix: str, orig_name: str, kw_rows: list[dict]) -> str:
    """
    Construct the SP00 Keyword campaign name using *sku_prefix* instead of ASIN.

    Logic:
        • If the original name has a KT-suffix pattern → reuse it (strip trailing Nguyen).
        • Otherwise → use first keyword text + match type from actual data.

    Args:
        sku_prefix: Sanitised SKU used as the name prefix.
        orig_name:  The original campaign name string.
        kw_rows:    All Keyword entity rows for this campaign.

    Returns:
        Fully formed SP00 campaign name string.
    """
    # Check for KT pattern in original name
    kt_match = _KT_PATTERN.search(orig_name)
    if kt_match:
        kt_body = kt_match.group(1)
        kt_body = _TRAILING_NGUYEN.sub("", kt_body).strip("_- ")
        kt_suffix = f"KT{kt_body}"
        return no_double_underscore(f"{sku_prefix}_SP00_{kt_suffix}_Nguyen")

    # Fallback: derive name from actual keyword data
    first_kw   = kw_rows[0]
    kw_text    = safe_str(first_kw.get("Keyword Text", ""))
    match_type = safe_str(first_kw.get("Match Type", "EXACT")).upper() or "EXACT"

    if kw_text:
        cleaned = clean_keyword(kw_text)
        if cleaned:
            return no_double_underscore(f"{sku_prefix}_SP00_{cleaned}_{match_type}_Nguyen")

    # Last resort: generic unknown keyword name
    return no_double_underscore(f"{sku_prefix}_SP00_Unknown_{match_type}_Nguyen")


def _build_sp03_name(sku_prefix: str, orig_name: str, pt_rows: list[dict]) -> str:
    """
    Construct the SP03 Product Targeting campaign name using *sku_prefix*.

    Logic:
        • If the original name has a PT-suffix pattern → reuse it (strip trailing Nguyen).
        • Otherwise → try to extract a competitor ASIN from the targeting expression.
        • Fallback → clean the expression text directly.

    Args:
        sku_prefix: Sanitised SKU used as the name prefix.
        orig_name:  The original campaign name string.
        pt_rows:    All Product Targeting entity rows for this campaign.

    Returns:
        Fully formed SP03 campaign name string.
    """
    # Check for PT pattern in original name
    pt_match = _PT_PATTERN.search(orig_name)
    if pt_match:
        pt_body = pt_match.group(1)
        pt_body = _TRAILING_NGUYEN.sub("", pt_body).strip("_- ")
        pt_suffix = f"PT{pt_body}"
        return no_double_underscore(f"{sku_prefix}_SP03_{pt_suffix}_Nguyen")

    # Fallback: use first Product Targeting expression
    first_pt = pt_rows[0]
    expr = safe_str(first_pt.get("Product Targeting Expression", ""))

    # Skip auto-target expressions (those belong to SP02)
    if expr.lower() in _AUTO_EXPRESSIONS:
        return no_double_underscore(f"{sku_prefix}_SP02_AUTO_Nguyen")

    # Try to extract competitor ASIN from the expression
    competitor_asin = extract_asin(expr)
    if competitor_asin:
        return no_double_underscore(f"{sku_prefix}_SP03_PT_{competitor_asin}_Nguyen")

    # Clean the expression text and use it directly
    if expr:
        cleaned_expr = clean_keyword(expr)
        if cleaned_expr:
            return no_double_underscore(f"{sku_prefix}_SP03_PT_{cleaned_expr}_Nguyen")

    # Absolute fallback
    return no_double_underscore(f"{sku_prefix}_SP03_PT_Unknown_Nguyen")


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
        {SKU}_SP00_{kw}_EXACT_Nguyen        (original)
        {SKU}_SP00_{kw}_EXACT_v2_Nguyen     (first duplicate)
        {SKU}_SP00_{kw}_EXACT_v3_Nguyen     (second duplicate)
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
    print("  MODULE 2 – Campaign Naming SOP (SKU-prefix variant)")
    print("=" * 70)
    print(f"[MODULE 2] Processing file: {sku_file}")
    print(f"[MODULE 2] Inferred SKU: '{sku_name}'\n")

    df_sku = load_bulksheet(sku_file)
    rmap   = generate_rename_map(df_sku, sku_name)

    print("\n[MODULE 2] Rename Map:")
    for campaign_id, new_name in rmap.items():
        print(f"  {campaign_id}  →  {new_name}")
