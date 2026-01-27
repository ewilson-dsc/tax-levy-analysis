# ============================================================
# Mammoth USD Parcel Tax Automation Script
# Version: 1.0
# Author: Ethan W
# ============================================================

# =============================================================================
# Mammoth USD Parcel Tax Automation Script
#
# PURPOSE
# -------
# This script automates the annual process of preparing Mammoth Unified School
# District’s parcel tax levy file for submission to Mono County.
#
# Mammoth USD collects a flat, voter-approved parcel tax each year. The county
# does NOT determine which parcels should be charged this tax. Instead, the
# district must provide the county with a list of parcel APNs that are subject
# to the tax, in a very specific file format.
#
# This script follows the official “Mammoth Unified School District Parcel Tax
# Levy Processing” manual (Version 2.0, January 2026) and implements the same
# steps in code.
#
#
# HIGH-LEVEL OVERVIEW
# -------------------
# The county provides a large spreadsheet each year (the “secured roll”) that
# contains ALL parcels in the county. This script gradually narrows that list
# down to only the parcels that should actually be charged the Mammoth USD
# parcel tax.



# ----------------------
# CONFIGURATION SECTION
# ----------------------

CURRENT_YEAR = 2026
PRIOR_YEAR = 2025

PARCEL_TAX_AMOUNT = 59.00
TAX_CODE = "64504"

# Update from Board resolution Exhibit A (manual Step 0)
MAMMOTH_TRAS = [
    "010-000", "010-001", "010-002", "010-003", "010-004", "010-005", "010-006", "010-007",
    "010-008", "010-009", "010-010", "010-011", "010-012", "010-013", "010-014", "010-015",
    "059-000", "059-005", "059-007", "059-012", "059-018",
]

# Input files (manual naming convention)
SECURED_CURRENT_FILENAME = f"601_Secured_{CURRENT_YEAR}.xlsx"                   # County secured roll for the current year (all parcels, all TRAs)
SECURED_PRIOR_FILENAME = f"601_Secured_{PRIOR_YEAR}.xlsx"                       # Prior year county secured roll (used only for reconciliation)
NONTAXABLE_PRIOR_FILENAME = f"Mammoth_USD_NonTaxable_APNs_{PRIOR_YEAR}.xlsx"    # APNs permanently exempt from parcel tax
SENIOR_PRIOR_FILENAME = f"Mammoth_USD_Senior_Exemptions_{PRIOR_YEAR}.xlsx"      # APNs exempt due to approved senior status

OUTPUT_DIRNAME = "output"
# output/Mammoth_Universe_[YEAR].xlsx                                           # All parcels inside Mammoth USD before exemptions (Stage 1)
# output/Mammoth_After_NonTaxable_[YEAR].xlsx                                   # Universe minus permanently non-taxable parcels (Stage 2A)
# output/Mammoth_After_Senior_[YEAR].xlsx                                       # Parcels remaining after senior exemptions (Stage 2B)
# output/Mammoth_Reconciliation_[YEAR].xlsx                                     # Year-over-year parcel comparison and sanity checks (Stage 3)
# output/64504_Mammoth_USD_[DATE].txt                                           # Final county submission file (Stage 4)

# ----------------------
# IMPORTS
# ----------------------

import pandas as pd
import re
from pathlib import Path
from datetime import datetime

# ----------------------
# STAGE 0: TRA VERIFICATION
# ----------------------

def stage0_tra_verification_prompt():
    print("=" * 50)
    print("MAMMOTH USD PARCEL TAX PROCESSING")
    print(f"Processing Year: {CURRENT_YEAR}")
    print(f"Comparison Year: {PRIOR_YEAR}")
    print("=" * 50)
    print("CRITICAL: TRA VERIFICATION REQUIRED")
    print("Verify TRAs against the current year Board resolution Exhibit A.")
    print("If TRAs changed, update MAMMOTH_TRAS in this script before proceeding.\n")
    print("TRAs currently in this script:")
    for tra in MAMMOTH_TRAS:
        print(f"  {tra}")
    input("\nPress Enter to continue once TRAs are verified...")

# =========================
# STAGE 1: EXTRACT UNIVERSE (baseline)
# =========================

def stage1_extract_universe():
    print("\nSTAGE 1: EXTRACTING MAMMOTH USD UNIVERSE")

    base_dir = Path(__file__).resolve().parent

    # Make sure output folder exists
    out_dir = base_dir / OUTPUT_DIRNAME
    out_dir.mkdir(exist_ok=True)

    # Load the current year county secured roll
    secured_path = base_dir / SECURED_CURRENT_FILENAME
    if not secured_path.exists():
        raise FileNotFoundError(f"Missing file: {SECURED_CURRENT_FILENAME} (put it in the same folder as this script)")

    df = pd.read_excel(secured_path)

    # Assumed county columns (change these if your file uses different headers)
    APN_COL = "APN"
    TRA_COL = "TRA"

    if APN_COL not in df.columns or TRA_COL not in df.columns:
        raise KeyError(
            f"Expected columns '{APN_COL}' and '{TRA_COL}' not found.\n"
            f"Found headers: {list(df.columns)}\n"
            "Update APN_COL / TRA_COL to match your file."
        )

    # Filter to Mammoth TRAs (this creates the Mammoth USD 'Universe')
    universe = df[df[TRA_COL].isin(MAMMOTH_TRAS)].copy()

    print(f"Universe parcels found: {len(universe)}")

    # Save the universe file (manual output)
    out_path = out_dir / f"Mammoth_Universe_{CURRENT_YEAR}.xlsx"
    universe.to_excel(out_path, index=False)
    print(f"Saved: output/{out_path.name}")

    return universe

# =========================
# STAGE 2A: NON-TAXABLE PARCELS (baseline)
# =========================

def stage2a_remove_nontaxable(universe):
    print("\nSTAGE 2A: REMOVING NON-TAXABLE PARCELS")

    base_dir = Path(__file__).resolve().parent
    out_dir = base_dir / OUTPUT_DIRNAME
    out_dir.mkdir(exist_ok=True)

    # Load prior year non-taxable list
    nontax_path = base_dir / NONTAXABLE_PRIOR_FILENAME
    if not nontax_path.exists():
        raise FileNotFoundError(
            f"Missing file: {NONTAXABLE_PRIOR_FILENAME} (put it in the same folder as this script)"
        )

    nontax_df = pd.read_excel(nontax_path)

    # Assumed APN column in the non-taxable list (change if needed)
    NONTAX_APN_COL = "APN"
    if NONTAX_APN_COL not in nontax_df.columns:
        raise KeyError(
            f"Expected column '{NONTAX_APN_COL}' not found in {NONTAXABLE_PRIOR_FILENAME}.\n"
            f"Found headers: {list(nontax_df.columns)}\n"
            "Update NONTAX_APN_COL to match your file."
        )

    # Assumed APN column in the universe (change if needed)
    UNIVERSE_APN_COL = "APN"
    if UNIVERSE_APN_COL not in universe.columns:
        raise KeyError(
            f"Expected column '{UNIVERSE_APN_COL}' not found in Universe.\n"
            "Update UNIVERSE_APN_COL to match your universe APN column."
        )

    # Remove any universe rows whose APN appears in the non-taxable list
    nontax_apns = set(nontax_df[NONTAX_APN_COL].dropna().astype(str))
    remaining = universe[~universe[UNIVERSE_APN_COL].astype(str).isin(nontax_apns)].copy()

    removed_count = len(universe) - len(remaining)
    print(f"Non-taxable APNs loaded: {len(nontax_apns)}")
    print(f"Parcels removed as non-taxable: {removed_count}")
    print(f"Remaining after Stage 2A: {len(remaining)}")

    # Save output for next stage
    out_path = out_dir / f"Mammoth_After_NonTaxable_{CURRENT_YEAR}.xlsx"
    remaining.to_excel(out_path, index=False)
    print(f"Saved: output/{out_path.name}")

    return remaining


# ----------------------
# MAIN
# ----------------------

if __name__ == "__main__":
    stage0_tra_verification_prompt()
    universe = stage1_extract_universe()
    print("\nStage 1 complete.")
    after_nontax = stage2a_remove_nontaxable(universe)
    print("\nStage 2A complete.")

