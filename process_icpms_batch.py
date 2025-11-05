"""
process_icpms_batch.py

Batch processor for Agilent MassHunter ICP-MS data exported to CSV.

Pipeline:
1. Load SORT (instrument export) → normalize headers → long format.
2. Load DIGEST (sample_id, df).
3. Load ICV (icv_target, srm_target, tolerances).
4. Compute blank stats (avg blank, SD, MDL).
5. Correct sample values: (raw – avg_blank) * df / 1000  (the /1000 is optional).
6. Identify ICV and SRM rows; compute % recovery.
7. For elements with multiple channels/modes, choose the "best" channel based on QC.
8. Write Excel with 3 sheets:
   - "Corrected ppm"
   - "QC Summary"
   - "Below Detection Limit"
"""

import os
import pandas as pd
import numpy as np


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def normalize_sample_id(x: str) -> str:
    """Normalize sample names so SORT and DIGEST can match."""
    return (x or "").strip().upper().replace(" ", "_")


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def load_sort_file(path: str,
                   sample_col="sample_id",
                   date_col="acq_time") -> pd.DataFrame:
    """
    SORT = MassHunter export.
    Assumptions:
      - Has a date/time column (first column or contains 'date', 'time', 'acq')
      - Has a sample name column (contains 'sample', 'name', or second column)
      - col 2+ = analyte columns (e.g. '63  Cu  [ He ]', '75 -> 91  As  [ O2 ]')
      - possibly a second row of 'Conc.' that we should drop
    """
    # Try reading with default header (row 0)
    df_test = pd.read_csv(path, nrows=0)
    
    # Check if first row looks like it doesn't have proper column names (lots of Unnamed)
    unnamed_count = sum(1 for col in df_test.columns[:5] if str(col).startswith('Unnamed:'))
    
    if unnamed_count >= 3:
        # First row is probably element symbols, actual headers are in row 2
        print("Debug: Detected header in row 2, reading with header=1")
        df = pd.read_csv(path, header=1)
    else:
        df = pd.read_csv(path)

    # Debug: show what columns we have
    print(f"Debug SORT file - First 5 columns: {list(df.columns[:5])}")
    
    # Find date/time column
    date_col_found = None
    for col in df.columns[:5]:  # Check first 5 columns
        col_lower = str(col).lower()
        if any(keyword in col_lower for keyword in ['date', 'time', 'acq']):
            date_col_found = col
            print(f"Debug: Found date column: '{date_col_found}'")
            break
    if not date_col_found:
        date_col_found = df.columns[0]  # Default to first column
        print(f"Debug: No date column found, using first column: '{date_col_found}'")

    # Find sample name column
    sample_col_found = None
    for col in df.columns[:5]:  # Check first 5 columns
        col_lower = str(col).lower()
        if any(keyword in col_lower for keyword in ['sample', 'name']) and col != date_col_found:
            sample_col_found = col
            print(f"Debug: Found sample column: '{sample_col_found}'")
            break
    if not sample_col_found:
        # Default to second column if different from date column
        sample_col_found = df.columns[1] if df.columns[1] != date_col_found else df.columns[0]
        print(f"Debug: No sample column found, using: '{sample_col_found}'")

    # Rename to standard names
    df = df.rename(columns={date_col_found: date_col,
                            sample_col_found: sample_col})
    
    print(f"Debug: Renamed '{date_col_found}' → 'acq_time'")
    print(f"Debug: Renamed '{sample_col_found}' → 'sample_id'")

    # drop "Conc." row if present right under header
    if df.shape[0] > 1:
        second_row = df.iloc[1, 2:]
        if (second_row == "Conc.").all():
            df = df.iloc[1:].copy()

    # normalize sample_id
    df[sample_col] = df[sample_col].astype(str).apply(normalize_sample_id)
    return df


def load_digest_file(path: str,
                     sample_col="sample_id",
                     df_col="df") -> pd.DataFrame:
    """
    DIGEST = file that has dilution/digestion factor for each sample_id.
    Minimal schema: sample_id, df
    Supports CSV and Excel (.xlsx, .xls) files.
    """
    # Detect file type by extension
    if path.lower().endswith(('.xlsx', '.xls')):
        d = pd.read_excel(path)
    else:
        d = pd.read_csv(path)
    
    # standardize column names
    d = d.rename(columns={sample_col: "sample_id", df_col: "df"})
    d["sample_id"] = d["sample_id"].astype(str).apply(normalize_sample_id)
    return d[["sample_id", "df"]]


def load_icv_file(path: str) -> pd.DataFrame:
    """
    ICV = file that holds ICV target, SRM target, tolerances.
    Minimal schema: element, icv_target, srm_target
    """
    icv = pd.read_csv(path)
    
    # Debug: show what columns we have
    print(f"Debug ICV file columns: {list(icv.columns)}")
    
    # Check if required columns exist
    if "element" not in icv.columns:
        raise ValueError(f"ICV file must have 'element' column. Found columns: {list(icv.columns)}")
    
    icv["element"] = icv["element"].astype(str).str.strip()
    return icv


# ---------------------------------------------------------------------------
# Channel parsing
# ---------------------------------------------------------------------------

def parse_channel_headers(df: pd.DataFrame,
                          meta_cols=("acq_time", "sample_id")):
    """
    We turn the wide table into a long table and attach channel metadata.

    header examples:
      "63  Cu  [ He ]"
      "75 -> 91  As  [ O2 ]"

    Returns:
      long_df: rows = samples × channels
      chan_map: info about each channel
    """
    chan_cols = [c for c in df.columns if c not in meta_cols]
    meta_list = []

    for c in chan_cols:
        s = str(c).strip()
        
        # Skip unnamed columns or non-element columns
        if s.startswith("Unnamed:") or not s:
            continue

        try:
            if "->" in s:
                # mass-shifted
                # "75 -> 91  As  [ O2 ]"
                left, right = s.split("->", 1)
                nom_mass = int(left.strip().split()[0])
                right = right.strip()
                analyzed = int(right.split()[0])
                element = right.split()[1]
                gas = right.split("[")[-1].split("]")[0].strip()
                chan_id = f"{element}{nom_mass}to{analyzed}_{gas}"
                is_shift = True
            else:
                # plain mass
                # "63  Cu  [ He ]"
                parts = s.split()
                nom_mass = int(parts[0])
                element = parts[1]
                gas = parts[-1].strip("[]")
                analyzed = nom_mass
                chan_id = f"{element}{analyzed}_{gas}"
                is_shift = False

            meta_list.append({
                "original": c,
                "channel_id": chan_id,
                "element": element,
                "nominal_mass": nom_mass,
                "analyzed_mass": analyzed,
                "gas_mode": gas,
                "is_shift": is_shift
            })
        except (ValueError, IndexError) as e:
            # Skip columns that don't match expected format
            print(f"Warning: Skipping column '{c}' - doesn't match expected element format")
            continue

    chan_map = pd.DataFrame(meta_list)

    # Get list of valid columns (only those we successfully parsed)
    valid_columns = [row["original"] for _, row in chan_map.iterrows()]
    
    # Keep only metadata columns + valid element columns
    # Get actual meta columns that exist in df (in case names are different)
    actual_meta_cols = [c for c in df.columns if c in meta_cols]
    
    # Debug: show what we found
    print(f"Debug: Looking for meta_cols: {meta_cols}")
    print(f"Debug: Found in df: {actual_meta_cols}")
    print(f"Debug: All df columns: {list(df.columns[:5])}")  # Show first 5
    
    columns_to_keep = actual_meta_cols + valid_columns
    df_filtered = df[columns_to_keep].copy()
    
    # rename the wide dataframe to use channel_ids
    rename_map = {row["original"]: row["channel_id"]
                  for _, row in chan_map.iterrows()}
    df_renamed = df_filtered.rename(columns=rename_map)

    # Use actual meta columns found in the dataframe for melting
    long_df = df_renamed.melt(id_vars=actual_meta_cols,
                              var_name="channel_id",
                              value_name="raw_conc")
    long_df = long_df.merge(chan_map, on="channel_id", how="left")

    # convert to numeric
    long_df["raw_conc"] = pd.to_numeric(long_df["raw_conc"], errors="coerce")
    long_df["sample_id"] = long_df["sample_id"].astype(str).apply(
        normalize_sample_id
    )

    return long_df, chan_map


# ---------------------------------------------------------------------------
# Blank stats
# ---------------------------------------------------------------------------

def compute_blank_stats(long_df: pd.DataFrame):
    """
    Compute avg, SD, MDL per channel, and also per element.
    Blank is detected by sample_id containing 'BLANK'.
    """
    blanks = long_df[long_df["sample_id"].str.contains("BLANK", na=False)]

    stats_chan = (
        blanks.groupby("channel_id")["raw_conc"]
        .agg(["mean", "std", "count"])
        .reset_index()
        .rename(columns={"mean": "avg_blank",
                         "std": "sd_blank",
                         "count": "n_blank"})
    )
    stats_chan["mdl"] = 3 * stats_chan["sd_blank"]

    stats_elem = (
        blanks.groupby("element")["raw_conc"]
        .agg(["mean", "std"])
        .reset_index()
        .rename(columns={"mean": "avg_blank",
                         "std": "sd_blank"})
    )
    stats_elem["mdl"] = 3 * stats_elem["sd_blank"]

    return stats_chan, stats_elem


# ---------------------------------------------------------------------------
# Correction
# ---------------------------------------------------------------------------

def join_and_correct(long_df: pd.DataFrame,
                     digest_df: pd.DataFrame,
                     blank_stats_chan: pd.DataFrame,
                     blank_stats_elem: pd.DataFrame,
                     default_df=1.0,
                     apply_divide1000=True):
    """
    (ICP-MS sample value – avg blank) × df (/1000)
    """
    df = long_df.copy()

    # join DF
    df = df.merge(digest_df, on="sample_id", how="left")
    
    # Report unmatched samples and return list
    unmatched = df[df["df"].isna()]["sample_id"].unique()
    
    # Filter out QC samples (ICV, ICB, BLANK, SRM) - these don't need DIGEST data
    unmatched_samples_only = [s for s in unmatched 
                               if not any(qc in s for qc in ['ICV', 'ICB', 'BLANK', 'SRM_'])]
    
    if len(unmatched_samples_only) > 0:
        print(f"\n⚠️  Warning: {len(unmatched_samples_only)} UNKNOWN samples not found in DIGEST file (using df=1.0):")
        for s in unmatched_samples_only:
            print(f"  ❌ {s}")
        print(f"\n  Available in DIGEST: {list(digest_df['sample_id'].unique()[:5])}")  # Show first 5
    
    df["df"] = df["df"].fillna(default_df)
    
    # Store only unknown samples in metadata (not QC samples)
    df.attrs['unmatched_samples'] = unmatched_samples_only

    # join channel blank
    df = df.merge(blank_stats_chan[["channel_id", "avg_blank"]],
                  on="channel_id", how="left")
    df["avg_blank"] = df["avg_blank"].fillna(0.0)

    # corrected
    df["corrected"] = (df["raw_conc"] - df["avg_blank"]) * df["df"]
    if apply_divide1000:
        df["corrected"] = df["corrected"] / 1000.0

    # clamp negatives
    df.loc[df["corrected"] < 0, "corrected"] = 0

    return df


# ---------------------------------------------------------------------------
# ICV / SRM calculations
# ---------------------------------------------------------------------------

def compute_icv_ref(corrected_df: pd.DataFrame,
                    icv_df: pd.DataFrame,
                    ref_values_df: pd.DataFrame = None):
    """
    Find ICV and REF rows and calculate % recovery.
    """
    # ICV rows
    icv = corrected_df[corrected_df["sample_id"].str.contains("ICV", na=False)].copy()
    
    # Check if ICV file has required columns
    if "element" not in icv_df.columns or "icv_target" not in icv_df.columns:
        print(f"Warning: ICV file columns: {list(icv_df.columns)}")
        print(f"Warning: Expected 'element' and 'icv_target'")
    
    icv = icv.merge(icv_df[["element", "icv_target"]], on="element", how="left")
    icv["icv_recovery_pct"] = icv["corrected"] / icv["icv_target"] * 100

    # REF rows (must start with SRM_ - will rename to REF_ in sample naming convention)
    ref = corrected_df[
        corrected_df["sample_id"].str.startswith("SRM_", na=False)
    ].copy()

    # Only process REF if we have REF samples
    if len(ref) > 0:
        if ref_values_df is not None:
            # Extract REF identifier from sample name
            # Expected format: SRM_[NAME]_[#] (using SRM_ for backward compatibility)
            # e.g., "SRM_DOLT-5_1" -> "DOLT-5"
            #       "SRM_NIST_2710_1" -> "NIST_2710"
            #       "SRM_DORM-4_2" -> "DORM-4"
            def extract_ref_name(sample_id):
                # Split by underscore and get the middle part(s)
                # SRM_DOLT-5_1 -> ['SRM', 'DOLT-5', '1']
                parts = sample_id.split('_')
                if len(parts) >= 3 and parts[0] == 'SRM':
                    # Join all middle parts (in case REF name has underscores like NIST_2710)
                    return '_'.join(parts[1:-1])
                return None
            
            ref["ref_name"] = ref["sample_id"].apply(extract_ref_name)
            
            # Debug: show extracted REF names
            print(f"Debug: Extracted REF names: {ref[['sample_id', 'ref_name']].drop_duplicates().values.tolist()}")
            
            # Check if REF file has ref_name column
            if "ref_name" in ref_values_df.columns:
                # Match on both ref_name and element
                ref_small = ref_values_df.rename(columns={"target_value": "ref_target"})
                
                print(f"Debug: REF merge - sample REF data:")
                print(ref_small[ref_small['ref_name'] == 'DOLT-5'].head(10))
                
                ref = ref.merge(ref_small[["ref_name", "element", "ref_target"]],
                                on=["ref_name", "element"], how="left")
                
                print(f"Debug: After merge - sample REF rows with targets:")
                print(ref[ref['ref_target'].notna()][['sample_id', 'element', 'corrected', 'ref_target']].head(10))
            else:
                # Old format: just element and target_value (assume all samples are same REF)
                ref_small = ref_values_df.rename(columns={"target_value": "ref_target"})
                ref = ref.merge(ref_small[["element", "ref_target"]],
                                on="element", how="left")
        else:
            # fall back to ICV file (if it has ref_target column)
            if "ref_target" in icv_df.columns:
                ref = ref.merge(icv_df[["element", "ref_target"]],
                                on="element", how="left")
            else:
                # No REF targets available
                ref["ref_target"] = np.nan
        
        ref["ref_recovery_pct"] = ref["corrected"] / ref["ref_target"] * 100
        
        # Debug: show REF data
        print(f"\nDebug REF Recovery:")
        print(f"  Total REF rows: {len(ref)}")
        print(f"  Unique samples: {ref['sample_id'].unique()}")
        print(f"  Unique elements: {ref['element'].unique()}")
        print(f"  Has ref_target?: {ref['ref_target'].notna().sum()} of {len(ref)}")
        print(f"  Sample recovery %:\n{ref[['sample_id', 'element', 'channel_id', 'ref_target', 'corrected', 'ref_recovery_pct']].head(10)}")
    else:
        # No REF samples, create empty dataframe with correct structure
        ref["ref_target"] = np.nan
        ref["ref_recovery_pct"] = np.nan

    return icv, ref


# ---------------------------------------------------------------------------
# Choose best channel per element
# ---------------------------------------------------------------------------

def select_best_channels(icv_df: pd.DataFrame,
                         ref_df: pd.DataFrame,
                         chan_map: pd.DataFrame,
                         icv_lo=90,
                         icv_hi=110,
                         ref_lo=80,
                         ref_hi=120) -> pd.DataFrame:
    """
    For each element:
      - look at all channels
      - prefer channel that passes BOTH ICV (90-110) and REF (80-120)
      - else choose channel with REF recovery closest to 100
    """
    summary_rows = []
    elements = chan_map["element"].unique()

    for el in elements:
        icv_sub = icv_df[icv_df["element"] == el]
        ref_sub = ref_df[ref_df["element"] == el]

        merged = pd.merge(
            icv_sub[["channel_id", "icv_recovery_pct"]],
            ref_sub[["channel_id", "ref_recovery_pct"]],
            on="channel_id", how="outer"
        )

        if merged.empty:
            # element may not have ICV/REF rows
            summary_rows.append({
                "element": el,
                "selected_channel_id": None,
                "icv_recovery_pct": np.nan,
                "icv_pass": False,
                "ref_recovery_pct": np.nan,
                "ref_pass": False
            })
            continue

        merged["icv_pass"] = (
            (merged["icv_recovery_pct"] >= icv_lo) &
            (merged["icv_recovery_pct"] <= icv_hi)
        )
        merged["ref_pass"] = (
            (merged["ref_recovery_pct"] >= ref_lo) &
            (merged["ref_recovery_pct"] <= ref_hi)
        )

        # best = passes both
        both = merged[merged["icv_pass"] & merged["ref_pass"]]
        if not both.empty:
            chosen = both.iloc[0]
        else:
            # choose closest to 100% REF
            merged["dist100"] = (merged["ref_recovery_pct"] - 100).abs()
            merged = merged.sort_values("dist100")
            chosen = merged.iloc[0]

        summary_rows.append({
            "element": el,
            "selected_channel_id": chosen["channel_id"],
            "icv_recovery_pct": chosen.get("icv_recovery_pct", np.nan),
            "icv_pass": chosen.get("icv_pass", False),
            "ref_recovery_pct": chosen.get("ref_recovery_pct", np.nan),
            "ref_pass": chosen.get("ref_pass", False)
        })

    return pd.DataFrame(summary_rows)


# ---------------------------------------------------------------------------
# Build BDL sheet
# ---------------------------------------------------------------------------

def build_bdl_table(corrected_df: pd.DataFrame,
                    blank_elem_stats: pd.DataFrame) -> pd.DataFrame:
    """
    Below detection: (raw - avg_blank) < mdl
    """
    df = corrected_df.merge(
        blank_elem_stats[["element", "mdl"]],
        on="element", how="left"
    )
    cond = (df["raw_conc"] - df["avg_blank"]) < df["mdl"]
    bdl = df[cond].copy()
    return bdl[["sample_id", "element", "channel_id",
                "raw_conc", "avg_blank", "mdl"]]


# ---------------------------------------------------------------------------
# Write Excel
# ---------------------------------------------------------------------------

def build_output_workbook(results_corrected: pd.DataFrame,
                          qc_summary: pd.DataFrame,
                          bdl_df: pd.DataFrame,
                          blank_elem_stats: pd.DataFrame,
                          out_path: str,
                          unmatched_samples: list = None):
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # Define number formats
        ppm_format = workbook.add_format({'num_format': '0.000'})  # 3 decimal places
        percent_format = workbook.add_format({'num_format': '0"%"'})  # 0 decimal places
        percent_fail_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0"%"'})  # Red for failed QC
        
        # Create wide-format pivot table
        if 'element' in results_corrected.columns and 'corrected' in results_corrected.columns:
            # Select best channel per element-sample combination
            pivot_data = results_corrected.sort_values('channel_id').drop_duplicates(['sample_id', 'element'], keep='first')
            wide_df = pivot_data.pivot(index='sample_id', columns='element', values='corrected')
            wide_df = wide_df.reset_index()
            
            # Round to 3 decimal places
            for col in wide_df.columns[1:]:  # Skip sample_id column
                wide_df[col] = wide_df[col].round(3)
            
            # Write wide format sheet
            wide_df.to_excel(writer, sheet_name="Corrected ppm (Wide)", index=False)
            
            # Apply conditional formatting for BDL and unmatched samples
            worksheet = writer.sheets["Corrected ppm (Wide)"]
            
            # Format for BDL values (red fill, 3 decimals)
            bdl_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0.000'})
            
            # Format for unmatched samples (yellow fill, red text, 3 decimals)
            unmatched_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C0006', 'bold': True, 'num_format': '0.000'})
            
            # Get MDL values per element
            mdl_dict = dict(zip(blank_elem_stats['element'], blank_elem_stats['mdl']))
            
            # Apply formatting
            for row_idx, sample_id in enumerate(wide_df['sample_id'], start=1):  # start=1 for header
                # Check if sample is unmatched
                is_unmatched = unmatched_samples and sample_id in unmatched_samples
                
                if is_unmatched:
                    # Highlight entire row for unmatched samples
                    worksheet.set_row(row_idx, None, unmatched_format)
                else:
                    # Check each element for BDL
                    for col_idx, element in enumerate(wide_df.columns[1:], start=1):  # skip sample_id column
                        value = wide_df.iloc[row_idx-1][element]
                        mdl = mdl_dict.get(element, float('inf'))
                        
                        if pd.notna(value) and value < mdl:
                            worksheet.write(row_idx, col_idx, value, bdl_format)
        
        # Write other sheets
        results_corrected.to_excel(writer, sheet_name="Corrected ppm (Long)", index=False)
        qc_summary.to_excel(writer, sheet_name="QC Summary", index=False)
        bdl_df.to_excel(writer, sheet_name="Below Detection Limit", index=False)
        
        # Format QC Summary sheet - apply percentage formatting and conditional formatting
        qc_worksheet = writer.sheets["QC Summary"]
        
        # Find recovery and pass columns in QC Summary
        if 'icv_recovery_pct' in qc_summary.columns:
            icv_col_idx = qc_summary.columns.get_loc('icv_recovery_pct')
            qc_worksheet.set_column(icv_col_idx, icv_col_idx, 15, percent_format)
            
            # Apply conditional formatting for failed ICV (if icv_pass column exists)
            if 'icv_pass' in qc_summary.columns:
                icv_pass_col_idx = qc_summary.columns.get_loc('icv_pass')
                for row_idx in range(1, len(qc_summary) + 1):  # Start at 1 (skip header)
                    if row_idx <= len(qc_summary) and not qc_summary.iloc[row_idx - 1]['icv_pass']:
                        val = qc_summary.iloc[row_idx - 1]['icv_recovery_pct']
                        if pd.notna(val):
                            qc_worksheet.write(row_idx, icv_col_idx, val, percent_fail_format)
        
        if 'ref_recovery_pct' in qc_summary.columns:
            ref_col_idx = qc_summary.columns.get_loc('ref_recovery_pct')
            qc_worksheet.set_column(ref_col_idx, ref_col_idx, 15, percent_format)
            
            # Apply conditional formatting for failed REF (if ref_pass column exists)
            if 'ref_pass' in qc_summary.columns:
                ref_pass_col_idx = qc_summary.columns.get_loc('ref_pass')
                for row_idx in range(1, len(qc_summary) + 1):  # Start at 1 (skip header)
                    if row_idx <= len(qc_summary) and not qc_summary.iloc[row_idx - 1]['ref_pass']:
                        val = qc_summary.iloc[row_idx - 1]['ref_recovery_pct']
                        if pd.notna(val):
                            qc_worksheet.write(row_idx, ref_col_idx, val, percent_fail_format)
        
        # Auto-size columns (but keep Wide sheet columns narrower)
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            
            # Get the dataframe for this sheet to calculate widths
            if sheet_name == "Corrected ppm (Wide)":
                df_to_size = wide_df if 'wide_df' in locals() else None
                # For wide format, use fixed narrow width for element columns
                if df_to_size is not None:
                    # First column (sample_id) gets auto-sized
                    max_len = max(len('sample_id'), df_to_size['sample_id'].astype(str).str.len().max() if len(df_to_size) > 0 else 0)
                    worksheet.set_column(0, 0, min(max_len + 2, 30))
                    # Element columns get fixed narrow width with ppm format
                    for idx in range(1, len(df_to_size.columns)):
                        worksheet.set_column(idx, idx, 10, ppm_format)  # Fixed width + 3 decimal format
            elif sheet_name == "Corrected ppm (Long)":
                df_to_size = results_corrected
            elif sheet_name == "QC Summary":
                df_to_size = qc_summary
            elif sheet_name == "Below Detection Limit":
                df_to_size = bdl_df
            else:
                df_to_size = None
            
            # Auto-size for non-Wide sheets
            if df_to_size is not None and sheet_name != "Corrected ppm (Wide)":
                for idx, col in enumerate(df_to_size.columns):
                    # Calculate max width: max of column name and column values
                    max_len = max(
                        len(str(col)),
                        df_to_size[col].astype(str).str.len().max() if len(df_to_size) > 0 else 0
                    )
                    # Add some padding and cap at reasonable width
                    adjusted_width = min(max_len + 2, 50)
                    worksheet.set_column(idx, idx, adjusted_width)
        
    print(f"[OK] Wrote workbook: {out_path}")


# ---------------------------------------------------------------------------
# Main batch runner
# ---------------------------------------------------------------------------

def load_ref_file(path: str) -> pd.DataFrame:
    """
    Load REF (reference material) values file. Supports two formats:
    
    Format 1 (long): ref_name, element, target_value
    Format 2 (wide): Row 1 = element symbols, Row 2 = names, Row 3 = units, 
                     Row 4+ = REF_name in column 2, values across columns
    """
    df = pd.read_csv(path)
    
    # Check if it's wide format (has many columns, first few might be empty or metadata)
    if df.shape[1] > 10:  # Likely wide format with many elements
        # Read first row WITHOUT using it as header to get element SYMBOLS
        element_row = pd.read_csv(path, header=None, nrows=1)
        element_symbols = list(element_row.iloc[0, 2:])  # Skip first 2 columns, get symbols
        
        # Skip first 3 rows (symbols, names, units) and read data
        df = pd.read_csv(path, skiprows=3, header=None)
        
        # Column 1 should be REF name, columns 2+ are element values
        ref_name_col = df.columns[1]  # Second column is REF name
        value_cols = df.columns[2:]    # Rest are element values
        
        long_rows = []
        for _, row in df.iterrows():
            ref_name = row[ref_name_col]
            if pd.isna(ref_name) or not str(ref_name).strip():
                continue
            for i, col in enumerate(value_cols):
                if i < len(element_symbols):
                    element_symbol = str(element_symbols[i]).strip()
                    value = row[col]
                    # Only add if element symbol exists and value is valid
                    if element_symbol and not pd.isna(value) and value != '':
                        try:
                            long_rows.append({
                                'ref_name': str(ref_name).strip(),
                                'element': element_symbol,  # Use symbol, not full name
                                'target_value': float(value)
                            })
                        except (ValueError, TypeError):
                            continue
        
        ref_df = pd.DataFrame(long_rows)
        
        # Convert target values from mg/kg (ppm) to µg/kg (ppb)
        # Assuming REF file values are in mg/kg
        ref_df['target_value'] = ref_df['target_value'] * 1000
        
        print(f"Debug: Converted wide-format REF file to long format ({len(ref_df)} rows)")
        print(f"Debug REF: Unique ref_names: {ref_df['ref_name'].unique()}")
        print(f"Debug REF: Sample rows (converted to ppb):\n{ref_df.head(10)}")
        return ref_df
    else:
        # Already in long format
        return df


def process_batch(sort_path: str,
                  digest_path: str,
                  icv_path: str,
                  output_path: str,
                  ref_values_path: str = None,
                  apply_divide1000: bool = True):
    # load
    sort_df = load_sort_file(sort_path)
    digest_df = load_digest_file(digest_path)
    icv_df = load_icv_file(icv_path)
    ref_df = None
    if ref_values_path and os.path.exists(ref_values_path):
        ref_df = load_ref_file(ref_values_path)

    # parse
    long_df, chan_map = parse_channel_headers(sort_df)

    # blanks
    blank_chan_stats, blank_elem_stats = compute_blank_stats(long_df)

    # correct
    corrected_df = join_and_correct(
        long_df,
        digest_df,
        blank_chan_stats,
        blank_elem_stats,
        default_df=1.0,
        apply_divide1000=apply_divide1000
    )

    # QC
    icv_data, ref_data = compute_icv_ref(corrected_df, icv_df, ref_df)
    qc_summary_df = select_best_channels(icv_data, ref_data, chan_map)

    # samples only (not blanks/ICV/SRM/DUP)
    # Exclude anything starting with BLANK_, ICV_, SRM_, or containing DUP
    samples_df = corrected_df[
        ~(corrected_df["sample_id"].str.contains("BLANK|ICV|DUP", na=False, case=False) |
          corrected_df["sample_id"].str.startswith("SRM_", na=False))
    ].copy()

    # BDL
    bdl_df = build_bdl_table(corrected_df, blank_elem_stats)
    
    # Get unmatched samples list
    unmatched_samples = getattr(corrected_df, 'attrs', {}).get('unmatched_samples', [])

    # write Excel
    build_output_workbook(samples_df, qc_summary_df, bdl_df, blank_elem_stats, output_path, unmatched_samples)
    
    # Return summary statistics for GUI
    stats = {
        'unmatched_samples': unmatched_samples,
        'total_samples': len(samples_df['sample_id'].unique()),
        'total_icv': len(icv_data['sample_id'].unique()) if len(icv_data) > 0 else 0,
        'total_ref': len(ref_data['sample_id'].unique()) if len(ref_data) > 0 else 0,
        'total_blanks': len(corrected_df[corrected_df['sample_id'].str.contains('BLANK', na=False)]['sample_id'].unique()),
        'icv_pass_rate': (qc_summary_df['icv_pass'].sum() / len(qc_summary_df) * 100) if len(qc_summary_df) > 0 else 0,
        'ref_pass_rate': (qc_summary_df['ref_pass'].sum() / len(qc_summary_df) * 100) if len(qc_summary_df) > 0 else 0,
        'elements_analyzed': len(qc_summary_df),
    }
    
    return stats


# Allow running from command line if you want
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Process ICP-MS batch.")
    parser.add_argument("--sort", required=True, help="Path to sort.csv")
    parser.add_argument("--digest", required=True, help="Path to digest.csv")
    parser.add_argument("--icv", required=True, help="Path to icv.csv")
    parser.add_argument("--crm", required=False, help="Optional CRM values csv")
    parser.add_argument("--out", required=True, help="Output Excel path")
    parser.add_argument("--no-div1000", action="store_true", help="Do NOT divide by 1000 at the end")
    args = parser.parse_args()

    process_batch(
        args.sort,
        args.digest,
        args.icv,
        args.out,
        crm_values_path=args.crm,
        apply_divide1000=not args.no_div1000
    )
