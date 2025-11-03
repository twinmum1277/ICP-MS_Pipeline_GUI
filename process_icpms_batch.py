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
      - col 0 = date/time
      - col 1 = sample name
      - col 2+ = analyte columns (e.g. '63  Cu  [ He ]', '75 -> 91  As  [ O2 ]')
      - possibly a second row of 'Conc.' that we should drop
    """
    df = pd.read_csv(path)

    # rename first two columns to standard names
    df = df.rename(columns={df.columns[0]: date_col,
                            df.columns[1]: sample_col})

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

    chan_map = pd.DataFrame(meta_list)

    # rename the wide dataframe to use channel_ids
    rename_map = {row["original"]: row["channel_id"]
                  for _, row in chan_map.iterrows()}
    df_renamed = df.rename(columns=rename_map)

    long_df = df_renamed.melt(id_vars=list(meta_cols),
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
    df["df"] = df["df"].fillna(default_df)

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

def compute_icv_srm(corrected_df: pd.DataFrame,
                    icv_df: pd.DataFrame,
                    crm_values_df: pd.DataFrame = None):
    """
    Find ICV and SRM rows and calculate % recovery.
    """
    # ICV rows
    icv = corrected_df[corrected_df["sample_id"].str.contains("ICV", na=False)].copy()
    icv = icv.merge(icv_df[["element", "icv_target"]], on="element", how="left")
    icv["icv_recovery_pct"] = icv["corrected"] / icv["icv_target"] * 100

    # SRM rows (could be SRM / NIST / CRM / USGS)
    srm = corrected_df[
        corrected_df["sample_id"].str.contains("SRM|NIST|CRM|USGS", na=False)
    ].copy()

    if crm_values_df is not None:
        # use actual CRM values file if provided
        crm_small = crm_values_df.rename(columns={"target_value": "srm_target"})
        srm = srm.merge(crm_small[["element", "srm_target"]],
                        on="element", how="left")
    else:
        # fall back to ICV file
        if "srm_target" in icv_df.columns:
            srm = srm.merge(icv_df[["element", "srm_target"]],
                            on="element", how="left")
        else:
            srm["srm_target"] = np.nan

    srm["srm_recovery_pct"] = srm["corrected"] / srm["srm_target"] * 100

    return icv, srm


# ---------------------------------------------------------------------------
# Choose best channel per element
# ---------------------------------------------------------------------------

def select_best_channels(icv_df: pd.DataFrame,
                         srm_df: pd.DataFrame,
                         chan_map: pd.DataFrame,
                         icv_lo=90,
                         icv_hi=110,
                         srm_lo=80,
                         srm_hi=120) -> pd.DataFrame:
    """
    For each element:
      - look at all channels
      - prefer channel that passes BOTH ICV (90-110) and SRM (80-120)
      - else choose channel with SRM recovery closest to 100
    """
    summary_rows = []
    elements = chan_map["element"].unique()

    for el in elements:
        icv_sub = icv_df[icv_df["element"] == el]
        srm_sub = srm_df[srm_df["element"] == el]

        merged = pd.merge(
            icv_sub[["channel_id", "icv_recovery_pct"]],
            srm_sub[["channel_id", "srm_recovery_pct"]],
            on="channel_id", how="outer"
        )

        if merged.empty:
            # element may not have ICV/SRM rows
            summary_rows.append({
                "element": el,
                "selected_channel_id": None,
                "icv_recovery_pct": np.nan,
                "icv_pass": False,
                "srm_recovery_pct": np.nan,
                "srm_pass": False
            })
            continue

        merged["icv_pass"] = (
            (merged["icv_recovery_pct"] >= icv_lo) &
            (merged["icv_recovery_pct"] <= icv_hi)
        )
        merged["srm_pass"] = (
            (merged["srm_recovery_pct"] >= srm_lo) &
            (merged["srm_recovery_pct"] <= srm_hi)
        )

        # best = passes both
        both = merged[merged["icv_pass"] & merged["srm_pass"]]
        if not both.empty:
            chosen = both.iloc[0]
        else:
            # choose closest to 100% SRM
            merged["dist100"] = (merged["srm_recovery_pct"] - 100).abs()
            merged = merged.sort_values("dist100")
            chosen = merged.iloc[0]

        summary_rows.append({
            "element": el,
            "selected_channel_id": chosen["channel_id"],
            "icv_recovery_pct": chosen.get("icv_recovery_pct", np.nan),
            "icv_pass": chosen.get("icv_pass", False),
            "srm_recovery_pct": chosen.get("srm_recovery_pct", np.nan),
            "srm_pass": chosen.get("srm_pass", False)
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
                          out_path: str):
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        results_corrected.to_excel(writer, sheet_name="Corrected ppm", index=False)
        qc_summary.to_excel(writer, sheet_name="QC Summary", index=False)
        bdl_df.to_excel(writer, sheet_name="Below Detection Limit", index=False)
    print(f"[OK] Wrote workbook: {out_path}")


# ---------------------------------------------------------------------------
# Main batch runner
# ---------------------------------------------------------------------------

def process_batch(sort_path: str,
                  digest_path: str,
                  icv_path: str,
                  output_path: str,
                  crm_values_path: str = None,
                  apply_divide1000: bool = True):
    # load
    sort_df = load_sort_file(sort_path)
    digest_df = load_digest_file(digest_path)
    icv_df = load_icv_file(icv_path)
    crm_df = None
    if crm_values_path and os.path.exists(crm_values_path):
        crm_df = pd.read_csv(crm_values_path)

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
    icv_data, srm_data = compute_icv_srm(corrected_df, icv_df, crm_df)
    qc_summary_df = select_best_channels(icv_data, srm_data, chan_map)

    # samples only (not blanks/ICV/SRM)
    samples_df = corrected_df[
        ~corrected_df["sample_id"].str.contains("BLANK|ICV|SRM|DUP", na=False)
    ].copy()

    # BDL
    bdl_df = build_bdl_table(corrected_df, blank_elem_stats)

    # write Excel
    build_output_workbook(samples_df, qc_summary_df, bdl_df, output_path)


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
