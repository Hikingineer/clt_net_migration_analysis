import os
import pandas as pd

APPEND_CSV = "data_append.csv"
COMPLETED_CSV = "completed_dates.csv"
OUT_XLSX = "daily_sales_history_2011_2025_merged.xlsx"

DATA_SHEET = "Data"
CAPPED_SHEET = "CappedDates"
SUMMARY_SHEET = "Summary"
PARCEL_SUMM_SHEET = "Parcel_Summ"

SALEDATE_COL = "SaleDate"
APPRAISED_COL = "Appraised Value"
OWNERS_COL = "Owners"

EXPECTED_MERGED_COLS = ["Address", "Appraised Value", "Parcel", "Owners", "Website"]  # original 5


def _coerce_money_numeric(series: pd.Series) -> pd.Series:
    s = series.astype(str)
    s = s.str.replace(r"[$,]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def owners_has_trust(owners) -> bool:
    if owners is None or pd.isna(owners):
        return False
    return "TRUST" in str(owners).upper()


def main():
    if not os.path.exists(APPEND_CSV):
        raise RuntimeError(f"Missing {APPEND_CSV}. Run the downloader first.")

    data = pd.read_csv(APPEND_CSV).drop_duplicates()

    required = {SALEDATE_COL, APPRAISED_COL, "Parcel"}
    missing = required - set(data.columns)
    if missing:
        raise RuntimeError(f"{APPEND_CSV} missing required columns: {sorted(missing)}")

    # Parse SaleDate / Year
    data[SALEDATE_COL] = pd.to_datetime(data[SALEDATE_COL], errors="coerce")
    data["SaleYear"] = data[SALEDATE_COL].dt.year

    # TRUST flag (kept from your prior work)
    if OWNERS_COL in data.columns:
        data["IsTrust"] = data[OWNERS_COL].apply(owners_has_trust)
    else:
        data["IsTrust"] = False

    # Appraised numeric
    data["AppraisedValueNumeric"] = _coerce_money_numeric(data[APPRAISED_COL])

    # ---------------- Bucket transactions by appraised value ----------------
    # Buckets requested:
    # <150k, 150k-499k, 500k-999k, 1M to <10M, >10M
    bins = [-float("inf"), 150_000, 499_999.999999, 999_999.999999, 10_000_000, float("inf")]
    labels = ["lt150k", "150k-499k", "500k-999k", "1M-<10M", ">10M"]

    # Use NA if AppraisedValueNumeric is missing
    data["AppraisedBucket"] = pd.cut(
        data["AppraisedValueNumeric"],
        bins=bins,
        labels=labels,
        include_lowest=True,
        right=True,
    )
    # Ensure missing bucket stays missing (so it won't count)
    # (pd.cut already yields NaN for NaN inputs)

    # ---------------- Yearly Summary (existing logic) ----------------
    def min_nonzero(s: pd.Series):
        s = s.dropna()
        nz = s[s != 0]
        return nz.min() if len(nz) else pd.NA

    summary = (
        data.groupby("SaleYear", dropna=False)
        .agg(
            DeedsTransactions=("SaleYear", "size"),
            TotalAppraisedValue=("AppraisedValueNumeric", "sum"),
            AvgAppraisedValue_per_deed=("AppraisedValueNumeric", "mean"),
            MinAppraisedValue_nonzero=("AppraisedValueNumeric", min_nonzero),
            MaxAppraisedValue=("AppraisedValueNumeric", "max"),
            TrustDeedsCount=("IsTrust", "sum"),
            TrustTotalAppraisedValue=(
                "AppraisedValueNumeric",
                lambda s: s[data.loc[s.index, "IsTrust"]].sum(),
            ),
        )
        .reset_index()
    )

    summary["MinAppraisedValueUsed"] = summary["MinAppraisedValue_nonzero"]
    summary["TotalAppraisedValueMM"] = summary["TotalAppraisedValue"] / 1_000_000
    summary["TrustPctOfDeeds"] = summary["TrustDeedsCount"] / summary["DeedsTransactions"]
    summary["TrustAvgAppraisedValue_per_deed"] = (
        summary["TrustTotalAppraisedValue"] / summary["TrustDeedsCount"]
    )
    summary["AvgAppraisedValue_per_deed"] = summary["AvgAppraisedValue_per_deed"]
    summary = summary.sort_values("SaleYear")

    # ---------------- Add bucket stats to Summary sheet ----------------
    # Counts per year per bucket
    bucket_counts = (
        data.groupby(["SaleYear", "AppraisedBucket"], dropna=True)
        .size()
        .unstack(fill_value=0)
    )

    # Total appraised value per year per bucket (optional but usually useful)
    bucket_totals = (
        data.groupby(["SaleYear", "AppraisedBucket"], dropna=True)["AppraisedValueNumeric"]
        .sum()
        .unstack(fill_value=0)
    )

    # Merge into summary and create clear column names
    # Count columns
    for lab in labels:
        if lab in bucket_counts.columns:
            summary[f"TxCount_{lab}"] = bucket_counts[lab].values
        else:
            summary[f"TxCount_{lab}"] = 0

    # Value columns (in raw dollars; also in MM if you want)
    for lab in labels:
        if lab in bucket_totals.columns:
            summary[f"TxValue_{lab}"] = bucket_totals[lab].values
            summary[f"TxValueMM_{lab}"] = (bucket_totals[lab].values / 1_000_000)
        else:
            summary[f"TxValue_{lab}"] = 0
            summary[f"TxValueMM_{lab}"] = 0

    # ---------------- capped sheet ----------------
    if os.path.exists(COMPLETED_CSV):
        ledger = pd.read_csv(COMPLETED_CSV)
        capped = (
            ledger[ledger["Capped"] == True][["SaleDate", "Rows"]]
            .drop_duplicates()
            .sort_values("SaleDate")
        )
    else:
        capped = pd.DataFrame(columns=["SaleDate", "Rows"])

    # ---------------- Parcel_Summ ----------------
    parcel_stats = (
        data.groupby("Parcel", dropna=False)
        .agg(
            UniqueSaleYears=("SaleYear", lambda s: s.dropna().nunique()),
            DeedsCount=("Parcel", "size"),
            AvgAppraisedValue=("AppraisedValueNumeric", "mean"),
            MinAppraisedValue=("AppraisedValueNumeric", "min"),
            MaxAppraisedValue=("AppraisedValueNumeric", "max"),
        )
        .reset_index()
    )

    def parcel_rate_of_change(g: pd.DataFrame):
        g2 = g.dropna(subset=["SaleYear", "AppraisedValueNumeric"])
        if len(g2) < 2:
            return pd.NA
        years = g2["SaleYear"].astype(float)
        if years.nunique() < 2:
            return pd.NA
        g_first = g2.sort_values("SaleYear").iloc[0]
        g_last = g2.sort_values("SaleYear").iloc[-1]
        dy = g_last["AppraisedValueNumeric"] - g_first["AppraisedValueNumeric"]
        dt = g_last["SaleYear"] - g_first["SaleYear"]
        if dt == 0 or pd.isna(dt):
            return pd.NA
        return dy / dt

    parcel_rate = (
        data.groupby("Parcel", dropna=False)
        .apply(parcel_rate_of_change)
        .reset_index(name="AppraisedValueIncreasePerYear")
    )

    parcel_summ = parcel_stats.merge(parcel_rate, on="Parcel", how="left")
    parcel_summ["HasMultipleDeeds"] = parcel_summ["DeedsCount"] > 1

    # ---------------- Write workbook ----------------
    missing_merge_cols = [c for c in EXPECTED_MERGED_COLS if c not in data.columns]
    if missing_merge_cols:
        raise RuntimeError(f"{APPEND_CSV} missing expected merged columns: {missing_merge_cols}")

    merged_raw = data[EXPECTED_MERGED_COLS].copy()
    merged_raw[SALEDATE_COL] = data[SALEDATE_COL]
    merged_raw["SaleYear"] = data["SaleYear"]
    merged_raw["IsTrust"] = data["IsTrust"]

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        merged_raw.to_excel(writer, sheet_name=DATA_SHEET, index=False)
        capped.to_excel(writer, sheet_name=CAPPED_SHEET, index=False)
        summary.to_excel(writer, sheet_name=SUMMARY_SHEET, index=False)
        parcel_summ.to_excel(writer, sheet_name=PARCEL_SUMM_SHEET, index=False)

    print(f"Done. Wrote: {OUT_XLSX}")
    print(f"Merged raw rows: {len(merged_raw)}")
    print(f"Parcel rows: {len(parcel_summ)}")


if __name__ == "__main__":
    main()
