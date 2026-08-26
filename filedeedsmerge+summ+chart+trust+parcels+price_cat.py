import os
import pandas as pd
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList


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

EXPECTED_MERGED_COLS = [
    "Address",
    "Appraised Value",
    "Parcel",
    "Owners",
    "Website",
]


# Bucket format:
# (column suffix, lower bound, upper bound)
#
# Buckets use lower-inclusive, upper-exclusive ranges:
#   [lower bound, upper bound)
#
# Examples:
#   $150,000 belongs to 150k-250k
#   $1,000,000 belongs to 1M-2M
#   $10,000,000 belongs to 10M-plus
VALUE_BUCKETS = [
    ("lt100k", -float("inf"), 100_000),
    ("100k-150k", 100_000, 150_000),
    ("150k-250k", 150_000, 250_000),
    ("250k-350k", 250_000, 350_000),
    ("350k-500k", 350_000, 500_000),
    ("500k-750k", 500_000, 750_000),
    ("750k-1M", 750_000, 1_000_000),
    ("1M-2M", 1_000_000, 2_000_000),
    ("2M-5M", 2_000_000, 5_000_000),
    ("5M-10M", 5_000_000, 10_000_000),
    ("10M-plus", 10_000_000, float("inf")),
]

VALUE_BUCKET_LABELS = [
    bucket[0]
    for bucket in VALUE_BUCKETS
]


def _coerce_money_numeric(series: pd.Series) -> pd.Series:
    """
    Convert values such as '$1,250,000' or '1,250,000'
    into numeric values.
    """
    s = series.astype(str)
    s = s.str.replace(r"[$,]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def owners_has_trust(owners) -> bool:
    if owners is None or pd.isna(owners):
        return False

    return "TRUST" in str(owners).upper()


def min_nonzero(series: pd.Series):
    series = series.dropna()
    nonzero = series[series != 0]

    if len(nonzero):
        return nonzero.min()

    return pd.NA


def create_value_buckets(data: pd.DataFrame) -> pd.DataFrame:
    """
    Add the AppraisedBucket column using the defined value ranges.
    """
    bucket_bins = [bucket[1] for bucket in VALUE_BUCKETS]
    bucket_bins.append(VALUE_BUCKETS[-1][2])

    data["AppraisedBucket"] = pd.cut(
        data["AppraisedValueNumeric"],
        bins=bucket_bins,
        labels=VALUE_BUCKET_LABELS,
        right=False,
        include_lowest=True,
    )

    return data


def add_bucket_statistics_to_summary(
    data: pd.DataFrame,
    summary: pd.DataFrame,
) -> pd.DataFrame:
    """
    Add yearly transaction counts and appraised-value totals
    for each value bucket.
    """

    # ---------------- Detailed yearly value statistics ----------------
    yearly_value_stats = (
        data.groupby("SaleYear", dropna=False)["AppraisedValueNumeric"]
        .agg(
            MedianAppraisedValue="median",
            AppraisedValueP25=lambda s: s.quantile(0.25),
            AppraisedValueP75=lambda s: s.quantile(0.75),
            AppraisedValueP90=lambda s: s.quantile(0.90),
            AppraisedValueP95=lambda s: s.quantile(0.95),
            AppraisedValueP99=lambda s: s.quantile(0.99),
        )
        .reset_index()
    )

    summary = summary.merge(
        yearly_value_stats,
        on="SaleYear",
        how="left",
    )

    # ---------------- Transaction counts by bucket ----------------
    bucket_counts = (
        data.dropna(subset=["AppraisedBucket"])
        .groupby(
            ["SaleYear", "AppraisedBucket"],
            observed=False,
        )
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    # ---------------- Appraised totals by bucket ----------------
    bucket_values = (
        data.dropna(subset=["AppraisedBucket"])
        .groupby(
            ["SaleYear", "AppraisedBucket"],
            observed=False,
        )["AppraisedValueNumeric"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    # Ensure every bucket column exists, even if no transactions
    # fall into that bucket during the reporting period.
    for bucket_name in VALUE_BUCKET_LABELS:
        if bucket_name not in bucket_counts.columns:
            bucket_counts[bucket_name] = 0

        if bucket_name not in bucket_values.columns:
            bucket_values[bucket_name] = 0

    # Put bucket columns in the desired order
    bucket_counts = bucket_counts[
        ["SaleYear"] + VALUE_BUCKET_LABELS
    ]

    bucket_values = bucket_values[
        ["SaleYear"] + VALUE_BUCKET_LABELS
    ]

    # Rename count columns
    bucket_counts = bucket_counts.rename(
        columns={
            bucket_name: f"TxCount_{bucket_name}"
            for bucket_name in VALUE_BUCKET_LABELS
        }
    )

    # Rename total-value columns
    bucket_values = bucket_values.rename(
        columns={
            bucket_name: f"TxValue_{bucket_name}"
            for bucket_name in VALUE_BUCKET_LABELS
        }
    )

    # Merge counts and values into Summary
    summary = summary.merge(
        bucket_counts,
        on="SaleYear",
        how="left",
    )

    summary = summary.merge(
        bucket_values,
        on="SaleYear",
        how="left",
    )

    # Add bucket totals in millions
    for bucket_name in VALUE_BUCKET_LABELS:
        summary[f"TxValueMM_{bucket_name}"] = (
            summary[f"TxValue_{bucket_name}"] / 1_000_000
        )

    # ---------------- Threshold percentages ----------------
    #
    # These use all transactions for the year as the denominator.
    # Transactions with missing appraised values are not above or
    # below any threshold.

    yearly_total_transactions = (
        data.groupby("SaleYear", dropna=False)
        .size()
        .rename("YearlyTransactionCount")
        .reset_index()
    )

    yearly_under_150k = (
        data.assign(
            Under150k=data["AppraisedValueNumeric"] < 150_000
        )
        .groupby("SaleYear", dropna=False)["Under150k"]
        .sum()
        .rename("Under150kCount")
        .reset_index()
    )

    yearly_over_1m = (
        data.assign(
            Over1M=data["AppraisedValueNumeric"] >= 1_000_000
        )
        .groupby("SaleYear", dropna=False)["Over1M"]
        .sum()
        .rename("Over1MCount")
        .reset_index()
    )

    yearly_over_10m = (
        data.assign(
            Over10M=data["AppraisedValueNumeric"] >= 10_000_000
        )
        .groupby("SaleYear", dropna=False)["Over10M"]
        .sum()
        .rename("Over10MCount")
        .reset_index()
    )

    threshold_stats = yearly_total_transactions.merge(
        yearly_under_150k,
        on="SaleYear",
        how="left",
    )

    threshold_stats = threshold_stats.merge(
        yearly_over_1m,
        on="SaleYear",
        how="left",
    )

    threshold_stats = threshold_stats.merge(
        yearly_over_10m,
        on="SaleYear",
        how="left",
    )

    threshold_stats["PctTransactions_Under_150k"] = (
        threshold_stats["Under150kCount"]
        / threshold_stats["YearlyTransactionCount"]
        * 100
    )

    threshold_stats["PctTransactions_Over_1M"] = (
        threshold_stats["Over1MCount"]
        / threshold_stats["YearlyTransactionCount"]
        * 100
    )

    threshold_stats["PctTransactions_Over_10M"] = (
        threshold_stats["Over10MCount"]
        / threshold_stats["YearlyTransactionCount"]
        * 100
    )

    threshold_stats = threshold_stats[
        [
            "SaleYear",
            "PctTransactions_Under_150k",
            "PctTransactions_Over_1M",
            "PctTransactions_Over_10M",
        ]
    ]

    summary = summary.merge(
        threshold_stats,
        on="SaleYear",
        how="left",
    )

    # Replace missing bucket results with zero
    count_columns = [
        f"TxCount_{bucket_name}"
        for bucket_name in VALUE_BUCKET_LABELS
    ]

    value_columns = [
        f"TxValue_{bucket_name}"
        for bucket_name in VALUE_BUCKET_LABELS
    ]

    value_mm_columns = [
        f"TxValueMM_{bucket_name}"
        for bucket_name in VALUE_BUCKET_LABELS
    ]

    summary[
        count_columns + value_columns + value_mm_columns
    ] = summary[
        count_columns + value_columns + value_mm_columns
    ].fillna(0)

    return summary

def add_summary_charts(writer, summary):
    """
    Add charts to the Summary worksheet.

    Chart 1:
        Number of transactions by price range and year

    Chart 2:
        Total appraised value by price range and year
    """

    worksheet = writer.sheets[SUMMARY_SHEET]

    # Map column names to Excel column numbers
    header_columns = {
        cell.value: cell.column
        for cell in worksheet[1]
    }

    sale_year_column = header_columns["SaleYear"]

    count_columns = [
        f"TxCount_{bucket_name}"
        for bucket_name in VALUE_BUCKET_LABELS
    ]

    value_columns = [
        f"TxValueMM_{bucket_name}"
        for bucket_name in VALUE_BUCKET_LABELS
    ]

    # Excel rows containing yearly data
    first_data_row = 1
    last_data_row = len(summary) + 1

    # ---------------------------------------------------------
    # Chart 1: Number of transactions by price range
    # ---------------------------------------------------------
    transaction_chart = LineChart()

    transaction_chart.title = (
        "Number of Sales by Appraised-Value Range"
    )
    transaction_chart.style = 13
    transaction_chart.y_axis.title = "Number of Sales"
    transaction_chart.x_axis.title = "Sale Year"

    transaction_chart.height = 12
    transaction_chart.width = 24

    year_reference = Reference(
        worksheet,
        min_col=sale_year_column,
        min_row=first_data_row,
        max_row=last_data_row,
    )

    for column_name in count_columns:
        column_number = header_columns[column_name]

        data_reference = Reference(
            worksheet,
            min_col=column_number,
            min_row=first_data_row,
            max_row=last_data_row,
        )

        transaction_chart.add_data(
            data_reference,
            titles_from_data=True,
        )

    transaction_chart.set_categories(year_reference)
    transaction_chart.legend.position = "b"

    # Add the chart to the right of the summary table
    transaction_chart.anchor = "A35"

    worksheet.add_chart(transaction_chart)

    # ---------------------------------------------------------
    # Chart 2: Total appraised value by price range
    # ---------------------------------------------------------
    value_chart = LineChart()

    value_chart.title = (
        "Total Appraised Value by Price Range"
    )
    value_chart.style = 12
    value_chart.y_axis.title = "Total Appraised Value ($ Millions)"
    value_chart.x_axis.title = "Sale Year"

    value_chart.height = 12
    value_chart.width = 24

    for column_name in value_columns:
        column_number = header_columns[column_name]

        data_reference = Reference(
            worksheet,
            min_col=column_number,
            min_row=first_data_row,
            max_row=last_data_row,
        )

        value_chart.add_data(
            data_reference,
            titles_from_data=True,
        )

    value_chart.set_categories(year_reference)
    value_chart.legend.position = "b"

    value_chart.anchor = "A60"

    worksheet.add_chart(value_chart)


def main():
    if not os.path.exists(APPEND_CSV):
        raise RuntimeError(
            f"Missing {APPEND_CSV}. Run the downloader first."
        )

    data = pd.read_csv(APPEND_CSV).drop_duplicates()

    required = {
        SALEDATE_COL,
        APPRAISED_COL,
        "Parcel",
    }

    missing = required - set(data.columns)

    if missing:
        raise RuntimeError(
            f"{APPEND_CSV} missing required columns: {sorted(missing)}"
        )

    # ---------------- Parse dates and years ----------------
    data[SALEDATE_COL] = pd.to_datetime(
        data[SALEDATE_COL],
        errors="coerce",
    )

    data["SaleYear"] = data[SALEDATE_COL].dt.year

    # ---------------- Trust flag ----------------
    if OWNERS_COL in data.columns:
        data["IsTrust"] = data[OWNERS_COL].apply(
            owners_has_trust
        )
    else:
        data["IsTrust"] = False

    # ---------------- Numeric appraised value ----------------
    data["AppraisedValueNumeric"] = _coerce_money_numeric(
        data[APPRAISED_COL]
    )

    # ---------------- Appraised-value buckets ----------------
    data = create_value_buckets(data)

    # ---------------- Yearly Summary ----------------
    summary = (
        data.groupby("SaleYear", dropna=False)
        .agg(
            DeedsTransactions=("SaleYear", "size"),
            TotalAppraisedValue=(
                "AppraisedValueNumeric",
                "sum",
            ),
            AvgAppraisedValue_per_deed=(
                "AppraisedValueNumeric",
                "mean",
            ),
            MinAppraisedValue_nonzero=(
                "AppraisedValueNumeric",
                min_nonzero,
            ),
            MaxAppraisedValue=(
                "AppraisedValueNumeric",
                "max",
            ),
            TrustDeedsCount=(
                "IsTrust",
                "sum",
            ),
            TrustTotalAppraisedValue=(
                "AppraisedValueNumeric",
                lambda s: s[
                    data.loc[s.index, "IsTrust"]
                ].sum(),
            ),
        )
        .reset_index()
    )

    summary["MinAppraisedValueUsed"] = (
        summary["MinAppraisedValue_nonzero"]
    )

    summary["TotalAppraisedValueMM"] = (
        summary["TotalAppraisedValue"] / 1_000_000
    )

    summary["TrustPctOfDeeds"] = (
        summary["TrustDeedsCount"]
        / summary["DeedsTransactions"]
        * 100
    )

    summary["TrustAvgAppraisedValue_per_deed"] = (
        summary["TrustTotalAppraisedValue"]
        / summary["TrustDeedsCount"]
    )

    summary = summary.sort_values("SaleYear")

    # Add detailed value statistics and bucket statistics
    summary = add_bucket_statistics_to_summary(
        data,
        summary,
    )

    summary = summary.sort_values("SaleYear")

    # ---------------- CappedDates sheet ----------------
    if os.path.exists(COMPLETED_CSV):
        ledger = pd.read_csv(COMPLETED_CSV)

        capped = (
            ledger[ledger["Capped"] == True][
                ["SaleDate", "Rows"]
            ]
            .drop_duplicates()
            .sort_values("SaleDate")
        )
    else:
        capped = pd.DataFrame(
            columns=["SaleDate", "Rows"]
        )

    # ---------------- Parcel_Summ sheet ----------------
    parcel_stats = (
        data.groupby("Parcel", dropna=False)
        .agg(
            UniqueSaleYears=(
                "SaleYear",
                lambda s: s.dropna().nunique(),
            ),
            DeedsCount=(
                "Parcel",
                "size",
            ),
            AvgAppraisedValue=(
                "AppraisedValueNumeric",
                "mean",
            ),
            MinAppraisedValue=(
                "AppraisedValueNumeric",
                "min",
            ),
            MaxAppraisedValue=(
                "AppraisedValueNumeric",
                "max",
            ),
        )
        .reset_index()
    )

    def parcel_rate_of_change(group):
        group = group.dropna(
            subset=[
                "SaleYear",
                "AppraisedValueNumeric",
            ]
        )

        if len(group) < 2:
            return pd.NA

        if group["SaleYear"].nunique() < 2:
            return pd.NA

        group = group.sort_values("SaleYear")

        first_row = group.iloc[0]
        last_row = group.iloc[-1]

        value_change = (
            last_row["AppraisedValueNumeric"]
            - first_row["AppraisedValueNumeric"]
        )

        year_change = (
            last_row["SaleYear"]
            - first_row["SaleYear"]
        )

        if year_change == 0 or pd.isna(year_change):
            return pd.NA

        return value_change / year_change

    parcel_rate = (
        data.groupby("Parcel", dropna=False)
        .apply(parcel_rate_of_change)
        .reset_index(
            name="AppraisedValueIncreasePerYear"
        )
    )

    parcel_summ = parcel_stats.merge(
        parcel_rate,
        on="Parcel",
        how="left",
    )

    parcel_summ["HasMultipleDeeds"] = (
        parcel_summ["DeedsCount"] > 1
    )

    # ---------------- Data sheet ----------------
    missing_merge_cols = [
        column
        for column in EXPECTED_MERGED_COLS
        if column not in data.columns
    ]

    if missing_merge_cols:
        raise RuntimeError(
            f"{APPEND_CSV} missing expected merged columns: "
            f"{missing_merge_cols}"
        )

    merged_raw = data[EXPECTED_MERGED_COLS].copy()

    merged_raw[SALEDATE_COL] = data[SALEDATE_COL]
    merged_raw["SaleYear"] = data["SaleYear"]
    merged_raw["IsTrust"] = data["IsTrust"]

    # ---------------- Write workbook ----------------
    with pd.ExcelWriter(
        OUT_XLSX,
        engine="openpyxl",
    ) as writer:
        merged_raw.to_excel(
            writer,
            sheet_name=DATA_SHEET,
            index=False,
        )

        capped.to_excel(
            writer,
            sheet_name=CAPPED_SHEET,
            index=False,
        )

        summary.to_excel(
            writer,
            sheet_name=SUMMARY_SHEET,
            index=False,
        )

        parcel_summ.to_excel(
            writer,
            sheet_name=PARCEL_SUMM_SHEET,
            index=False,
        )

    # Add charts after the Summary sheet has been written
    add_summary_charts(writer, summary)

    print(f"Done. Wrote: {OUT_XLSX}")
    print(f"Merged raw rows: {len(merged_raw)}")
    print(f"Parcel rows: {len(parcel_summ)}")
    print(f"Summary rows: {len(summary)}")


if __name__ == "__main__":
    main()
