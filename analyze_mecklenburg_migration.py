from pathlib import Path
import re
import zipfile
import io

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import geopandas as gpd
import requests


# ============================================================
# SETTINGS
# ============================================================

INPUT_FILE = Path(
    r"C:\Users\jjsha\Downloads\all_states_county_inflow_outflow.xls"
)

OUTPUT_FOLDER = Path(
    r"C:\Users\jjsha\Downloads\mecklenburg_migration_analysis"
)

OUTPUT_FOLDER.mkdir(exist_ok=True)

RAW_SHEET = "raw data"

# Mecklenburg County, North Carolina
MECKLENBURG_STATE = 37
MECKLENBURG_COUNTY = 119

# Download location for Census county boundary shapefile
COUNTY_SHAPEFILE_URL = (
    "https://www2.census.gov/geo/tiger/TIGER2023/COUNTY/"
    "tl_2023_us_county.zip"
)


# ============================================================
# COLUMN UTILITIES
# ============================================================

def normalize_column_name(column):
    """Convert column names to easy-to-match lowercase names."""

    column = str(column).strip().lower()
    column = re.sub(r"[^a-z0-9]+", "_", column)
    return column.strip("_")


def find_column(df, possible_names, required=True):
    """
    Find a dataframe column using several possible names.
    """

    normalized = {
        normalize_column_name(col): col
        for col in df.columns
    }

    for name in possible_names:
        name = normalize_column_name(name)

        if name in normalized:
            return normalized[name]

    if required:
        raise KeyError(
            f"Could not find one of these columns: {possible_names}\n"
            f"Available columns are:\n{list(df.columns)}"
        )

    return None


def prepare_data(df):
    """
    Standardize important fields while retaining all original columns.
    """

    year_col = find_column(
        df,
        ["year", "tax_year", "ending_year"]
    )

    from_state_col = find_column(
        df,
        ["from_state", "origin_state", "origin_state_fips"]
    )

    from_county_col = find_column(
        df,
        ["from_county", "origin_county", "origin_county_fips"]
    )

    to_state_col = find_column(
        df,
        ["to_state", "destination_state", "destination_state_fips"]
    )

    to_county_col = find_column(
        df,
        ["to_county", "destination_county", "destination_county_fips"]
    )

    returns_col = find_column(
        df,
        ["returns", "tax_returns", "number_of_returns"]
    )

    exemptions_col = find_column(
        df,
        ["exemptions", "number_of_exemptions"],
        required=False
    )

    agi_col = find_column(
        df,
        ["agi", "adjusted_gross_income"],
        required=False
    )

    data = pd.DataFrame()

    data["year"] = pd.to_numeric(
        df[year_col],
        errors="coerce"
    )

    data["from_state"] = pd.to_numeric(
        df[from_state_col],
        errors="coerce"
    )

    data["from_county"] = pd.to_numeric(
        df[from_county_col],
        errors="coerce"
    )

    data["to_state"] = pd.to_numeric(
        df[to_state_col],
        errors="coerce"
    )

    data["to_county"] = pd.to_numeric(
        df[to_county_col],
        errors="coerce"
    )

    data["returns"] = pd.to_numeric(
        df[returns_col],
        errors="coerce"
    ).fillna(0)

    if exemptions_col:
        data["exemptions"] = pd.to_numeric(
            df[exemptions_col],
            errors="coerce"
        ).fillna(0)
    else:
        data["exemptions"] = np.nan

    if agi_col:
        data["agi"] = pd.to_numeric(
            df[agi_col],
            errors="coerce"
        ).fillna(0)
    else:
        data["agi"] = np.nan

    # Remove invalid geography records.
    data = data.dropna(
        subset=[
            "year",
            "from_state",
            "from_county",
            "to_state",
            "to_county"
        ]
    )

    for column in [
        "year",
        "from_state",
        "from_county",
        "to_state",
        "to_county"
    ]:
        data[column] = data[column].astype(int)

    # Remove same-county records and obvious aggregate geography.
    data = data[
        ~(
            (data["from_state"] == data["to_state"]) &
            (data["from_county"] == data["to_county"])
        )
    ]

    return data


# ============================================================
# LOAD DATA
# ============================================================

raw = pd.read_excel(
    INPUT_FILE,
    sheet_name=RAW_SHEET
)

print("Original columns:")
print(list(raw.columns))

data = prepare_data(raw)

print(f"\nRows loaded: {len(data):,}")
print(f"Years: {data['year'].min()}–{data['year'].max()}")


# ============================================================
# IDENTIFY MECKLENBURG INFLOWS AND OUTFLOWS
# ============================================================

incoming = data[
    (data["to_state"] == MECKLENBURG_STATE) &
    (data["to_county"] == MECKLENBURG_COUNTY)
].copy()

outgoing = data[
    (data["from_state"] == MECKLENBURG_STATE) &
    (data["from_county"] == MECKLENBURG_COUNTY)
].copy()

print(f"Incoming flow rows: {len(incoming):,}")
print(f"Outgoing flow rows: {len(outgoing):,}")


# ============================================================
# 1. NET MIGRATION TO MECKLENBURG COUNTY
# ============================================================

annual_in = (
    incoming
    .groupby("year", as_index=False)
    [["returns", "exemptions", "agi"]]
    .sum()
    .rename(columns={
        "returns": "in_returns",
        "exemptions": "in_exemptions",
        "agi": "in_agi"
    })
)

annual_out = (
    outgoing
    .groupby("year", as_index=False)
    [["returns", "exemptions", "agi"]]
    .sum()
    .rename(columns={
        "returns": "out_returns",
        "exemptions": "out_exemptions",
        "agi": "out_agi"
    })
)

annual = annual_in.merge(
    annual_out,
    on="year",
    how="outer"
).fillna(0)

annual["net_returns"] = (
    annual["in_returns"] -
    annual["out_returns"]
)

annual["net_exemptions"] = (
    annual["in_exemptions"] -
    annual["out_exemptions"]
)

annual["net_agi"] = (
    annual["in_agi"] -
    annual["out_agi"]
)

annual["total_return_flow"] = (
    annual["in_returns"] +
    annual["out_returns"]
)

annual["net_returns_percent_of_flow"] = np.where(
    annual["total_return_flow"] != 0,
    annual["net_returns"] /
    annual["total_return_flow"] * 100,
    np.nan
)

annual.to_csv(
    OUTPUT_FOLDER / "mecklenburg_annual_migration_summary.csv",
    index=False
)

print("\nAnnual Mecklenburg migration summary:")
print(annual.to_string(index=False))


# ============================================================
# 2. WEIGHTED AVERAGE AGI
# ============================================================

def weighted_average_agi(group):
    """
    Weighted average AGI per tax return.

    IRS AGI may be reported in thousands of dollars depending
    on the source file. Keep the units consistent with the raw file.
    """

    total_returns = group["returns"].sum()
    total_agi = group["agi"].sum()

    if total_returns == 0:
        return np.nan

    return total_agi / total_returns


in_agi = (
    incoming
    .groupby("year")
    .apply(weighted_average_agi, include_groups=False)
    .rename("weighted_avg_agi_in")
)

out_agi = (
    outgoing
    .groupby("year")
    .apply(weighted_average_agi, include_groups=False)
    .rename("weighted_avg_agi_out")
)

agi_trend = pd.concat(
    [in_agi, out_agi],
    axis=1
).reset_index()


# ------------------------------------------------------------
# Optional base-population AGI
# ------------------------------------------------------------
#
# The raw county-flow sheet does not necessarily contain the
# average AGI of all Mecklenburg residents. It contains AGI
# associated with movers between counties.
#
# If you have a county-total file, place it here and update
# the column names below.
#
# Expected columns:
#   year
#   state
#   county
#   returns
#   agi
#
# Example:
# BASE_FILE = Path(r"C:\Users\jjsha\Downloads\county_totals.xlsx")
#
# If BASE_FILE remains None, the script will not fabricate a
# base-population estimate.

BASE_FILE = None


if BASE_FILE is not None and Path(BASE_FILE).exists():

    base_raw = pd.read_excel(BASE_FILE)

    base = pd.DataFrame({
        "year": pd.to_numeric(base_raw["year"], errors="coerce"),
        "state": pd.to_numeric(base_raw["state"], errors="coerce"),
        "county": pd.to_numeric(base_raw["county"], errors="coerce"),
        "returns": pd.to_numeric(
            base_raw["returns"],
            errors="coerce"
        ).fillna(0),
        "agi": pd.to_numeric(
            base_raw["agi"],
            errors="coerce"
        ).fillna(0),
    })

    base_mecklenburg = base[
        (base["state"] == MECKLENBURG_STATE) &
        (base["county"] == MECKLENBURG_COUNTY)
    ].copy()

    base_mecklenburg["weighted_avg_agi_base"] = np.where(
        base_mecklenburg["returns"] != 0,
        base_mecklenburg["agi"] /
        base_mecklenburg["returns"],
        np.nan
    )

    agi_trend = agi_trend.merge(
        base_mecklenburg[
            ["year", "weighted_avg_agi_base"]
        ],
        on="year",
        how="left"
    )

else:
    agi_trend["weighted_avg_agi_base"] = np.nan
    print(
        "\nNo base-population file supplied. "
        "Base AGI cannot be calculated from county-flow data alone."
    )


agi_trend.to_csv(
    OUTPUT_FOLDER / "mecklenburg_weighted_average_agi_trend.csv",
    index=False
)


# ------------------------------------------------------------
# Plot weighted average AGI trend
# ------------------------------------------------------------

plt.figure(figsize=(11, 6))

plt.plot(
    agi_trend["year"],
    agi_trend["weighted_avg_agi_in"],
    marker="o",
    linewidth=2,
    label="In-migration"
)

plt.plot(
    agi_trend["year"],
    agi_trend["weighted_avg_agi_out"],
    marker="o",
    linewidth=2,
    label="Out-migration"
)

if agi_trend["weighted_avg_agi_base"].notna().any():
    plt.plot(
        agi_trend["year"],
        agi_trend["weighted_avg_agi_base"],
        marker="o",
        linewidth=2,
        label="Mecklenburg base population"
    )

plt.title("Weighted Average AGI of Mecklenburg Migration")
plt.xlabel("IRS tax year")
plt.ylabel("AGI per tax return")
plt.grid(alpha=0.3)
plt.legend()
plt.tight_layout()
plt.savefig(
    OUTPUT_FOLDER / "weighted_average_agi_trend.png",
    dpi=300
)
plt.close()


# ============================================================
# 3. COUNTY-LEVEL FLOW DATA
# ============================================================

# Annual total Mecklenburg inflow and outflow.
total_in_by_year = (
    incoming
    .groupby("year")["returns"]
    .sum()
    .rename("total_mecklenburg_in_returns")
)

total_out_by_year = (
    outgoing
    .groupby("year")["returns"]
    .sum()
    .rename("total_mecklenburg_out_returns")
)


# Incoming counties.
inflow_counties = (
    incoming
    .groupby(
        ["year", "from_state", "from_county"],
        as_index=False
    )
    [["returns", "exemptions", "agi"]]
    .rename(columns={
        "from_state": "county_state",
        "from_county": "county",
        "returns": "in_returns",
        "exemptions": "in_exemptions",
        "agi": "in_agi"
    })
)

inflow_counties["flow_direction"] = "To Mecklenburg"

inflow_counties = inflow_counties.merge(
    total_in_by_year,
    on="year",
    how="left"
)

inflow_counties["flow_percent"] = np.where(
    inflow_counties["total_mecklenburg_in_returns"] != 0,
    inflow_counties["in_returns"] /
    inflow_counties["total_mecklenburg_in_returns"] * 100,
    0
)


# Outgoing counties.
outflow_counties = (
    outgoing
    .groupby(
        ["year", "to_state", "to_county"],
        as_index=False
    )
    [["returns", "exemptions", "agi"]]
    .rename(columns={
        "to_state": "county_state",
        "to_county": "county",
        "returns": "out_returns",
        "exemptions": "out_exemptions",
        "agi": "out_agi"
    })
)

outflow_counties["flow_direction"] = "From Mecklenburg"

outflow_counties = outflow_counties.merge(
    total_out_by_year,
    on="year",
    how="left"
)

outflow_counties["flow_percent"] = np.where(
    outflow_counties["total_mecklenburg_out_returns"] != 0,
    outflow_counties["out_returns"] /
    outflow_counties["total_mecklenburg_out_returns"] * 100,
    0
)


# Save flow tables.
inflow_counties.to_csv(
    OUTPUT_FOLDER / "county_inflows_to_mecklenburg.csv",
    index=False
)

outflow_counties.to_csv(
    OUTPUT_FOLDER / "county_outflows_from_mecklenburg.csv",
    index=False
)


# ============================================================
# DOWNLOAD COUNTY GEOMETRIES
# ============================================================

def get_county_shapes():
    """
    Download and read US county boundaries.
    """

    response = requests.get(
        COUNTY_SHAPEFILE_URL,
        timeout=120
    )

    response.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(response.content)) as zip_file:
        shapefile_name = [
            name for name in zip_file.namelist()
            if name.endswith(".shp")
        ][0]

        with zip_file.open(shapefile_name) as shp:
            pass

        # GeoPandas can read the zipped shapefile directly from bytes
        local_zip = OUTPUT_FOLDER / "county_boundaries.zip"

        with open(local_zip, "wb") as file:
            file.write(response.content)

    return gpd.read_file(f"zip://{local_zip}")


counties = get_county_shapes()

# TIGER/Line GEOID is 5-digit state FIPS + county FIPS.
counties["county_state"] = pd.to_numeric(
    counties["GEOID"].str[:2],
    errors="coerce"
)

counties["county"] = pd.to_numeric(
    counties["GEOID"].str[2:],
    errors="coerce"
)

# Use a projected coordinate system for plotting.
counties = counties.to_crs("EPSG:5070")


# ============================================================
# PREPARE AGI COLOR VARIABLE
# ============================================================

# Average AGI by county from all observed county-flow records.
#
# This represents AGI per return associated with that county's
# migration flows, not necessarily AGI of every county resident.

county_agi = (
    data
    .groupby(
        ["from_state", "from_county"],
        as_index=False
    )
    [["returns", "agi"]]
    .sum()
)

county_agi["avg_agi"] = np.where(
    county_agi["returns"] != 0,
    county_agi["agi"] / county_agi["returns"],
    np.nan
)

county_agi = county_agi.rename(columns={
    "from_state": "county_state",
    "from_county": "county"
})

# Mecklenburg reference average.
mecklenburg_avg_agi = county_agi.loc[
    (county_agi["county_state"] == MECKLENBURG_STATE) &
    (county_agi["county"] == MECKLENBURG_COUNTY),
    "avg_agi"
].iloc[0]

county_agi["agi_relative_to_mecklenburg"] = (
    county_agi["avg_agi"] /
    mecklenburg_avg_agi
)


# ============================================================
# CREATE FLOW MAPS
# ============================================================

def make_flow_map(year, direction, number_of_arrows=30):
    """
    Create a county-level flow map.

    Arrow width is proportional to the percentage of total
    Mecklenburg inflow or outflow represented by that county.

    County fill color is based on average AGI relative to
    Mecklenburg County.
    """

    if direction == "in":

        flow = inflow_counties[
            inflow_counties["year"] == year
        ].copy()

        flow = flow.sort_values(
            "flow_percent",
            ascending=False
        ).head(number_of_arrows)

        source_col = "county_state"
        target_col = "county"

        output_name = (
            f"flow_map_to_mecklenburg_{year}.png"
        )

        title = (
            f"Top County Flows to Mecklenburg County, {year}"
        )

    elif direction == "out":

        flow = outflow_counties[
            outflow_counties["year"] == year
        ].copy()

        flow = flow.sort_values(
            "flow_percent",
            ascending=False
        ).head(number_of_arrows)

        source_col = "county_state"
        target_col = "county"

        output_name = (
            f"flow_map_from_mecklenburg_{year}.png"
        )

        title = (
            f"Top County Flows from Mecklenburg County, {year}"
        )

    else:
        raise ValueError("direction must be 'in' or 'out'")

    # Join AGI information to county shapes.
    map_counties = counties.merge(
        county_agi,
        on=["county_state", "county"],
        how="left"
    )

    fig, ax = plt.subplots(figsize=(13, 11))

    # Diverging color scale:
    # blue = below Mecklenburg
    # white = similar to Mecklenburg
    # red = above Mecklenburg
    map_counties.plot(
        ax=ax,
        column="agi_relative_to_mecklenburg",
        cmap="RdBu_r",
        vmin=0.25,
        vmax=2.0,
        legend=True,
        legend_kwds={
            "label": "County average AGI / Mecklenburg average AGI"
        },
        linewidth=0.15,
        edgecolor="gray",
        missing_kwds={
            "color": "lightgray",
            "label": "No AGI data"
        }
    )

    # Plot Mecklenburg County prominently.
    mecklenburg = map_counties[
        (map_counties["county_state"] == MECKLENBURG_STATE) &
        (map_counties["county"] == MECKLENBURG_COUNTY)
    ]

    mecklenburg.plot(
        ax=ax,
        facecolor="none",
        edgecolor="black",
        linewidth=2.5
    )

    # County centroids.
    centroids = map_counties.copy()
    centroids["centroid"] = centroids.geometry.centroid

    center = centroids[
        (centroids["county_state"] == MECKLENBURG_STATE) &
        (centroids["county"] == MECKLENBURG_COUNTY)
    ]["centroid"].iloc[0]

    # Draw arrows.
    for _, row in flow.iterrows():

        county_point = centroids[
            (centroids["county_state"] == row[source_col]) &
            (centroids["county"] == row[target_col])
        ]

        if county_point.empty:
            continue

        origin = county_point["centroid"].iloc[0]

        if direction == "in":
            start = origin
            end = center
            color = "darkgreen"
        else:
            start = center
            end = origin
            color = "purple"

        # Scale line width using percentage of total flow.
        # Minimum width keeps small flows visible.
        width = max(0.5, row["flow_percent"] * 0.20)

        ax.annotate(
            "",
            xy=(end.x, end.y),
            xytext=(start.x, start.y),
            arrowprops=dict(
                arrowstyle="->",
                color=color,
                linewidth=width,
                alpha=0.65,
                shrinkA=3,
                shrinkB=8,
                connectionstyle="arc3,rad=0.08"
            )
        )

    ax.set_title(
        title +
        "\nArrow width = percentage of Mecklenburg county flow; "
        "color = relative average AGI",
        fontsize=13
    )

    ax.set_axis_off()

    plt.tight_layout()
    plt.savefig(
        OUTPUT_FOLDER / output_name,
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    print(f"Created {output_name}")


# Create maps for every year in the data.
for year in sorted(data["year"].unique()):

    make_flow_map(
        year=year,
        direction="in",
        number_of_arrows=30
    )

    make_flow_map(
        year=year,
        direction="out",
        number_of_arrows=30
    )


# ============================================================
# PLOT NET MIGRATION TREND
# ============================================================

fig, axes = plt.subplots(
    3,
    1,
    figsize=(11, 13),
    sharex=True
)

axes[0].bar(
    annual["year"],
    annual["net_returns"],
    color=np.where(
        annual["net_returns"] >= 0,
        "seagreen",
        "firebrick"
    )
)
axes[0].axhline(0, color="black", linewidth=0.8)
axes[0].set_ylabel("Net returns")
axes[0].set_title("Net Migration to Mecklenburg County")

axes[1].bar(
    annual["year"],
    annual["net_exemptions"],
    color=np.where(
        annual["net_exemptions"] >= 0,
        "seagreen",
        "firebrick"
    )
)
axes[1].axhline(0, color="black", linewidth=0.8)
axes[1].set_ylabel("Net exemptions")

axes[2].bar(
    annual["year"],
    annual["net_agi"],
    color=np.where(
        annual["net_agi"] >= 0,
        "seagreen",
        "firebrick"
    )
)
axes[2].axhline(0, color="black", linewidth=0.8)
axes[2].set_ylabel("Net AGI")
axes[2].set_xlabel("IRS tax year")

plt.tight_layout()
plt.savefig(
    OUTPUT_FOLDER / "mecklenburg_net_migration_trend.png",
    dpi=300
)
plt.close()


print("\nAnalysis complete.")
print(f"Results saved to:\n{OUTPUT_FOLDER}")
