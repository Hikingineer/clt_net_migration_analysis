import os
import re
import json
import time
import random

import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright


# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

INPUT_CSV = "deeds_data_append.csv"
OUTPUT_CSV = "data_enriched.csv"

PAGE_TIMEOUT_SECONDS = 60
PAGE_RENDER_WAIT_SECONDS = 8
SLEEP_BETWEEN_URLS_SECONDS = 1.5
MAX_RETRIES = 3

HEADLESS = True


ENRICH_COLS = [
    "Zipcode",
    "Land Use Code",
    "Neighborhood",
    "Land Use Desc",
    "Land",
    "Exemption / Deferment",
    "Municipality",
    "Last Sale Date (KeyInfo)",
    "Fire District",
    "Special District",
    "Legal Description",
    "Last Sale Price",
    "Land Value",
    "Building Value",
    "Features",
    "Assessment Total",
    "Finished Area",
    "Year Built",
    "Built Use / Style",
    "Grade",
    "Story",
    "Heat",
    "Fuel",
    "Foundation",
    "External Wall",
    "Fireplace(s)",
    "Full Bath(s)",
    "Half Bath(s)",
    "Bedroom(s)",
    "Building Total (SqFt)",
    "ValueChanges_JSON",
]


# ------------------------------------------------------------
# General utilities
# ------------------------------------------------------------

def normalize_website(url):
    """Normalize URLs from the Website column."""
    if pd.isna(url):
        return ""

    url = str(url).strip()

    if not url:
        return ""

    if url.startswith("//"):
        return "https:" + url

    if url.startswith("/"):
        return "https://property.spatialest.com" + url

    return url


def default_record(url):
    """Create an empty output record."""
    record = {column: "" for column in ENRICH_COLS}
    record["Website"] = url
    record["ValueChanges_JSON"] = "[]"
    return record


def clean_text(value):
    """Collapse whitespace."""
    if value is None:
        return ""

    return " ".join(str(value).split())


def append_record(record):
    """Append one record to the output CSV."""
    row = pd.DataFrame([record])

    file_exists = os.path.exists(OUTPUT_CSV)

    row.to_csv(
        OUTPUT_CSV,
        mode="a",
        index=False,
        header=not file_exists,
        encoding="utf-8-sig",
    )


# ------------------------------------------------------------
# Browser functions
# ------------------------------------------------------------

def open_and_read_page(page, url):
    """
    Open a property URL and return:
      rendered_html: HTML after JavaScript executes
      visible_text: text visible on the page
    """

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            print(f"    Opening page, attempt {attempt}/{MAX_RETRIES}")

            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_SECONDS * 1000,
            )

            # The property application loads data after the initial page.
            page.wait_for_timeout(PAGE_RENDER_WAIT_SECONDS * 1000)

            visible_text = page.locator("body").inner_text(
                timeout=30000
            ).strip()

            rendered_html = page.content()

            if not visible_text:
                raise RuntimeError("The page contained no visible text")

            return rendered_html, visible_text

        except Exception as error:
            last_error = error

            wait_time = attempt * attempt + random.random()
            print(f"    Page error: {error}")
            print(f"    Retrying in {wait_time:.1f} seconds")
            time.sleep(wait_time)

    raise RuntimeError(
        f"Could not load page after {MAX_RETRIES} attempts: {last_error}"
    )


# ------------------------------------------------------------
# Parsing functions
# ------------------------------------------------------------

def extract_zipcode(soup, visible_text):
    """
    Try to extract the ZIP code specifically from the mailing section.
    Fall back to the visible page text.
    """

    # Based on the XPath supplied earlier:
    # header.record-card-header ... div.mailing ... div.value
    mailing_value = soup.select_one(
        "header.record-card-header div.mailing div.value"
    )

    if mailing_value:
        mailing_text = clean_text(mailing_value.get_text(" ", strip=True))
        match = re.search(r"\b(\d{5})(?:-\d{4})?\b", mailing_text)

        if match:
            return match.group(1)

    # Additional CSS fallback
    mailing_elements = soup.select(
        "header.record-card-header .mailing, "
        ".record-card-header .mailing, "
        ".mailing"
    )

    for element in mailing_elements:
        text = clean_text(element.get_text(" ", strip=True))
        match = re.search(r"\b(\d{5})(?:-\d{4})?\b", text)

        if match:
            return match.group(1)

    # General fallback
    match = re.search(r"\b(\d{5})(?:-\d{4})?\b", visible_text)

    if match:
        return match.group(1)

    return ""


def get_table_label_values(soup):
    """
    Extract label/value pairs from HTML tables.

    Example:
        <td>Neighborhood</td><td>Example</td>
    becomes:
        {"Neighborhood": "Example"}
    """

    label_map = {}

    for table in soup.find_all("table"):
        rows = table.find_all("tr")

        for row in rows:
            cells = row.find_all(["td", "th"])

            if len(cells) < 2:
                continue

            # Handle a normal two-column row.
            if len(cells) == 2:
                label = clean_text(cells[0].get_text(" ", strip=True))
                value = clean_text(cells[1].get_text(" ", strip=True))

                if label and value:
                    label_map[label] = value

            # Handle rows containing multiple label/value pairs.
            else:
                for position in range(0, len(cells) - 1, 2):
                    label = clean_text(
                        cells[position].get_text(" ", strip=True)
                    )
                    value = clean_text(
                        cells[position + 1].get_text(" ", strip=True)
                    )

                    if label and value:
                        label_map[label] = value

    return label_map


def extract_label_value_from_text(visible_text, label):
    """
    Fallback parser for visible text where a label and value occur
    on adjacent lines.
    """

    lines = [
        clean_text(line)
        for line in visible_text.splitlines()
        if clean_text(line)
    ]

    for index, line in enumerate(lines):
        normalized_line = line.rstrip(":").strip().lower()
        normalized_label = label.rstrip(":").strip().lower()

        if normalized_line == normalized_label:
            if index + 1 < len(lines):
                next_line = lines[index + 1]

                # Avoid returning another label as the value.
                if next_line.lower() != normalized_line:
                    return next_line

        # Also handle "Label: Value"
        if normalized_line.startswith(normalized_label + ":"):
            value = line[len(label):].lstrip(" :")
            if value:
                return value

    return ""


def get_value(label_map, visible_text, possible_labels):
    """
    Try several possible spellings for a field.
    First searches tables, then visible text.
    """

    for label in possible_labels:
        if label in label_map:
            return label_map[label]

    for label in possible_labels:
        value = extract_label_value_from_text(visible_text, label)

        if value:
            return value

    return ""


def extract_value_changes(soup):
    """
    Extract value-change rows from a table if available.
    """

    results = []

    for table in soup.find_all("table"):
        table_text = clean_text(table.get_text(" ", strip=True)).lower()

        if "value change" not in table_text:
            continue

        for row in table.find_all("tr"):
            cells = row.find_all(["td", "th"])

            if len(cells) < 4:
                continue

            values = [
                clean_text(cell.get_text(" ", strip=True))
                for cell in cells[:4]
            ]

            # Skip an apparent header row.
            if values[0].lower() in {"date", "date of value change"}:
                continue

            if any(values):
                results.append({
                    "date_of_value_change": values[0],
                    "effective_for_tax_year": values[1],
                    "reason_for_change": values[2],
                    "new_value": values[3],
                })

        if results:
            break

    return json.dumps(results, ensure_ascii=False)


def parse_property(rendered_html, visible_text, url):
    """
    Parse the rendered page into one CSV record.
    """

    record = default_record(url)

    soup = BeautifulSoup(rendered_html, "html.parser")
    label_map = get_table_label_values(soup)

    # ZIP code
    record["Zipcode"] = extract_zipcode(soup, visible_text)

    # Key information
    field_labels = {
        "Land Use Code": ["Land Use Code"],
        "Neighborhood": ["Neighborhood"],
        "Land Use Desc": ["Land Use Desc", "Land Use Description"],
        "Land": ["Land"],
        "Exemption / Deferment": [
            "Exemption / Deferment",
            "Exemption/Deferment",
        ],
        "Municipality": ["Municipality"],
        "Last Sale Date (KeyInfo)": [
            "Last Sale Date",
            "Last Sale Date (KeyInfo)",
        ],
        "Fire District": ["Fire District"],
        "Special District": ["Special District"],
        "Legal Description": ["Legal Description"],
        "Last Sale Price": ["Last Sale Price"],
    }

    # Assessment information
    field_labels.update({
        "Land Value": ["Land Value"],
        "Building Value": ["Building Value"],
        "Features": ["Features"],
        "Assessment Total": [
            "Assessment Total",
            "Total",
        ],
    })

    # Building information
    field_labels.update({
        "Finished Area": ["Finished Area"],
        "Year Built": ["Year Built"],
        "Built Use / Style": [
            "Built Use / Style",
            "Built Use",
            "Style",
        ],
        "Grade": ["Grade"],
        "Story": ["Story", "Stories"],
        "Heat": ["Heat"],
        "Fuel": ["Fuel"],
        "Foundation": ["Foundation"],
        "External Wall": ["External Wall"],
        "Fireplace(s)": ["Fireplace(s)", "Fireplaces"],
        "Full Bath(s)": ["Full Bath(s)", "Full Baths"],
        "Half Bath(s)": ["Half Bath(s)", "Half Baths"],
        "Bedroom(s)": ["Bedroom(s)", "Bedrooms"],
        "Building Total (SqFt)": [
            "Building Total (SqFt)",
            "Total (SqFt)",
            "Building Total",
        ],
    })

    for output_column, possible_labels in field_labels.items():
        record[output_column] = get_value(
            label_map,
            visible_text,
            possible_labels,
        )

    record["ValueChanges_JSON"] = extract_value_changes(soup)

    return record


# ------------------------------------------------------------
# Main program
# ------------------------------------------------------------

def main():
    if not os.path.exists(INPUT_CSV):
        raise FileNotFoundError(
            f"Input file not found: {INPUT_CSV}"
        )

    input_df = pd.read_csv(INPUT_CSV)

    if "Website" not in input_df.columns:
        raise ValueError(
            "The input CSV must contain a column named 'Website'."
        )

    # These are the URLs the script will scrape.
    urls = (
        input_df["Website"]
        .apply(normalize_website)
        .dropna()
        .tolist()
    )

    # Remove blanks and duplicates while preserving order.
    unique_urls = list(dict.fromkeys(
        url for url in urls if url
    ))

    # Resume capability: skip URLs already in the output CSV.
    already_scraped = set()

    if os.path.exists(OUTPUT_CSV):
        output_df = pd.read_csv(OUTPUT_CSV)

        if "Website" in output_df.columns:
            already_scraped = set(
                output_df["Website"]
                .dropna()
                .astype(str)
                .tolist()
            )

    print(f"Input URLs: {len(unique_urls)}")
    print(f"Already scraped: {len(already_scraped)}")
    print(f"Remaining: {len(unique_urls) - len(already_scraped)}")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=HEADLESS
        )

        context = browser.new_context(
            viewport={"width": 1440, "height": 1000},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
            locale="en-US",
        )

        page = context.new_page()

        for index, url in enumerate(unique_urls, start=1):
            if url in already_scraped:
                continue

            print(f"\n[{index}/{len(unique_urls)}] {url}")

            try:
                rendered_html, visible_text = open_and_read_page(
                    page,
                    url,
                )

                record = parse_property(
                    rendered_html,
                    visible_text,
                    url,
                )

                append_record(record)

                print(
                    "    Saved:"
                    f" Zipcode={record['Zipcode']!r},"
                    f" Neighborhood={record['Neighborhood']!r},"
                    f" Year Built={record['Year Built']!r}"
                )

                already_scraped.add(url)

            except Exception as error:
                print(f"    FAILED: {error}")

                record = default_record(url)
                record["scrape_error"] = str(error)

                append_record(record)

                # Mark it as processed so the script does not repeatedly
                # retry the same failed URL on the next run.
                already_scraped.add(url)

            time.sleep(
                SLEEP_BETWEEN_URLS_SECONDS
                + random.random() * 0.5
            )

        browser.close()

    print("\nFinished.")
    print(f"Results saved to: {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
