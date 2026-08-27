import os, time, random, json, re
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

DATA_APPEND_CSV = "data_append.csv"
ENRICHED_CSV = "data_enriched.csv"

DETAIL_TIMEOUT_SECONDS = 60
DETAIL_SLEEP_SECONDS = 0.8
DETAIL_MAX_RETRIES = 5

# Playwright defaults
PLAYWRIGHT_HEADLESS = True
PLAYWRIGHT_NAV_WAIT_UNTIL = "networkidle"

ENRICH_COLS = [
    # Zipcode
    "Zipcode",

    # Key Information
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

    # Assessment details
    "Land Value",
    "Building Value",
    "Features",
    "Assessment Total",

    # Building table
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

    # Value changes
    "ValueChanges_JSON",
]


def default_record(url: str) -> dict:
    rec = {c: "" for c in ENRICH_COLS}
    rec["Website"] = url
    return rec


def text_or_blank(el) -> str:
    if not el:
        return ""
    return " ".join(el.get_text(" ", strip=True).split())


def fetch_html(url: str) -> str:
    """
    Uses Playwright to render the JS-driven SPA so the record data is present in the DOM.
    """
    last_err = None
    for attempt in range(1, DETAIL_MAX_RETRIES + 1):
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=PLAYWRIGHT_HEADLESS)
                page = browser.new_page()

                page.set_default_timeout(30000)
                page.goto(url, wait_until=PLAYWRIGHT_NAV_WAIT_UNTIL)

                # Wait for something that looks like the record-card to exist.
                # If this ever times out, you can loosen/remove it and rely on page.content().
                page.wait_for_selector(
                    "header.record-card-header, div.record-card-header",
                    timeout=15000
                )

                html = page.content()
                browser.close()
                return html

        except Exception as e:
            last_err = e
            sleep_for = (attempt ** 2) * 0.4 + random.random() * 0.4
            print(f"  playwright fetch error attempt {attempt}/{DETAIL_MAX_RETRIES}: {e} -> sleep {sleep_for:.1f}s")
            time.sleep(sleep_for)

    raise RuntimeError(f"Failed fetching {url}. Last error: {last_err}")


def parse_detail(html: str, url: str) -> dict:
    """
    Never raises for missing sections/fields.
    Returns a record containing all ENRICH_COLS, defaulting to "".
    """
    rec = default_record(url)

    try:
        soup = BeautifulSoup(html, "html.parser")
        page_text = soup.get_text("\n", strip=True)

        # Zipcode: keep your best-effort 5-digit match for now
        m = re.search(r"\b(\d{5})\b", page_text)
        if m:
            rec["Zipcode"] = m.group(1)

        # Collect label/value pairs from any tables (best-effort)
        label_map = {}
        for t in soup.find_all("table"):
            try:
                for tr in t.find_all("tr"):
                    tds = tr.find_all(["td", "th"])
                    if len(tds) >= 2:
                        label = text_or_blank(tds[0])
                        value = text_or_blank(tds[1])
                        if label and value:
                            label_map[label] = value
            except Exception:
                continue

        # Some SPAs render key/value blocks without tables.
        # Best-effort: look for common label text patterns and grab nearby text.
        # (This is intentionally conservative; your table parsing is the main path.)
        # If tables work, this won't change anything.

        # Key Information mapping (best-effort)
        key_map = {
            "Land Use Code": "Land Use Code",
            "Neighborhood": "Neighborhood",
            "Land Use Desc": "Land Use Desc",
            "Land": "Land",
            "Exemption / Deferment": "Exemption / Deferment",
            "Municipality": "Municipality",
            "Last Sale Date": "Last Sale Date (KeyInfo)",
            "Fire District": "Fire District",
            "Special District": "Special District",
            "Legal Description": "Legal Description",
            "Last Sale Price": "Last Sale Price",
        }
        for src, dest in key_map.items():
            if src in label_map:
                rec[dest] = label_map[src]

        # Assessment mapping (best-effort)
        assess_map = {
            "Land Value": "Land Value",
            "Building Value": "Building Value",
            "Features": "Features",
            "Total": "Assessment Total",
            "Assessment Total": "Assessment Total",
        }
        for src, dest in assess_map.items():
            if src in label_map:
                rec[dest] = label_map[src]

        # Building table fields (best-effort):
        building_map = {
            "Finished Area": "Finished Area",
            "Year Built": "Year Built",
            "Built Use / Style": "Built Use / Style",
            "Grade": "Grade",
            "Story": "Story",
            "Heat": "Heat",
            "Fuel": "Fuel",
            "Foundation": "Foundation",
            "External Wall": "External Wall",
            "Fireplace(s)": "Fireplace(s)",
            "Full Bath(s)": "Full Bath(s)",
            "Half Bath(s)": "Half Bath(s)",
            "Bedroom(s)": "Bedroom(s)",
            "Total (SqFt)": "Building Total (SqFt)",
            "Building Total (SqFt)": "Building Total (SqFt)",
        }
        for src, dest in building_map.items():
            if src in label_map:
                rec[dest] = label_map[src]

        # Value changes: JSON extraction if it's in tables
        value_changes_rows = []
        for t in soup.find_all("table"):
            try:
                if "Value Changes" in t.get_text(" ", strip=True):
                    for tr in t.find_all("tr"):
                        cols = tr.find_all(["td", "th"])
                        if len(cols) >= 4:
                            vals = [text_or_blank(c) for c in cols[:4]]
                            if any(vals):
                                value_changes_rows.append({
                                    "date_of_value_change": vals[0],
                                    "effective_for_tax_year": vals[1],
                                    "reason_for_change": vals[2],
                                    "new_value": vals[3],
                                })
                    break
            except Exception:
                continue

        rec["ValueChanges_JSON"] = json.dumps(value_changes_rows, ensure_ascii=False)

    except Exception as e:
        rec["ParseError"] = str(e)

    return rec


def normalize_website(url: str) -> str:
    if not isinstance(url, str):
        return ""
    u = url.strip()
    if not u:
        return ""
    if u.startswith("//"):
        return "https:" + u
    if u.startswith("/"):
        return "https://property.spatialest.com" + u
    return u


def append_one_row(csv_path: str, rec: dict):
    df_row = pd.DataFrame([rec])
    file_exists = os.path.exists(csv_path)
    df_row.to_csv(csv_path, mode="a", index=False, header=not file_exists)


def main():
    if not os.path.exists(DATA_APPEND_CSV):
        raise FileNotFoundError(DATA_APPEND_CSV)

    df = pd.read_csv(DATA_APPEND_CSV)
    if "Website" not in df.columns:
        raise ValueError("data_append.csv must include a 'Website' column")

    df["Website"] = df["Website"].apply(normalize_website)
    unique_urls = df["Website"].dropna().unique().tolist()

    already = set()
    if os.path.exists(ENRICHED_CSV):
        done = pd.read_csv(ENRICHED_CSV)
        if "Website" in done.columns:
            already = set(done["Website"].astype(str).tolist())

    print(f"Total unique parcel URLs: {len(unique_urls)}; already scraped: {len(already)}")

    for idx, url in enumerate(unique_urls, start=1):
        if not url or url in already:
            continue

        print(f"[{idx}/{len(unique_urls)}] scraping: {url}")
        try:
            html = fetch_html(url)
            rec = parse_detail(html, url)
            append_one_row(ENRICHED_CSV, rec)
        except Exception as e:
            rec = default_record(url)
            rec["scrape_error"] = str(e)
            rec["ValueChanges_JSON"] = json.dumps([], ensure_ascii=False)
            append_one_row(ENRICHED_CSV, rec)

        time.sleep(DETAIL_SLEEP_SECONDS)


if __name__ == "__main__":
    main()
