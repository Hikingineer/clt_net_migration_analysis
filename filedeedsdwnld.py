import os
import time
import random
import datetime as dt
import requests
import pandas as pd
from io import BytesIO

URL = "https://property.spatialest.com/nc/mecklenburg/api/v1/search/getexcel"

DEEDS_DIR = "deeds"
COMPLETED_CSV = "completed_dates.csv"
APPEND_CSV = "data_append.csv"
NO_ENTRY_CSV = "no_entry_days.csv"

EXPECTED_COLS = ["Address", "Appraised Value", "Parcel", "Owners", "Website"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept": "application/vnd.ms-excel, application/octet-stream, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
    "Referer": "https://property.spatialest.com/",
    "Origin": "https://property.spatialest.com",
}

TIMEOUT_SECONDS = 120
SLEEP_SECONDS = 2
MAX_RETRIES = 8
CAP_THRESHOLD = 300

session = requests.Session()


def day_to_api_str(d: dt.date) -> str:
    return d.strftime("%m-%d-%Y")


def load_completed_dates() -> set[str]:
    if not os.path.exists(COMPLETED_CSV):
        return set()
    df = pd.read_csv(COMPLETED_CSV)
    if "SaleDate" not in df.columns:
        return set()
    return set(df["SaleDate"].astype(str).tolist())


def mark_completed(day: dt.date, rows: int, capped: bool):
    row = pd.DataFrame([{
        "SaleDate": day.isoformat(),
        "Rows": int(rows),
        "Capped": bool(capped),
    }])

    if os.path.exists(COMPLETED_CSV):
        existing = pd.read_csv(COMPLETED_CSV)
        out = pd.concat([existing, row], ignore_index=True)
        out = out.drop_duplicates(subset=["SaleDate"], keep="last")
        out.to_csv(COMPLETED_CSV, index=False)
    else:
        row.to_csv(COMPLETED_CSV, index=False)


def append_to_csv(df_new: pd.DataFrame):
    file_exists = os.path.exists(APPEND_CSV)
    df_new.to_csv(APPEND_CSV, mode="a", index=False, header=not file_exists)


def append_no_entry(day: dt.date, reason: str):
    row = pd.DataFrame([{"SaleDate": day.isoformat(), "Reason": reason}])
    file_exists = os.path.exists(NO_ENTRY_CSV)
    row.to_csv(NO_ENTRY_CSV, mode="a", header=not file_exists, index=False)


def weekday_name(day: dt.date) -> str:
    return ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][day.weekday()]


def should_skip_weekend(day: dt.date) -> bool:
    return day.weekday() in (5, 6)  # Sat=5, Sun=6


def fetch_day_excel_bytes(day: dt.date) -> bytes:
    params = {
        "filters[SaleDate][min]": day_to_api_str(day),
        "filters[SaleDate][max]": day_to_api_str(day),
    }

    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = session.get(URL, params=params, headers=HEADERS, timeout=TIMEOUT_SECONDS)

            if r.status_code in (403, 429):
                sleep_for = (attempt ** 2) + random.random() * 2
                print(f"  Blocked HTTP {r.status_code}. Retry {attempt}/{MAX_RETRIES} after {sleep_for:.1f}s...")
                time.sleep(sleep_for)
                continue

            r.raise_for_status()
            return r.content

        except Exception as e:
            last_err = e
            sleep_for = (attempt ** 2) + random.random() * 2
            print(f"  Error: {e}. Retry {attempt}/{MAX_RETRIES} after {sleep_for:.1f}s...")
            time.sleep(sleep_for)

    raise RuntimeError(f"Failed downloading {day.isoformat()} after retries. Last error: {last_err}")


def month_out_dir(day: dt.date) -> str:
    return os.path.join(DEEDS_DIR, day.strftime("%Y-%m"))


def out_file_path(day: dt.date) -> str:
    return os.path.join(month_out_dir(day), f"{day.isoformat()}.xlsx")


def main():
    start_date = dt.date(2025, 12, 31)
    end_date = dt.date(2011, 1, 1)

    completed = load_completed_dates()

    total = (start_date - end_date).days + 1
    for i in range(total):
        day = start_date - dt.timedelta(days=i)
        key = day.isoformat()

        if key in completed:
            print(f"[{i+1}/{total}] {key} already completed - skipping")
            continue

        print(f"[{i+1}/{total}] {key} ({weekday_name(day)}) ...")

        # Weekend: no requests; log and mark completed so it won't re-run.
        if should_skip_weekend(day):
            append_no_entry(day, reason="Weekend (Sat/Sun) skipped")
            mark_completed(day, rows=0, capped=False)
            print("  Weekend skipped.")
            continue

        # For weekday: try normal flow, but if columns missing, do bypass/log and continue.
        try:
            out_dir = month_out_dir(day)
            os.makedirs(out_dir, exist_ok=True)
            path = out_file_path(day)

            # If you want the "public holiday bypass when the file doesn't exist":
            # - if file doesn't exist, download (normal path)
            # - if file exists already, we can attempt to parse it (no extra request)
            if not os.path.exists(path):
                content = fetch_day_excel_bytes(day)
                with open(path, "wb") as f:
                    f.write(content)
            else:
                # read existing file bytes to keep everything uniform
                with open(path, "rb") as f:
                    content = f.read()

            df = pd.read_excel(BytesIO(content))

            missing = [c for c in EXPECTED_COLS if c not in df.columns]
            if len(df) == 0:
                append_no_entry(day, reason="No rows returned")
                mark_completed(day, rows=0, capped=False)
                print("  No rows returned.")
                continue

            if missing:
                # This is the “public holiday / missing columns” case.
                # Bypass: log, mark completed, continue (no crash).
                append_no_entry(day, reason=f"Missing expected columns: {missing}")
                mark_completed(day, rows=len(df), capped=False)
                print(f"  Missing columns -> logged. {missing}")
                continue

            df = df[EXPECTED_COLS].copy()
            df["SaleDate"] = key

            rows = len(df)
            capped = rows >= CAP_THRESHOLD

            append_to_csv(df)
            mark_completed(day, rows=rows, capped=capped)

            print(f"  Saved/append done. rows={rows}, capped={capped}")

        except Exception as e:
            # Always keep going
            append_no_entry(day, reason=f"Download/parse error: {e}")
            mark_completed(day, rows=0, capped=False)
            print(f"  ERROR -> logged and continuing: {e}")

        time.sleep(SLEEP_SECONDS)


if __name__ == "__main__":
    main()
