import json
import re
from datetime import datetime, date
from pathlib import Path

import pandas as pd

INPUT_XLSX = Path("events.xlsx")
OUTPUT_JSON = Path("events.json")

REQUIRED_COLS = ["Event", "Date", "Location", "Handle"]

def normalize_date(value) -> str:
    if pd.isna(value):
        raise ValueError("Missing Date")

    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()

    s = str(value).strip()

    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass

    try:
        return datetime.strptime(s, "%B %d, %Y").date().isoformat()
    except ValueError:
        pass

    raise ValueError(f"Unrecognized Date format: {value!r}. Use YYYY-MM-DD.")

def normalize_handle(handle: str) -> str:
    s = str(handle).strip()
    if not s:
        raise ValueError("Missing Handle")

    s = re.sub(r"^https?://[^/]+/", "", s)
    s = re.sub(r"^products/", "", s)
    s = re.sub(r"^/products/", "", s)
    s = s.strip("/")

    if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", s):
        raise ValueError(
            f"Handle looks invalid: {handle!r}. Use lowercase letters, numbers, and hyphens only."
        )

    return s

def main():
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(
            f"Missing {INPUT_XLSX}. Put your spreadsheet in this folder and name it events.xlsx."
        )

    df = pd.read_excel(INPUT_XLSX)

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Spreadsheet is missing columns: {missing}. Required: {REQUIRED_COLS}")

    events = []
    errors = []

    for i, row in df.iterrows():
        try:
            event = str(row["Event"]).strip()
            location = str(row["Location"]).strip()
            date_iso = normalize_date(row["Date"])
            handle = normalize_handle(row["Handle"])

            if not event:
                raise ValueError("Missing Event")
            if not location:
                raise ValueError("Missing Location")

            events.append({
                "event": event,
                "date": date_iso,
                "location": location,
                "handle": handle
            })
        except Exception as e:
            errors.append(f"Row {i + 2}: {e}")

    if errors:
        print("Fix these spreadsheet issues and rerun:")
        for err in errors:
            print(err)
        raise SystemExit(1)

    events.sort(key=lambda x: x["date"])

    OUTPUT_JSON.write_text(json.dumps(events, indent=2), encoding="utf-8")
    print(f"Wrote {OUTPUT_JSON} with {len(events)} events.")

if __name__ == "__main__":
    main()