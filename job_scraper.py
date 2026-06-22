"""
Multi-Role IT Job Scraper
=========================

Searches multiple job titles via the free Adzuna Jobs API, categorizes the
results (Dev / Support roles), filters out roles you don't want using a list of
negative keywords, and exports everything to a tidy multi-sheet Excel workbook.

Why Adzuna instead of scraping Indeed/LinkedIn directly?
    Direct HTML scraping of those sites violates their Terms of Service, breaks
    constantly (anti-bot walls, CAPTCHAs, JavaScript-rendered pages) and rarely
    yields clean structured data. The Adzuna API is free, legal, and returns
    title / company / location / URL / date as proper JSON. See README.md for
    how to get your free credentials (takes ~2 minutes).

Run:
    python job_scraper.py

Configure: edit the CONFIG section below. Everything you'd normally want to
change (job titles, categories, locations, negative keywords) lives there.
"""

from __future__ import annotations

import os
import sys
import time
from datetime import datetime

import requests

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas is required. Run:  pip install -r requirements.txt")

# Optional: load credentials from a .env file if python-dotenv is installed.
try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass


# ---------------------------------------------------------------------------
# CONFIG  —  this is the only section you normally need to touch
# ---------------------------------------------------------------------------

# Adzuna country code. Common: "us", "gb", "ca", "au", "in", "de", "fr".
COUNTRY = "in"

# Map each job title you want to search -> the category (Excel sheet) it lands in.
# Add or remove lines freely. New categories automatically become new sheets.
# Focused on support / helpdesk roles, with several title variants so we catch
# postings worded differently across companies.
JOB_TITLES = {
    "Technical Support Specialist": "Support Roles",
    "Customer Support Engineer":    "Support Roles",
    "IT Helpdesk":                  "Support Roles",
    "Help Desk Support":            "Support Roles",
    "Desktop Support":              "Support Roles",
    "IT Support":                   "Support Roles",
    "Service Desk":                 "Support Roles",
    "Technical Support Engineer":   "Support Roles",
}

# Locations to search. Use "" (empty string) to search the whole country.
# We want Bangalore-based OR remote roles.
LOCATIONS = [
    "Bangalore",
    "Remote",
]

# Any job whose TITLE or DESCRIPTION contains one of these (case-insensitive)
# is excluded. Tune this to your experience level.
NEGATIVE_KEYWORDS = [
    "senior",
    "sr.",
    "lead",
    "staff",
    "principal",
    "manager",
    "director",
    "head of",
    "vp ",
    "architect",
    "10+ years",
    "8+ years",
    "7+ years",
    "security clearance",
]

# How many results to pull per (title, location) query, and how many API pages
# to walk (Adzuna returns up to 50 per page). MAX_PAGES=2 => up to 100 per query.
RESULTS_PER_PAGE = 50
MAX_PAGES = 2

# Politeness delay between API calls (seconds). Keeps you under rate limits.
REQUEST_DELAY = 1.0

# Output file. A timestamp keeps each day's hunt as its own file.
OUTPUT_FILE = f"support_jobs_{datetime.now():%Y-%m-%d}.xlsx"

# ---------------------------------------------------------------------------
# End of CONFIG
# ---------------------------------------------------------------------------


API_BASE = "https://api.adzuna.com/v1/api/jobs"


def get_credentials() -> tuple[str, str]:
    """Read Adzuna API credentials from environment variables."""
    app_id = os.getenv("ADZUNA_APP_ID")
    app_key = os.getenv("ADZUNA_APP_KEY")
    if not app_id or not app_key:
        sys.exit(
            "Missing Adzuna credentials.\n"
            "  1. Sign up (free) at https://developer.adzuna.com/\n"
            "  2. Set them as environment variables, e.g. in PowerShell:\n"
            '       $env:ADZUNA_APP_ID  = "your_app_id"\n'
            '       $env:ADZUNA_APP_KEY = "your_app_key"\n'
            "     ...or create a .env file (see README.md)."
        )
    return app_id, app_key


def is_excluded(*texts: str) -> bool:
    """True if any negative keyword appears in any of the given text fields."""
    haystack = " ".join(t for t in texts if t).lower()
    return any(kw.lower() in haystack for kw in NEGATIVE_KEYWORDS)


def fetch_page(app_id, app_key, query, location, page) -> list[dict]:
    """Fetch a single page of results for one (query, location) pair."""
    url = f"{API_BASE}/{COUNTRY}/search/{page}"
    params = {
        "app_id": app_id,
        "app_key": app_key,
        "results_per_page": RESULTS_PER_PAGE,
        "what": query,
        "content-type": "application/json",
    }
    if location:
        params["where"] = location

    try:
        resp = requests.get(url, params=params, timeout=30)
        resp.raise_for_status()
    except requests.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "?"
        if status == 401:
            sys.exit("Adzuna rejected your credentials (401). Check APP_ID / APP_KEY.")
        print(f"   ! HTTP {status} for '{query}' in '{location or 'anywhere'}' "
              f"(page {page}) — skipping.")
        return []
    except requests.RequestException as exc:
        print(f"   ! Network error: {exc} — skipping.")
        return []

    return resp.json().get("results", [])


def normalize(raw: dict, query: str, category: str) -> dict:
    """Pull the fields we care about out of a raw Adzuna result."""
    created = raw.get("created", "")
    # Adzuna dates look like '2024-06-01T12:34:56Z' -> keep just the date.
    date_posted = created[:10] if created else ""
    return {
        "Job Title":   raw.get("title", "").replace("<strong>", "").replace("</strong>", "").strip(),
        "Company":     (raw.get("company") or {}).get("display_name", ""),
        "Location":    (raw.get("location") or {}).get("display_name", ""),
        "Date Posted": date_posted,
        "URL":         raw.get("redirect_url", ""),
        "Matched Search": query,
        "Category":    category,
    }


def scrape_all(app_id, app_key) -> "pd.DataFrame":
    """Run every (title, location) search, filter, dedupe, and return a DataFrame."""
    rows: list[dict] = []
    seen: set[str] = set()
    excluded_count = 0

    for query, category in JOB_TITLES.items():
        for location in LOCATIONS:
            loc_label = location or "anywhere"
            print(f"-> Searching '{query}' in {loc_label} [{category}]")

            for page in range(1, MAX_PAGES + 1):
                results = fetch_page(app_id, app_key, query, location, page)
                if not results:
                    break

                for raw in results:
                    job = normalize(raw, query, category)

                    # Skip jobs matching negative keywords.
                    if is_excluded(job["Job Title"], raw.get("description", "")):
                        excluded_count += 1
                        continue

                    # Dedupe (same posting often appears across searches).
                    key = job["URL"] or f'{job["Job Title"]}|{job["Company"]}|{job["Location"]}'
                    if key in seen:
                        continue
                    seen.add(key)
                    rows.append(job)

                time.sleep(REQUEST_DELAY)

    print(f"\nKept {len(rows)} jobs. Filtered out {excluded_count} by negative keywords.")
    columns = ["Job Title", "Company", "Location", "Date Posted",
               "URL", "Matched Search", "Category"]
    return pd.DataFrame(rows, columns=columns)


def export_to_excel(df: "pd.DataFrame", filename: str) -> None:
    """Write one sheet per category, plus an 'All Jobs' overview sheet."""
    if df.empty:
        print("No jobs to export — try loosening NEGATIVE_KEYWORDS or adding LOCATIONS.")
        return

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        # Combined sheet first.
        _write_sheet(writer, "All Jobs", df)
        # One sheet per category, in the order they first appear.
        for category in dict.fromkeys(JOB_TITLES.values()):
            subset = df[df["Category"] == category]
            if not subset.empty:
                _write_sheet(writer, category, subset)

    print(f"Saved {len(df)} jobs to {os.path.abspath(filename)}")


def _write_sheet(writer, sheet_name: str, frame: "pd.DataFrame") -> None:
    """Write a frame to a sheet, auto-size columns, freeze the header, link URLs."""
    # Excel sheet names max 31 chars and can't contain : \ / ? * [ ]
    safe = sheet_name[:31]
    frame = frame.sort_values("Date Posted", ascending=False)
    frame.to_excel(writer, sheet_name=safe, index=False)

    sheet = writer.sheets[safe]
    sheet.freeze_panes = "A2"

    url_col = list(frame.columns).index("URL") + 1  # 1-based for openpyxl
    for col_idx, col_name in enumerate(frame.columns, start=1):
        # Auto-fit column width to the longest value (capped so URLs stay sane).
        longest = max(
            [len(str(col_name))] + [len(str(v)) for v in frame[col_name]],
            default=10,
        )
        sheet.column_dimensions[sheet.cell(row=1, column=col_idx).coordinate[0]].width = min(longest + 2, 60)

    # Turn the URL cells into clickable hyperlinks.
    for row in range(2, len(frame) + 2):
        cell = sheet.cell(row=row, column=url_col)
        if cell.value:
            cell.hyperlink = cell.value
            cell.style = "Hyperlink"


def main() -> None:
    app_id, app_key = get_credentials()
    print(f"Multi-Role IT Job Scraper — country: {COUNTRY}\n")
    df = scrape_all(app_id, app_key)
    export_to_excel(df, OUTPUT_FILE)


if __name__ == "__main__":
    main()
