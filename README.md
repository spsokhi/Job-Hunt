# Multi-Role IT Job Scraper

Automates a daily IT job hunt. It searches several job titles at once via the
free [Adzuna Jobs API](https://developer.adzuna.com/), tags each result with a
**Category** (Dev / Support roles), filters out roles you don't want, and exports
everything into a clean Excel workbook with **one tab per category** so your
daily application process stays organized.

## Why the Adzuna API (and not scraping Indeed/LinkedIn)?

Direct HTML scraping of Indeed, LinkedIn, etc. **violates their Terms of Service**,
breaks constantly (anti-bot walls, CAPTCHAs, JavaScript-rendered pages), and
rarely yields clean structured data. The Adzuna API is **free**, **legal**, and
returns title / company / location / URL / date as proper JSON — exactly the
fields you need.

---

## 1. Setup

```powershell
# From the project folder (c:\code\JobHunt)
python -m venv venv
.\venv\Scripts\Activate.ps1          # PowerShell
pip install -r requirements.txt
```

> On macOS/Linux the activate step is `source venv/bin/activate`.

## 2. Get free Adzuna API credentials (~2 minutes)

1. Sign up at **https://developer.adzuna.com/signup**
2. Open your dashboard — you'll see an **Application ID** and **Application Key**.

Then provide them to the script in **either** of these ways:

**Option A — environment variables (PowerShell):**
```powershell
$env:ADZUNA_APP_ID  = "your_app_id"
$env:ADZUNA_APP_KEY = "your_app_key"
```

**Option B — a `.env` file** (auto-loaded; create it next to `job_scraper.py`):
```
ADZUNA_APP_ID=your_app_id
ADZUNA_APP_KEY=your_app_key
```

## 3. Run it

```powershell
python job_scraper.py
```

You'll get a file like `jobs_2026-06-18.xlsx` with these sheets:

| Sheet             | Contents                                         |
|-------------------|--------------------------------------------------|
| **All Jobs**      | Everything found, newest first                   |
| **Dev Roles**     | Software Engineer, Python Developer              |
| **Support Roles** | Technical Support, Customer Support, IT Helpdesk |

URLs in the **URL** column are clickable — open the workbook and apply straight from it.

---

## 4. Customizing (all in the `CONFIG` block at the top of `job_scraper.py`)

**Add job titles / change categories** — just edit the dictionary. A brand-new
category name automatically becomes a new Excel tab:
```python
JOB_TITLES = {
    "Software Engineer":  "Dev Roles",
    "Python Developer":   "Dev Roles",
    "Data Analyst":       "Data Roles",   # <- new title AND new sheet
    "IT Helpdesk":        "Support Roles",
}
```

**Add or change locations** — `""` means search the whole country:
```python
LOCATIONS = ["Remote", "New York", "Austin", ""]
```

**Change the country** — Adzuna codes: `us`, `gb`, `ca`, `au`, `in`, `de`, `fr`, ...
```python
COUNTRY = "gb"
```

**Tune the filter** — add words/phrases to exclude over- or under-qualified roles:
```python
NEGATIVE_KEYWORDS = ["senior", "lead", "manager", "10+ years", "clearance"]
```

**Pull more results per search** — `MAX_PAGES = 2` returns up to 100 per query
(50 per page). Raise it for deeper searches.

---

## Notes & limits

- Adzuna's free tier is generous for personal daily use; the built-in 1-second
  delay between calls keeps you well under the rate limit.
- Re-running creates one file per day (timestamped), so you keep a history.
- If you get **0 results**, loosen `NEGATIVE_KEYWORDS` or add more `LOCATIONS`.
- A `401` error means the API ID/key are wrong or not yet active.
