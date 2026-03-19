import requests
import pandas as pd
import random
import time
import smtplib
import os
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup
from collections import defaultdict

TODAY = datetime.now()

# -------------------------------------------------
# CONFIG
# -------------------------------------------------

# BUG 1 FIXED: env var names now match the GitHub Secrets defined in GOOGLE_SHEETS_SETUP.md
EMAIL      = os.environ.get("EMAIL")       # was "EMAIL_USER" — secret is named "EMAIL"
EMAIL_PASS = os.environ.get("PASSWORD")    # was "EMAIL_PASS" — secret is named "PASSWORD"
TO_EMAIL   = os.environ.get("EMAIL")       # send digest to yourself; add EMAIL_TO secret if different
GSHEET_ID  = os.environ.get("GSHEET_ID")
SA_JSON    = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

OUTPUT_FILE = "internships.xlsx"

# -------------------------------------------------
# TIER 1 COMPANIES
# -------------------------------------------------

TIER1 = [
    "goldman sachs", "morgan stanley", "jpmorgan", "hsbc", "barclays",
    "citibank", "blackrock", "bcg", "bain", "mckinsey",
    "kotak", "axis bank", "hdfc", "icici"
]

# -------------------------------------------------
# BLOCKED LINKS (SPECIFIC OLD INTERNSHIPS)
# -------------------------------------------------

EXCLUDED_LINKS = {
    "https://www.jobaaj.com/job/barclays-corporate-finance-intern-delhi-ncr-0-to-1-years-760847",
    "https://www.jobaaj.com/job/hsbc-investment-banking-intern-delhi-ncr-0-to-1-years-670612"
}

# -------------------------------------------------
# PROXY ROTATION
# -------------------------------------------------

PROXIES = [None]

def get_proxy():
    p = random.choice(PROXIES)
    if not p:
        return None
    return {"http": p, "https": p}

# -------------------------------------------------
# RETRY LOGIC
# BUG 2 FIXED: sleep reduced from 30s to 8s so CI doesn't waste time
# -------------------------------------------------

def safe_request(url, headers=None):
    for i in range(2):
        try:
            r = requests.get(
                url,
                headers=headers,
                proxies=get_proxy(),
                timeout=20
            )
            if r.status_code == 200:
                return r
            else:
                print(f"[WARN] {url} returned status {r.status_code}")
        except Exception as e:
            print(f"[ERROR] Request failed for {url}: {e}")
        time.sleep(8)   # was 30 — excessive on CI
    return None

# -------------------------------------------------
# GOOGLE SHEETS: LOAD + SAVE SEEN LISTINGS
# BUG 6 FIXED: Google Sheets deduplication was completely missing in v6
# -------------------------------------------------

def get_gsheet_client():
    if not SA_JSON:
        print("[WARN] GOOGLE_SERVICE_ACCOUNT_JSON not set — cross-run dedup disabled")
        return None
    try:
        info = json.loads(SA_JSON)
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(info, scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        print(f"[ERROR] Failed to build gsheet client: {e}")
        return None

def load_seen_links():
    """Return a set of listing links already emailed in past runs."""
    client = get_gsheet_client()
    if not client or not GSHEET_ID:
        return set()
    try:
        sh = client.open_by_key(GSHEET_ID)
        try:
            ws = sh.worksheet("seen_listings")
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title="seen_listings", rows=1, cols=1)
            ws.append_row(["link"])
            return set()
        values = ws.col_values(1)
        return set(v.strip() for v in values[1:] if v.strip())  # skip header
    except Exception as e:
        print(f"[ERROR] load_seen_links failed: {e}")
        return set()

def save_new_links(new_links):
    """Append newly seen links to the Google Sheet for future dedup."""
    client = get_gsheet_client()
    if not client or not GSHEET_ID or not new_links:
        return
    try:
        sh = client.open_by_key(GSHEET_ID)
        try:
            ws = sh.worksheet("seen_listings")
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title="seen_listings", rows=1, cols=1)
            ws.append_row(["link"])
        for link in new_links:
            ws.append_row([link])
        print(f"[INFO] Saved {len(new_links)} new links to Google Sheet")
    except Exception as e:
        print(f"[ERROR] save_new_links failed: {e}")

# -------------------------------------------------
# SHINE SCRAPER
# -------------------------------------------------

def scrape_shine():
    jobs = []
    url = "https://www.shine.com/job-search/finance-intern-jobs"
    headers = {"User-Agent": "Mozilla/5.0"}
    r = safe_request(url, headers=headers)
    if not r:
        print("[WARN] Shine scraper got no response")
        return jobs

    soup = BeautifulSoup(r.text, "html.parser")
    cards = soup.select("li.job-listing") or soup.select("div[class*='jobCard']")

    for c in cards:
        try:
            title_el = c.select_one("a[class*='title'], h2, h3")
            company_el = c.select_one("[class*='company'], [class*='employer']")
            link_el = c.select_one("a[href]")

            if not title_el or not link_el:
                continue

            title = title_el.text.strip()
            company = company_el.text.strip() if company_el else "Unknown"
            href = link_el["href"]
            link = href if href.startswith("http") else "https://www.shine.com" + href

            jobs.append({
                "title": title,
                "company": company,
                "location": "India",
                "link": link,
                "source": "Shine",
                "deadline": "",
                "posted_date": TODAY.strftime("%Y-%m-%d")
            })
        except Exception as e:
            print(f"[WARN] Shine card parse error: {e}")

    print(f"[INFO] Shine scraped {len(jobs)} jobs")
    return jobs

# -------------------------------------------------
# TIMESJOBS SCRAPER
# -------------------------------------------------

def scrape_timesjobs():
    jobs = []
    url = "https://www.timesjobs.com/candidate/job-search.html?txtKeywords=finance+intern"
    headers = {"User-Agent": "Mozilla/5.0"}
    r = safe_request(url, headers=headers)
    if not r:
        print("[WARN] TimesJobs scraper got no response")
        return jobs

    soup = BeautifulSoup(r.text, "html.parser")
    cards = soup.select("li.clearfix.job-bx.wht-shd-bx") or soup.select("li[class*='job-bx']")

    for c in cards:
        try:
            title_el = c.select_one("h2 a") or c.select_one("h2")
            company_el = c.select_one("h3.joblist-comp-name") or c.select_one("[class*='company']")
            link_el = c.select_one("h2 a")

            if not title_el:
                continue

            title = title_el.text.strip()
            company = company_el.text.strip() if company_el else "Unknown"
            link = link_el["href"] if link_el and link_el.get("href") else url

            jobs.append({
                "title": title,
                "company": company,
                "location": "India",
                "link": link,
                "source": "TimesJobs",
                "deadline": "",
                "posted_date": TODAY.strftime("%Y-%m-%d")
            })
        except Exception as e:
            print(f"[WARN] TimesJobs card parse error: {e}")

    print(f"[INFO] TimesJobs scraped {len(jobs)} jobs")
    return jobs

# -------------------------------------------------
# FILTER: REMOVE EXCLUDED LINKS
# -------------------------------------------------

def remove_excluded_jobs(jobs):
    return [j for j in jobs if j.get("link", "").strip() not in EXCLUDED_LINKS]

# -------------------------------------------------
# FILTER: REMOVE OLD JOBAAJ INTERNSHIPS (>6 MONTHS)
# BUG 3 FIXED: posted_date is now a string "YYYY-MM-DD"; parse it for comparison
# -------------------------------------------------

def remove_old_jobaaj(jobs):
    cutoff = TODAY - timedelta(days=180)
    filtered = []
    for j in jobs:
        if "jobaaj.com" in j.get("link", ""):
            try:
                posted = datetime.strptime(j.get("posted_date", ""), "%Y-%m-%d")
                if posted < cutoff:
                    continue
            except ValueError:
                pass  # if date missing/malformed, keep the job
        filtered.append(j)
    return filtered

# -------------------------------------------------
# DEDUPLICATION (in-run, by link)
# -------------------------------------------------

def deduplicate(jobs):
    seen = set()
    unique = []
    for j in jobs:
        link = j.get("link", "").strip()
        if link and link not in seen:
            seen.add(link)
            unique.append(j)
    print(f"[INFO] After dedup: {len(unique)} unique jobs")
    return unique

# -------------------------------------------------
# DOMAIN CLASSIFIER
# -------------------------------------------------

def classify_domain(title):
    t = title.lower()
    if "consult" in t or "strategy" in t:
        return "Consulting"
    if "finance" in t or "investment" in t or "banking" in t:
        return "Finance"
    if "research" in t or "analyst" in t:
        return "Research"
    return "Other"

# -------------------------------------------------
# TIER1 CHECK
# -------------------------------------------------

def is_tier1(company):
    c = company.lower()
    return any(t in c for t in TIER1)

# -------------------------------------------------
# EMAIL DIGEST
# -------------------------------------------------

def send_email(jobs, new_count):
    if not all([EMAIL, EMAIL_PASS, TO_EMAIL]):
        print("[ERROR] Email credentials missing — skipping email send")
        print(f"  EMAIL={EMAIL}, EMAIL_PASS={'set' if EMAIL_PASS else 'MISSING'}, TO_EMAIL={TO_EMAIL}")
        return

    grouped = defaultdict(list)
    for j in jobs:
        grouped[j["domain"]].append(j)

    html = f"<h2>Internship Digest — {new_count} NEW today</h2>"
    for d, items in grouped.items():
        html += f"<h3>{d}</h3>"
        for j in items:
            tier_badge = " ⭐ <b>Tier 1</b>" if j.get("tier1") else ""
            is_new = " 🆕" if j.get("is_new") else ""
            html += f"""
            <p>
            <b>{j['title']}</b>{tier_badge}{is_new} – {j['company']}<br>
            {j['location']}<br>
            <a href="{j['link']}"
            style="background:#0073e6;color:white;padding:6px 12px;border-radius:6px;text-decoration:none">
            Apply Now
            </a>
            </p>
            """

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"{new_count} NEW Internships – {TODAY.strftime('%d %b %Y')}"
        msg["From"] = EMAIL
        msg["To"] = TO_EMAIL
        msg.attach(MIMEText(html, "html"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL, EMAIL_PASS)
            server.sendmail(EMAIL, TO_EMAIL, msg.as_string())

        print("[INFO] Email sent successfully")
    except Exception as e:
        print(f"[ERROR] Failed to send email: {e}")

# -------------------------------------------------
# MAIN
# -------------------------------------------------

def main():
    jobs = []
    jobs += scrape_shine()
    jobs += scrape_timesjobs()

    print(f"[INFO] Total scraped: {len(jobs)}")

    jobs = remove_excluded_jobs(jobs)
    jobs = remove_old_jobaaj(jobs)
    jobs = deduplicate(jobs)

    if not jobs:
        print("[WARN] No jobs found — skipping Excel export and email")
        return

    for j in jobs:
        j["domain"] = classify_domain(j["title"])
        j["tier1"] = is_tier1(j["company"])

    # BUG 6 FIXED: cross-run dedup via Google Sheets (restored from v5 design)
    seen_links = load_seen_links()
    new_jobs = []
    for j in jobs:
        link = j.get("link", "").strip()
        if link not in seen_links:
            j["is_new"] = True
            new_jobs.append(j)
        else:
            j["is_new"] = False

    save_new_links([j["link"] for j in new_jobs])
    new_count = len(new_jobs)
    print(f"[INFO] {new_count} new listings this run (not seen before)")

    df = pd.DataFrame(jobs)
    tier1_df = df[df["tier1"] == True]

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All Listings", index=False)
        tier1_df.to_excel(writer, sheet_name="Tier1 Only", index=False)

    print(f"[INFO] Excel saved: {OUTPUT_FILE} ({len(df)} rows, {len(tier1_df)} Tier1)")

    send_email(jobs, new_count)

if __name__ == "__main__":
    main()
