import requests
import pandas as pd
import random
import time
import smtplib
import os
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup
from collections import defaultdict

TODAY = datetime.now()

# -------------------------------------------------
# CONFIG
# -------------------------------------------------

EMAIL = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
TO_EMAIL = os.environ.get("EMAIL_TO")

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
# RETRY LOGIC  (FIX 6: log errors instead of silently swallowing)
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
        time.sleep(30)
    return None

# -------------------------------------------------
# SHINE SCRAPER  (FIX 3: corrected CSS selectors)
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

    # Shine's actual listing cards as of 2024 — update selectors if site changes
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
                "posted_date": TODAY   # FIX 4 note: real date parsing would need page detail scrape
            })
        except Exception as e:
            print(f"[WARN] Shine card parse error: {e}")

    print(f"[INFO] Shine scraped {len(jobs)} jobs")
    return jobs

# -------------------------------------------------
# TIMESJOBS SCRAPER  (FIX 3: corrected CSS selectors)
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

    # TimesJobs actual listing structure
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
                "posted_date": TODAY
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
# -------------------------------------------------

def remove_old_jobaaj(jobs):
    cutoff = TODAY - timedelta(days=180)
    filtered = []
    for j in jobs:
        if "jobaaj.com" in j.get("link", ""):
            posted = j.get("posted_date")
            if posted and posted < cutoff:
                continue
        filtered.append(j)
    return filtered

# -------------------------------------------------
# DEDUPLICATION  (FIX 8: remove duplicate links)
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
# EMAIL DIGEST  (FIX 1: guard against missing creds + empty jobs)
# -------------------------------------------------

def send_email(jobs):
    # Guard: check credentials exist
    if not all([EMAIL, EMAIL_PASS, TO_EMAIL]):
        print("[ERROR] Email credentials missing — skipping email send")
        return

    grouped = defaultdict(list)
    for j in jobs:
        grouped[j["domain"]].append(j)

    html = "<h2>Internship Digest</h2>"
    for d, items in grouped.items():
        html += f"<h3>{d}</h3>"
        for j in items:
            tier_badge = " ⭐ <b>Tier 1</b>" if j.get("tier1") else ""
            html += f"""
            <p>
            <b>{j['title']}</b>{tier_badge} – {j['company']}<br>
            {j['location']}<br>
            <a href="{j['link']}"
            style="background:#0073e6;color:white;padding:6px 12px;border-radius:6px;text-decoration:none">
            Apply Now
            </a>
            </p>
            """

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"Internship Digest – {TODAY.strftime('%d %b %Y')}"
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
    jobs = deduplicate(jobs)   # FIX 8

    # FIX 2: guard against empty jobs list before DataFrame ops
    if not jobs:
        print("[WARN] No jobs found — skipping Excel export and email")
        return

    for j in jobs:
        j["domain"] = classify_domain(j["title"])
        j["tier1"] = is_tier1(j["company"])

    df = pd.DataFrame(jobs)
    tier1_df = df[df["tier1"] == True]

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All Listings", index=False)
        tier1_df.to_excel(writer, sheet_name="Tier1 Only", index=False)

    print(f"[INFO] Excel saved: {OUTPUT_FILE} ({len(df)} rows, {len(tier1_df)} Tier1)")

    # FIX 7: only email if we actually have jobs
    send_email(jobs)

if __name__ == "__main__":
    main()
