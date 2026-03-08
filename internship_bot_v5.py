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
    "goldman sachs",
    "morgan stanley",
    "jpmorgan",
    "hsbc",
    "barclays",
    "citibank",
    "blackrock",
    "bcg",
    "bain",
    "mckinsey",
    "kotak",
    "axis bank",
    "hdfc",
    "icici"
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

PROXIES = [
    None
]

def get_proxy():
    p = random.choice(PROXIES)
    if not p:
        return None
    return {"http": p, "https": p}

# -------------------------------------------------
# RETRY LOGIC
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

        except:
            pass

        time.sleep(30)

    return None

# -------------------------------------------------
# SHINE SCRAPER
# -------------------------------------------------

def scrape_shine():

    jobs = []

    url = "https://www.shine.com/job-search/finance-intern-jobs"

    r = safe_request(url)

    if not r:
        return jobs

    soup = BeautifulSoup(r.text, "html.parser")

    cards = soup.select(".jobCard")

    for c in cards:

        try:

            title = c.select_one(".jobTitle").text.strip()
            company = c.select_one(".companyName").text.strip()
            link = "https://www.shine.com" + c.select_one("a")["href"]

            jobs.append({
                "title": title,
                "company": company,
                "location": "India",
                "link": link,
                "source": "Shine",
                "deadline": "",
                "posted_date": TODAY
            })

        except:
            pass

    return jobs

# -------------------------------------------------
# TIMESJOBS SCRAPER
# -------------------------------------------------

def scrape_timesjobs():

    jobs = []

    url = "https://www.timesjobs.com/candidate/job-search.html?txtKeywords=finance+intern"

    r = safe_request(url)

    if not r:
        return jobs

    soup = BeautifulSoup(r.text, "html.parser")

    cards = soup.select(".job-bx")

    for c in cards:

        try:

            title = c.select_one("h2").text.strip()
            company = c.select_one(".company-name").text.strip()
            link = c.select_one("a")["href"]

            jobs.append({
                "title": title,
                "company": company,
                "location": "India",
                "link": link,
                "source": "TimesJobs",
                "deadline": "",
                "posted_date": TODAY
            })

        except:
            pass

    return jobs

# -------------------------------------------------
# FILTER: REMOVE EXCLUDED LINKS
# -------------------------------------------------

def remove_excluded_jobs(jobs):

    filtered = []

    for j in jobs:

        link = j.get("link", "").strip()

        if link in EXCLUDED_LINKS:
            continue

        filtered.append(j)

    return filtered

# -------------------------------------------------
# FILTER: REMOVE OLD JOBAAJ INTERNSHIPS (>6 MONTHS)
# -------------------------------------------------

def remove_old_jobaaj(jobs):

    filtered = []

    cutoff = TODAY - timedelta(days=180)

    for j in jobs:

        link = j.get("link", "")

        if "jobaaj.com" in link:

            posted = j.get("posted_date")

            if posted and posted < cutoff:
                continue

        filtered.append(j)

    return filtered

# -------------------------------------------------
# DOMAIN CLASSIFIER
# -------------------------------------------------

def classify_domain(title):

    t = title.lower()

    if "consult" in t or "strategy" in t:
        return "Consulting"

    if "finance" in t or "investment" in t:
        return "Finance"

    if "research" in t:
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

def send_email(jobs):

    grouped = defaultdict(list)

    for j in jobs:
        grouped[j["domain"]].append(j)

    html = "<h2>Internship Digest</h2>"

    for d, items in grouped.items():

        html += f"<h3>{d}</h3>"

        for j in items:

            html += f"""
            <p>
            <b>{j['title']}</b> – {j['company']}<br>
            {j['location']}<br>
            <a href="{j['link']}" 
            style="background:#0073e6;color:white;padding:6px 12px;border-radius:6px;text-decoration:none">
            Apply Now
            </a>
            </p>
            """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Internship Digest"
    msg["From"] = EMAIL
    msg["To"] = TO_EMAIL

    msg.attach(MIMEText(html, "html"))

    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login(EMAIL, EMAIL_PASS)
    server.sendmail(EMAIL, TO_EMAIL, msg.as_string())
    server.quit()

# -------------------------------------------------
# MAIN
# -------------------------------------------------

def main():

    jobs = []

    jobs += scrape_shine()
    jobs += scrape_timesjobs()

    print("Jobs scraped:", len(jobs))

    # Remove specific old internships
    jobs = remove_excluded_jobs(jobs)

    # Remove old jobaaj listings
    jobs = remove_old_jobaaj(jobs)

    for j in jobs:

        j["domain"] = classify_domain(j["title"])
        j["tier1"] = is_tier1(j["company"])

    df = pd.DataFrame(jobs)

    tier1_df = df[df["tier1"] == True]

    with pd.ExcelWriter(OUTPUT_FILE) as writer:

        df.to_excel(writer, sheet_name="All Listings", index=False)
        tier1_df.to_excel(writer, sheet_name="Tier1 Only", index=False)

    send_email(jobs)

if __name__ == "__main__":
    main()
