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

RAPIDAPI_KEY = os.environ.get("RAPIDAPI_KEY")

OUTPUT_FILE = "internships.xlsx"

KEYWORDS = [
    "finance intern",
    "investment banking intern",
    "equity research intern",
    "consulting intern",
    "strategy intern",
]

LOCATIONS = ["India", "Mumbai", "Delhi", "Bangalore"]

# -------------------------------------------------
# TIER 1 COMPANIES
# -------------------------------------------------

TIER1 = [
    "goldman sachs",
    "morgan stanley",
    "jpmorgan",
    "jp morgan",
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
# PROXY ROTATION
# -------------------------------------------------

PROXIES = [
    None,
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

        except Exception:
            pass

        time.sleep(30)

    return None

# -------------------------------------------------
# LINKEDIN API
# -------------------------------------------------

def scrape_linkedin():

    jobs = []

    url = "https://linkedin-jobs-search.p.rapidapi.com/"

    headers = {
        "X-RapidAPI-Key": RAPIDAPI_KEY,
        "X-RapidAPI-Host": "linkedin-jobs-search.p.rapidapi.com"
    }

    for kw in KEYWORDS:

        params = {
            "keywords": kw,
            "location": "India",
            "datePosted": "past24Hours"
        }

        try:

            r = requests.get(url, headers=headers, params=params)
            data = r.json()

            for j in data:

                jobs.append({
                    "title": j.get("title"),
                    "company": j.get("company"),
                    "location": j.get("location"),
                    "link": j.get("url"),
                    "source": "LinkedIn",
                    "deadline": ""
                })

        except:
            pass

    return jobs

# -------------------------------------------------
# SHINE
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
                "deadline": ""
            })

        except:
            pass

    return jobs

# -------------------------------------------------
# TIMESJOBS
# -------------------------------------------------

def scrape_timesjobs():

    jobs = []

    url = "https://www.timesjobs.com/candidate/job-search.html?searchType=personalizedSearch&txtKeywords=finance+intern"

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
                "deadline": ""
            })

        except:
            pass

    return jobs

# -------------------------------------------------
# APNA
# -------------------------------------------------

def scrape_apna():

    jobs = []

    url = "https://apna.co/jobs?query=intern"

    r = safe_request(url)

    if not r:
        return jobs

    soup = BeautifulSoup(r.text, "html.parser")

    cards = soup.select(".JobCard")

    for c in cards:

        try:

            title = c.select_one(".JobCard_jobTitle").text.strip()
            company = c.select_one(".JobCard_companyName").text.strip()

            link = "https://apna.co"

            jobs.append({
                "title": title,
                "company": company,
                "location": "India",
                "link": link,
                "source": "Apna",
                "deadline": ""
            })

        except:
            pass

    return jobs

# -------------------------------------------------
# IIM BOARDS
# -------------------------------------------------

def scrape_iim():

    jobs = []

    boards = [
        "https://www.iima.ac.in",
        "https://www.iimb.ac.in",
        "https://www.iimcal.ac.in"
    ]

    for b in boards:

        try:

            r = safe_request(b)

            if not r:
                continue

            soup = BeautifulSoup(r.text, "html.parser")

            links = soup.find_all("a")

            for l in links:

                txt = l.text.lower()

                if "intern" in txt:

                    jobs.append({
                        "title": l.text.strip(),
                        "company": "Various",
                        "location": "India",
                        "link": l.get("href"),
                        "source": "IIM Board",
                        "deadline": ""
                    })

        except:
            pass

    return jobs

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
# DEADLINE COUNTDOWN
# -------------------------------------------------

def days_left(deadline):

    try:

        d = datetime.strptime(deadline, "%Y-%m-%d")

        return (d - TODAY).days

    except:

        return ""

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
# TIER1 ALERT
# -------------------------------------------------

def send_tier1_alert(jobs):

    if not jobs:
        return

    html = "<h2>🚨 TIER 1 INTERNSHIP ALERT</h2>"

    for j in jobs:

        html += f"""
        <p>
        <b>{j['company']}</b> – {j['title']}<br>
        <a href="{j['link']}">APPLY NOW</a>
        </p>
        """

    msg = MIMEMultipart("alternative")

    msg["Subject"] = "🚨 Tier1 Internship Alert"
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

    jobs += scrape_linkedin()
    jobs += scrape_shine()
    jobs += scrape_timesjobs()
    jobs += scrape_apna()
    jobs += scrape_iim()

    print("Jobs scraped:", len(jobs))

    for j in jobs:

        j["domain"] = classify_domain(j["title"])

        j["tier1"] = is_tier1(j["company"])

        j["days_left"] = days_left(j["deadline"])

    df = pd.DataFrame(jobs)

    tier1_df = df[df["tier1"] == True]

    new_df = df

    with pd.ExcelWriter(OUTPUT_FILE) as writer:

        df.to_excel(writer, sheet_name="All Listings", index=False)

        tier1_df.to_excel(writer, sheet_name="Tier1 Only", index=False)

        new_df.to_excel(writer, sheet_name="New Since Yesterday", index=False)

    send_email(jobs)

    send_tier1_alert(tier1_df.to_dict("records"))

if __name__ == "__main__":
    main()
