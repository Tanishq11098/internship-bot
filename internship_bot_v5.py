import smtplib, os, time, random, json
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests
from bs4 import BeautifulSoup

# ── Google Sheets (gspread) ──────────────────
import gspread
from google.oauth2.service_account import Credentials

# ── Anthropic Claude API ─────────────────────
import anthropic

# ─────────────────────────────────────────────
# ENV / CONFIG
# ─────────────────────────────────────────────
YOUR_EMAIL    = os.environ["EMAIL"]
YOUR_PASSWORD = os.environ["PASSWORD"]
SEND_TO_EMAIL = os.environ["EMAIL"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]

# Google Sheets: service account JSON stored as GitHub Secret
# Set secret name: GOOGLE_SERVICE_ACCOUNT_JSON  (full JSON string)
# Set secret name: GSHEET_ID  (the spreadsheet ID from the URL)
GSHEET_ID     = os.environ["GSHEET_ID"]
GSHEET_TAB    = "seen_listings"   # tab name for deduplication log

# ─────────────────────────────────────────────
# SCHEDULE: Mon, Thu, Sat, Sun
# ─────────────────────────────────────────────
RUN_ON_WEEKDAYS = {0, 1, 2, 3, 4, 5, 6}

# ─────────────────────────────────────────────
# TANISHQ'S RESUME (for AI fit scoring)
# ─────────────────────────────────────────────
RESUME_TEXT = """
Name: Tanishq Singhal
Degree: BBA Finance & Banking — IMS UCC Ghaziabad (2024–2027)
Certification: SEBI NISM Investor Certification (valid till March 2027)

Experience:
- Founder Fellow @ 23 Ventures (Dec 2025–Feb 2026): Secured startup credits (Cloudflare, OpenAI,
  Mixpanel), hosted offline hackathon (14 teams, 30+ participants), partnered with 25+ colleges
  across Delhi NCR and Pune for hackathons and webinars.
- Finance Intern @ Yhills, Noida (Jun–Aug 2025): Built financial models enabling 30% faster
  evaluations, prepared 5+ financial reports with company profiles and forecasts, conducted
  revenue/cost trend analysis using Advanced Excel and Google Sheets.
- Campus Ambassador @ Yhills (May–Jun 2025): Generated 40+ registrations (100%+ growth),
  coordinated 5+ events, managed 40+ member student community.

Skills:
- Technical: Advanced Excel (VLOOKUP, Index-Match, Pivot), Power BI, Google Sheets
- Finance: Financial Modeling, Ratio Analysis, EBITDA, Forecasting, Sales Growth
- Soft: Leadership, Event Management, Stakeholder Communication, Problem-Solving

Achievements:
- Winter Consulting 2025 – IIT Guwahati (Top 10%: Consulting & Strategy)
- McKinsey Forward Program (Dec 2025: Problem-Solving & Strategic Thinking)

Best-fit domains: Founder's Office, Strategy & Consulting, Investment Banking,
Equity Research, Valuation, FP&A, Operations / Growth, Management Trainee
"""

HEADERS_LIST = [
    {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"},
    {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/119.0.0.0 Safari/537.36"},
    {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/118.0.0.0 Safari/537.36"},
]

DOMAIN_COLORS = {
    "Investment Banking":   "C00000",
    "Equity Research":      "E26B0A",
    "Wealth Management":    "375623",
    "Private Equity":       "7030A0",
    "Venture Capital":      "7030A0",
    "Hedge Fund":           "7030A0",
    "Credit / Debt":        "1F4E79",
    "Treasury":             "1F4E79",
    "Compliance / Risk":    "833C00",
    "Audit / Assurance":    "833C00",
    "Valuation":            "E26B0A",
    "FP&A":                 "375623",
    "Fintech / Payments":   "0070C0",
    "Finance Operations":   "404040",
    "Accounting / Tax":     "404040",
    "Asset Management":     "375623",
    "Portfolio Management": "375623",
    "KPO Finance":          "1F4E79",
    "Founder's Office":     "A50021",
    "Management Trainee":   "BF8F00",
    "Strategy & Consulting":"1D6A96",
    "Operations / Growth":  "4A7C59",
}

def get_headers():
    return random.choice(HEADERS_LIST)

def detect_domain(title, company=""):
    text = (title + " " + company).lower()
    if any(x in text for x in ["founder", "chief of staff", "cxo office", "ceo office", "founders office"]):
        return "Founder's Office"
    if any(x in text for x in ["management trainee", "mt program", "graduate trainee", "management intern"]):
        return "Management Trainee"
    if any(x in text for x in ["strategy", "consulting intern", "management consulting", "business consulting", "mckinsey", "bcg", "bain", "deloitte", "kpmg consulting"]):
        return "Strategy & Consulting"
    if any(x in text for x in ["operations intern", "growth intern", "growth hacking", "business operations", "ops intern", "revenue operations"]):
        return "Operations / Growth"
    if any(x in text for x in ["investment bank", "ib ", "m&a", "capital market", "corporate finance"]): return "Investment Banking"
    if any(x in text for x in ["equity research", "equity analyst", "stock analyst", "market research", "research analyst"]): return "Equity Research"
    if any(x in text for x in ["wealth", "private banking", "relationship manager", "hni"]): return "Wealth Management"
    if any(x in text for x in ["private equity", " pe ", "buyout", "growth equity"]): return "Private Equity"
    if any(x in text for x in ["venture", " vc ", "startup invest"]): return "Venture Capital"
    if any(x in text for x in ["hedge", "quant", "algo", "derivatives"]): return "Hedge Fund"
    if any(x in text for x in ["credit", "debt", "lending", "fixed income", "bond", "nbfc"]): return "Credit / Debt"
    if any(x in text for x in ["treasury", "forex", " fx ", "currency", "liquidity"]): return "Treasury"
    if any(x in text for x in ["compliance", " risk", "regulatory", "kyc", "aml"]): return "Compliance / Risk"
    if any(x in text for x in ["audit", "assurance", "internal audit"]): return "Audit / Assurance"
    if any(x in text for x in ["valuat", "dcf", "modell", "financial model"]): return "Valuation"
    if any(x in text for x in ["fp&a", "financial planning", "budgeting", "forecasting"]): return "FP&A"
    if any(x in text for x in ["fintech", "payment", "wallet", "insurtech", "neobank"]): return "Fintech / Payments"
    if any(x in text for x in ["kpo", "research process"]): return "KPO Finance"
    if any(x in text for x in ["portfolio", "asset manag", "fund manag", "aum", "mutual fund"]): return "Asset Management"
    if any(x in text for x in ["account", " tax", "gst", "tally", "bookkeep"]): return "Accounting / Tax"
    return "Finance Operations"

# ═════════════════════════════════════════════
# GOOGLE SHEETS — DEDUPLICATION ACROSS RUNS
# ═════════════════════════════════════════════

def get_gsheet_client():
    """Authenticate with Google Sheets using service account JSON from env."""
    sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    sa_info = json.loads(sa_json)
    scopes  = ["https://www.googleapis.com/auth/spreadsheets"]
    creds   = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

def load_seen_keys(client):
    """Load all previously seen (title, company) keys from Google Sheet."""
    try:
        sh  = client.open_by_key(GSHEET_ID)
        try:
            ws = sh.worksheet(GSHEET_TAB)
        except gspread.exceptions.WorksheetNotFound:
            # First run — create the tab with headers
            ws = sh.add_worksheet(title=GSHEET_TAB, rows=5000, cols=4)
            ws.append_row(["title_key", "company_key", "domain", "first_seen"])
        rows = ws.get_all_values()
        seen = set()
        for row in rows[1:]:   # skip header
            if len(row) >= 2:
                seen.add((row[0], row[1]))
        print(f"Loaded {len(seen)} previously seen listings from Google Sheets.")
        return seen, ws
    except Exception as e:
        print(f"Google Sheets load error: {e}")
        return set(), None

def save_new_keys(ws, new_jobs):
    """Append newly seen jobs to the Google Sheet dedup log."""
    if not ws or not new_jobs:
        return
    try:
        today_str = date.today().isoformat()
        rows = [
            [j["title"].strip().lower()[:40],
             j["company"].strip().lower()[:30],
             j.get("domain", ""),
             today_str]
            for j in new_jobs
        ]
        ws.append_rows(rows, value_input_option="RAW")
        print(f"Saved {len(rows)} new listings to Google Sheets dedup log.")
    except Exception as e:
        print(f"Google Sheets save error: {e}")

def filter_new_only(all_jobs, seen_keys):
    """Return only jobs not seen in previous runs."""
    new_jobs, dup_count = [], 0
    local_seen = set()
    for j in all_jobs:
        key = (j["title"].strip().lower()[:40], j["company"].strip().lower()[:30])
        if key in seen_keys or key in local_seen:
            dup_count += 1
        else:
            local_seen.add(key)
            new_jobs.append(j)
    print(f"Dedup: {dup_count} duplicates removed, {len(new_jobs)} new listings kept.")
    return new_jobs

# ═════════════════════════════════════════════
# AI FIT SCORING — CLAUDE API
# ═════════════════════════════════════════════

def score_jobs_with_ai(jobs):
    """
    Batch score all jobs using Claude.
    Returns list of jobs with 'fit_score' (1-10) and 'fit_reason' (short string) added.
    To save API calls, we send all jobs in one prompt and parse JSON back.
    """
    if not jobs:
        return jobs

    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

    # Build a compact job list for the prompt
    job_summaries = []
    for i, j in enumerate(jobs):
        job_summaries.append(f"{i}: {j['title']} @ {j['company']} | Domain: {j['domain']} | Location: {j['location']} | Stipend: {j['stipend']}")
    job_text = "\n".join(job_summaries)

    prompt = f"""You are a career advisor. Score each internship listing for fit with the candidate below.

CANDIDATE RESUME:
{RESUME_TEXT}

INTERNSHIP LISTINGS (index: title @ company | domain | location | stipend):
{job_text}

For each listing, return a JSON array (same order, same indexes) like:
[
  {{"index": 0, "fit_score": 8, "fit_reason": "Matches finance modeling skills and SEBI cert"}},
  {{"index": 1, "fit_score": 5, "fit_reason": "Partial match, accounting-heavy role"}},
  ...
]

Scoring guide:
- 9-10: Perfect match (domain aligns, uses candidate's exact skills, good for BBA 1st year)
- 7-8:  Strong match (related domain, transferable skills)
- 5-6:  Moderate match (some overlap but gaps exist)
- 3-4:  Weak match (different domain, limited overlap)
- 1-2:  Poor match (irrelevant to profile)

Return ONLY the JSON array, no explanation, no markdown."""

    try:
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = response.content[0].text.strip()
        # Strip markdown fences if present
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        scores = json.loads(raw.strip())
        score_map = {item["index"]: item for item in scores}
        for i, job in enumerate(jobs):
            info = score_map.get(i, {})
            job["fit_score"]  = info.get("fit_score", 5)
            job["fit_reason"] = info.get("fit_reason", "Not scored")
        print(f"AI fit scoring complete for {len(jobs)} listings.")
    except Exception as e:
        print(f"AI scoring error: {e}")
        for job in jobs:
            job.setdefault("fit_score", 5)
            job.setdefault("fit_reason", "Scoring unavailable")

    return jobs

# ═════════════════════════════════════════════
# SOURCE 1: INTERNSHALA
# ═════════════════════════════════════════════
def scrape_internshala():
    results = []
    urls = [
        ("https://internshala.com/internships/finance-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/finance-internship-in-gurgaon/", "Gurugram"),
        ("https://internshala.com/internships/finance-internship-in-noida/", "Noida"),
        ("https://internshala.com/internships/investment-banking-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/equity-research-internship/", "Delhi NCR"),
        ("https://internshala.com/internships/audit-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/mba-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/accounting-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/strategy-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/operations-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/management-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/business-development-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/general-management-internship/", "Delhi NCR"),
        ("https://internshala.com/internships/consulting-internship-in-delhi/", "Delhi"),
    ]
    for url, loc in urls:
        try:
            time.sleep(random.uniform(1, 2))
            r = requests.get(url, headers=get_headers(), timeout=15)
            soup = BeautifulSoup(r.text, "html.parser")
            cards = soup.select(".internship_meta")[:20]
            for card in cards:
                title    = card.select_one(".profile")
                company  = card.select_one(".company_name")
                location = card.select_one(".location_link")
                stipend  = card.select_one(".stipend")
                duration = card.select_one(".item_body")
                deadline = card.select_one(".apply-by")
                link_tag = card.find_parent("a")
                if title:
                    t = title.get_text(strip=True)
                    c = company.get_text(strip=True) if company else "N/A"
                    results.append({
                        "title":     t, "company": c,
                        "firm_type": "Startup / Corporate",
                        "domain":    detect_domain(t, c),
                        "location":  location.get_text(strip=True) if location else loc,
                        "stipend":   stipend.get_text(strip=True) if stipend else "Not disclosed",
                        "duration":  duration.get_text(strip=True) if duration else "N/A",
                        "posted":    "< 7 days",
                        "deadline":  deadline.get_text(strip=True).replace("Apply By:", "").strip() if deadline else "Not mentioned",
                        "status":    "Not Applied", "platform": "Internshala",
                        "link":      "https://internshala.com" + link_tag["href"] if link_tag and link_tag.get("href") else url
                    })
        except Exception as e:
            print(f"Internshala error ({loc}): {e}")
    print(f"Internshala: {len(results)} results")
    return results

# ═════════════════════════════════════════════
# SOURCE 2: INDEED
# ═════════════════════════════════════════════
def scrape_indeed():
    results = []
    queries = [
        "finance+internship", "investment+banking+intern", "equity+research+intern",
        "audit+intern", "wealth+management+intern", "private+equity+intern",
        "fintech+intern", "valuation+intern", "credit+analyst+intern",
        "compliance+intern", "treasury+intern", "financial+analyst+intern",
        "accounting+intern", "KPO+finance+intern", "founder+office+intern",
        "chief+of+staff+intern", "management+trainee+delhi", "strategy+intern",
        "consulting+intern", "operations+intern", "growth+intern",
    ]
    for q in queries:
        try:
            time.sleep(random.uniform(1, 2))
            url = f"https://in.indeed.com/jobs?q={q}&l=Delhi+NCR&fromage=7"
            r = requests.get(url, headers=get_headers(), timeout=15)
            soup = BeautifulSoup(r.text, "html.parser")
            for job in soup.select(".job_seen_beacon")[:8]:
                title    = job.select_one('[class*="jobTitle"]')
                company  = job.select_one('[class*="companyName"]')
                location = job.select_one('[class*="companyLocation"]')
                salary   = job.select_one('[class*="salary"]')
                link_tag = job.select_one("a[id]")
                jk = link_tag["id"].replace("job_", "") if link_tag and link_tag.get("id") else ""
                t = title.get_text(strip=True) if title else "Intern"
                c = company.get_text(strip=True) if company else "N/A"
                results.append({
                    "title": t, "company": c, "firm_type": "Corporate",
                    "domain": detect_domain(t, c),
                    "location": location.get_text(strip=True) if location else "Delhi NCR",
                    "stipend": salary.get_text(strip=True) if salary else "Not disclosed",
                    "duration": "3-6 Months", "posted": "< 7 days",
                    "deadline": "Not mentioned", "status": "Not Applied",
                    "platform": "Indeed",
                    "link": f"https://in.indeed.com/viewjob?jk={jk}" if jk else url
                })
        except Exception as e:
            print(f"Indeed error ({q}): {e}")
    print(f"Indeed: {len(results)} results")
    return results

# ═════════════════════════════════════════════
# SOURCE 3: GLASSDOOR
# ═════════════════════════════════════════════
def scrape_glassdoor():
    results = []
    searches = [
        ("finance-internship-delhi", "Finance Operations"),
        ("investment-banking-intern-delhi", "Investment Banking"),
        ("equity-research-intern-gurgaon", "Equity Research"),
        ("audit-intern-delhi", "Audit / Assurance"),
        ("private-equity-intern-delhi", "Private Equity"),
        ("fintech-intern-noida", "Fintech / Payments"),
        ("wealth-management-intern-delhi", "Wealth Management"),
        ("credit-analyst-intern-gurugram", "Credit / Debt"),
        ("founders-office-intern-delhi", "Founder's Office"),
        ("management-trainee-delhi", "Management Trainee"),
        ("strategy-intern-delhi", "Strategy & Consulting"),
        ("consulting-intern-delhi", "Strategy & Consulting"),
        ("operations-intern-delhi", "Operations / Growth"),
    ]
    for slug, fallback in searches:
        try:
            time.sleep(random.uniform(1, 2))
            url = f"https://www.glassdoor.co.in/Job/{slug}-jobs-SRCH_KO0,{len(slug)}.htm"
            r = requests.get(url, headers=get_headers(), timeout=15)
            soup = BeautifulSoup(r.text, "html.parser")
            for job in soup.select("[data-test='jobListing']")[:8]:
                title   = job.select_one("[data-test='job-title']") or job.select_one(".job-title")
                company = job.select_one("[data-test='employer-name']") or job.select_one(".employer-name")
                loc     = job.select_one("[data-test='emp-location']") or job.select_one(".location")
                link    = job.select_one("a")
                t = title.get_text(strip=True) if title else f"{fallback} Intern"
                c = company.get_text(strip=True) if company else "N/A"
                href = ("https://www.glassdoor.co.in" + link["href"] if link and link.get("href") and link["href"].startswith("/") else (link["href"] if link and link.get("href") else url))
                d = detect_domain(t, c)
                results.append({
                    "title": t, "company": c, "firm_type": "Corporate / MNC",
                    "domain": d if d != "Finance Operations" else fallback,
                    "location": loc.get_text(strip=True) if loc else "Delhi NCR",
                    "stipend": "Not disclosed", "duration": "3-6 Months",
                    "posted": "< 7 days", "deadline": "Not mentioned",
                    "status": "Not Applied", "platform": "Glassdoor", "link": href
                })
        except Exception as e:
            print(f"Glassdoor error ({slug}): {e}")
    print(f"Glassdoor: {len(results)} results")
    return results

# ═════════════════════════════════════════════
# SOURCE 4: NAUKRI
# ═════════════════════════════════════════════
def scrape_naukri():
    results = []
    searches = [
        ("finance-internship-jobs-in-delhi-ncr", "Finance Operations"),
        ("investment-banking-internship-jobs-in-delhi-ncr", "Investment Banking"),
        ("equity-research-internship-jobs-in-delhi-ncr", "Equity Research"),
        ("audit-internship-jobs-in-delhi-ncr", "Audit / Assurance"),
        ("wealth-management-internship-jobs-in-delhi-ncr", "Wealth Management"),
        ("credit-analyst-internship-jobs-in-delhi-ncr", "Credit / Debt"),
        ("financial-analyst-internship-jobs-in-delhi-ncr", "Finance Operations"),
        ("compliance-internship-jobs-in-delhi-ncr", "Compliance / Risk"),
        ("fintech-internship-jobs-in-delhi-ncr", "Fintech / Payments"),
        ("treasury-internship-jobs-in-delhi-ncr", "Treasury"),
        ("private-equity-internship-jobs-in-delhi-ncr", "Private Equity"),
        ("valuation-internship-jobs-in-delhi-ncr", "Valuation"),
        ("founders-office-internship-jobs-in-delhi-ncr", "Founder's Office"),
        ("management-trainee-jobs-in-delhi-ncr", "Management Trainee"),
        ("strategy-internship-jobs-in-delhi-ncr", "Strategy & Consulting"),
        ("consulting-internship-jobs-in-delhi-ncr", "Strategy & Consulting"),
        ("operations-internship-jobs-in-delhi-ncr", "Operations / Growth"),
        ("growth-internship-jobs-in-delhi-ncr", "Operations / Growth"),
    ]
    for slug, fallback in searches:
        try:
            time.sleep(random.uniform(1.5, 2.5))
            url = f"https://www.naukri.com/{slug}"
            r = requests.get(url, headers=get_headers(), timeout=15)
            soup = BeautifulSoup(r.text, "html.parser")
            jobs = soup.select(".jobTuple") or soup.select(".job-container") or soup.select("article.jobTupleHeader")
            for job in jobs[:8]:
                title   = job.select_one(".title") or job.select_one(".jobTitle") or job.select_one("a.title")
                company = job.select_one(".companyInfo span") or job.select_one(".company-name") or job.select_one(".subTitle")
                loc     = job.select_one(".location") or job.select_one(".locWdth")
                link    = job.select_one("a.title") or job.select_one("a[href*='naukri']")
                t = title.get_text(strip=True) if title else f"{fallback} Intern"
                c = company.get_text(strip=True) if company else "N/A"
                href = link["href"] if link and link.get("href") else url
                d = detect_domain(t, c)
                results.append({
                    "title": t, "company": c, "firm_type": "Corporate / MNC",
                    "domain": d if d != "Finance Operations" else fallback,
                    "location": loc.get_text(strip=True) if loc else "Delhi NCR",
                    "stipend": "Not disclosed", "duration": "3-6 Months",
                    "posted": "< 7 days", "deadline": "Not mentioned",
                    "status": "Not Applied", "platform": "Naukri",
                    "link": href if href.startswith("http") else f"https://www.naukri.com/{slug}"
                })
        except Exception as e:
            print(f"Naukri error ({slug}): {e}")
    print(f"Naukri: {len(results)} results")
    return results

# ═════════════════════════════════════════════
# SOURCE 5: FOUNDIT
# ═════════════════════════════════════════════
def scrape_foundit():
    results = []
    searches = [
        ("finance", "Finance Operations"), ("investment-banking", "Investment Banking"),
        ("equity-research", "Equity Research"), ("audit", "Audit / Assurance"),
        ("wealth-management", "Wealth Management"), ("fintech", "Fintech / Payments"),
        ("credit-analyst", "Credit / Debt"), ("compliance", "Compliance / Risk"),
        ("founders-office", "Founder's Office"), ("management-trainee", "Management Trainee"),
        ("strategy-consulting", "Strategy & Consulting"), ("business-operations", "Operations / Growth"),
        ("growth", "Operations / Growth"),
    ]
    for keyword, fallback in searches:
        try:
            time.sleep(random.uniform(1.5, 2.5))
            url = f"https://www.foundit.in/search/{keyword}-internship-jobs-in-delhi?experienceRanges=0~0"
            r = requests.get(url, headers=get_headers(), timeout=15)
            soup = BeautifulSoup(r.text, "html.parser")
            jobs = soup.select(".jobCard") or soup.select(".card-apply-content") or soup.select(".job-container")
            for job in jobs[:8]:
                title   = job.select_one(".jobTitle") or job.select_one(".designation") or job.select_one("h3")
                company = job.select_one(".companyName") or job.select_one(".company") or job.select_one("h4")
                loc     = job.select_one(".location") or job.select_one(".loc")
                link    = job.select_one("a")
                t = title.get_text(strip=True) if title else f"{fallback} Intern"
                c = company.get_text(strip=True) if company else "N/A"
                href = ("https://www.foundit.in" + link["href"] if link and link.get("href") and link["href"].startswith("/") else (link["href"] if link and link.get("href") else url))
                d = detect_domain(t, c)
                results.append({
                    "title": t, "company": c, "firm_type": "Corporate",
                    "domain": d if d != "Finance Operations" else fallback,
                    "location": loc.get_text(strip=True) if loc else "Delhi NCR",
                    "stipend": "Not disclosed", "duration": "3-6 Months",
                    "posted": "< 7 days", "deadline": "Not mentioned",
                    "status": "Not Applied", "platform": "Foundit", "link": href
                })
        except Exception as e:
            print(f"Foundit error ({keyword}): {e}")
    print(f"Foundit: {len(results)} results")
    return results

# ═════════════════════════════════════════════
# SOURCE 6: JOBAAJ
# ═════════════════════════════════════════════
def scrape_jobaaj():
    results = []
    urls = [
        ("https://www.jobaaj.com/jobs/finance-internship-in-delhi", "Finance Operations"),
        ("https://www.jobaaj.com/jobs/investment-banking-internship-in-delhi", "Investment Banking"),
        ("https://www.jobaaj.com/jobs/equity-research-internship-in-delhi", "Equity Research"),
        ("https://www.jobaaj.com/jobs/audit-internship-in-delhi", "Audit / Assurance"),
        ("https://www.jobaaj.com/jobs/private-equity-internship-in-delhi", "Private Equity"),
        ("https://www.jobaaj.com/jobs/wealth-management-internship-in-delhi", "Wealth Management"),
        ("https://www.jobaaj.com/jobs/fintech-internship-in-delhi", "Fintech / Payments"),
        ("https://www.jobaaj.com/jobs/compliance-internship-in-delhi", "Compliance / Risk"),
        ("https://www.jobaaj.com/jobs/founders-office-internship-in-delhi", "Founder's Office"),
        ("https://www.jobaaj.com/jobs/management-trainee-in-delhi", "Management Trainee"),
        ("https://www.jobaaj.com/jobs/strategy-internship-in-delhi", "Strategy & Consulting"),
        ("https://www.jobaaj.com/jobs/operations-internship-in-delhi", "Operations / Growth"),
    ]
    for url, fallback in urls:
        try:
            time.sleep(random.uniform(1.5, 2.5))
            r = requests.get(url, headers=get_headers(), timeout=15)
            soup = BeautifulSoup(r.text, "html.parser")
            jobs = soup.select(".job-list-item") or soup.select(".job-card") or soup.select(".jobcard")
            for job in jobs[:8]:
                title   = job.select_one("h2") or job.select_one(".job-title") or job.select_one("h3")
                company = job.select_one(".company") or job.select_one(".employer") or job.select_one("h4")
                loc     = job.select_one(".location") or job.select_one(".city")
                link    = job.select_one("a")
                t = title.get_text(strip=True) if title else f"{fallback} Intern"
                c = company.get_text(strip=True) if company else "N/A"
                href = ("https://www.jobaaj.com" + link["href"] if link and link.get("href") and link["href"].startswith("/") else (link["href"] if link and link.get("href") else url))
                d = detect_domain(t, c)
                results.append({
                    "title": t, "company": c, "firm_type": "Corporate / MNC",
                    "domain": d if d != "Finance Operations" else fallback,
                    "location": loc.get_text(strip=True) if loc else "Delhi NCR",
                    "stipend": "Not disclosed", "duration": "3-6 Months",
                    "posted": "< 7 days", "deadline": "Not mentioned",
                    "status": "Not Applied", "platform": "Jobaaj",
                    "link": href if href.startswith("http") else url
                })
        except Exception as e:
            print(f"Jobaaj error ({url}): {e}")
    print(f"Jobaaj: {len(results)} results")
    return results

# ═════════════════════════════════════════════
# SOURCE 7: WELLFOUND (AngelList) — NEW
# ═════════════════════════════════════════════
def scrape_wellfound():
    """
    Scrapes Wellfound (formerly AngelList Talent) for startup internships.
    Best source for Founder's Office, Strategy, and early-stage roles.
    """
    results = []
    searches = [
        ("finance", "Finance Operations"),
        ("investment", "Investment Banking"),
        ("strategy", "Strategy & Consulting"),
        ("operations", "Operations / Growth"),
        ("founder", "Founder's Office"),
        ("growth", "Operations / Growth"),
        ("business-development", "Operations / Growth"),
        ("consulting", "Strategy & Consulting"),
        ("fintech", "Fintech / Payments"),
        ("venture-capital", "Venture Capital"),
    ]
    for keyword, fallback in searches:
        try:
            time.sleep(random.uniform(2, 3))
            url = f"https://wellfound.com/role/l/{keyword}-intern/india"
            r = requests.get(url, headers=get_headers(), timeout=20)
            soup = BeautifulSoup(r.text, "html.parser")

            # Wellfound job cards
            jobs = (soup.select("[data-test='StartupResult']") or
                    soup.select(".styles_component__Ey28k") or
                    soup.select("div[class*='JobListing']") or
                    soup.select("div[class*='job-listing']"))

            for job in jobs[:10]:
                title   = (job.select_one("a[class*='JobTitle']") or
                           job.select_one("[class*='title']") or
                           job.select_one("h2") or job.select_one("h3"))
                company = (job.select_one("[class*='startup-link']") or
                           job.select_one("[class*='company']") or
                           job.select_one("h4"))
                loc     = (job.select_one("[class*='location']") or
                           job.select_one("[class*='Location']"))
                comp    = (job.select_one("[class*='compensation']") or
                           job.select_one("[class*='salary']"))
                link    = job.select_one("a[href*='/jobs/']") or job.select_one("a")

                t = title.get_text(strip=True) if title else f"{fallback} Intern"
                c = company.get_text(strip=True) if company else "Startup"

                # Build absolute href
                if link and link.get("href"):
                    raw_href = link["href"]
                    href = ("https://wellfound.com" + raw_href
                            if raw_href.startswith("/") else raw_href)
                else:
                    href = url

                d = detect_domain(t, c)
                results.append({
                    "title":     t,
                    "company":   c,
                    "firm_type": "Startup (Wellfound)",
                    "domain":    d if d != "Finance Operations" else fallback,
                    "location":  loc.get_text(strip=True) if loc else "Delhi NCR / Remote",
                    "stipend":   comp.get_text(strip=True) if comp else "Not disclosed",
                    "duration":  "3-6 Months",
                    "posted":    "< 7 days",
                    "deadline":  "Not mentioned",
                    "status":    "Not Applied",
                    "platform":  "Wellfound",
                    "link":      href
                })
        except Exception as e:
            print(f"Wellfound error ({keyword}): {e}")

    print(f"Wellfound: {len(results)} results")
    return results

# ═════════════════════════════════════════════
# QUALITY SEED LISTINGS
# ═════════════════════════════════════════════
def get_quality_listings():
    return [
        {"title":"Finance Intern – PPO (Rs.25K/mo)","company":"TravClan Technology","firm_type":"Fintech Startup","domain":"Fintech / Payments","location":"Connaught Place, Delhi","stipend":"Rs.25,000/mo","duration":"6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Indeed","link":"https://in.indeed.com/viewjob?jk=44b725b689aa3123"},
        {"title":"Wealth Management Trainee","company":"Mint Global LLC","firm_type":"Wealth / Investment Firm","domain":"Wealth Management","location":"Delhi / Gurugram / Noida","stipend":"Rs.15,000/mo + Incentives","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Internshala","link":"https://internshala.com/internships/finance-internship-in-delhi/"},
        {"title":"Financial Audit Intern","company":"PwC India","firm_type":"Big 4 / MNC","domain":"Audit / Assurance","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Jobaaj","link":"https://www.jobaaj.com/job/pricewaterhousecoopers-pwc-financial-audit-intern-delhi-ncr-0-to-1-years-819084"},
        {"title":"Investment Banking Intern","company":"HSBC India","firm_type":"MNC Investment Bank","domain":"Investment Banking","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Jobaaj","link":"https://www.jobaaj.com/job/hsbc-investment-banking-intern-delhi-ncr-0-to-1-years-670612"},
        {"title":"Corporate Finance Intern","company":"Barclays","firm_type":"MNC Investment Bank","domain":"Investment Banking","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Jobaaj","link":"https://www.jobaaj.com/job/barclays-corporate-finance-intern-delhi-ncr-0-to-1-years-760847"},
        {"title":"Equity Research Intern","company":"Trade Brains","firm_type":"Fintech / Research","domain":"Equity Research","location":"Delhi NCR / Remote","stipend":"Rs.10,000-15,000/mo","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Internshala","link":"https://internshala.com/internships/finance-internship/"},
        {"title":"Founder's Office Intern","company":"Undisclosed D2C Startup","firm_type":"Early-Stage Startup","domain":"Founder's Office","location":"Delhi / Gurugram","stipend":"Rs.10,000-20,000/mo","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Wellfound","link":"https://wellfound.com/role/l/founder-intern/india"},
        {"title":"Chief of Staff Intern","company":"Undisclosed SaaS Startup","firm_type":"Tech Startup","domain":"Founder's Office","location":"Gurugram, Haryana","stipend":"Rs.15,000-25,000/mo","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Wellfound","link":"https://wellfound.com/role/l/chief-of-staff/india"},
        {"title":"Management Trainee – Finance & Strategy","company":"Undisclosed Conglomerate","firm_type":"Large Corporate","domain":"Management Trainee","location":"Delhi NCR","stipend":"As per norms","duration":"12 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Naukri","link":"https://www.naukri.com/management-trainee-jobs-in-delhi-ncr"},
        {"title":"Strategy & Consulting Intern","company":"Undisclosed Boutique","firm_type":"Management Consulting","domain":"Strategy & Consulting","location":"Delhi / Gurugram","stipend":"Rs.15,000-25,000/mo","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Internshala","link":"https://internshala.com/internships/consulting-internship-in-delhi/"},
        {"title":"Operations & Growth Intern","company":"Undisclosed Startup","firm_type":"Growth-Stage Startup","domain":"Operations / Growth","location":"Delhi / Noida","stipend":"Rs.10,000-15,000/mo","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Wellfound","link":"https://wellfound.com/role/l/operations-intern/india"},
        {"title":"Venture Capital Analyst Intern","company":"Undisclosed VC Firm","firm_type":"Venture Capital","domain":"Venture Capital","location":"Delhi / Gurugram","stipend":"Not disclosed","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Wellfound","link":"https://wellfound.com/role/l/venture-capital-intern/india"},
        {"title":"Credit Analyst Intern","company":"Undisclosed NBFC","firm_type":"NBFC / Lending","domain":"Credit / Debt","location":"Gurugram","stipend":"Rs.10,000-15,000/mo","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Naukri","link":"https://www.naukri.com/credit-analyst-internship-jobs-in-delhi-ncr"},
        {"title":"Valuation Analyst Intern","company":"Undisclosed Boutique","firm_type":"Boutique Advisory","domain":"Valuation","location":"Delhi / Gurugram","stipend":"Rs.10,000-15,000/mo","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Internshala","link":"https://internshala.com/internships/finance-internship-in-delhi/"},
        {"title":"KPO Financial Analyst Intern","company":"Undisclosed KPO","firm_type":"KPO Finance","domain":"KPO Finance","location":"Noida, UP","stipend":"Rs.10,000-12,000/mo","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Naukri","link":"https://www.naukri.com/kpo-internship-jobs-in-noida"},
    ]

# ═════════════════════════════════════════════
# BUILD EXCEL  (now with Fit Score column)
# ═════════════════════════════════════════════
def build_excel(internships, filepath):
    wb = Workbook()
    ws = wb.active
    ws.title = "All Internships"

    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    data_font   = Font(name="Arial", size=9)
    link_font   = Font(name="Arial", size=9, color="1558BB", underline="single")
    thin = Border(
        left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),  bottom=Side(style="thin", color="BFBFBF")
    )
    today = date.today().strftime("%B %d, %Y")

    # Title row
    ws.merge_cells("A1:P1")
    ws["A1"] = f"Finance & Management Internships — Delhi NCR | {today} | Sources: Internshala + Indeed + Glassdoor + Naukri + Foundit + Jobaaj + Wellfound"
    ws["A1"].font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = PatternFill("solid", start_color="D6E4F0")
    ws.row_dimensions[1].height = 30

    # Legend row
    ws.merge_cells("A2:P2")
    ws["A2"] = ("Color legend — Red: IB | Orange: Equity/Valuation | Purple: PE/VC/HF | "
                "Green: Wealth/AM/FP&A | Blue: Fintech/Credit/KPO | Brown: Audit/Compliance | "
                "Crimson: Founder's Office | Gold: Mgmt Trainee | Steel Blue: Strategy | Sage: Ops/Growth  "
                "★ AI Fit Score: 9-10 = Perfect | 7-8 = Strong | 5-6 = Moderate | 1-4 = Weak")
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="7F6000")
    ws["A2"].fill = PatternFill("solid", start_color="FFF2CC")
    ws["A2"].alignment = Alignment(horizontal="left", wrap_text=True)
    ws.row_dimensions[2].height = 30

    headers = ["#", "Role / Internship Title", "Company", "Firm Type", "Domain", "Location",
               "Stipend", "Duration", "Posted", "Deadline", "Platform", "Apply Link",
               "Status", "AI Fit Score", "Why It Fits (AI)", "Notes"]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin
    ws.row_dimensions[3].height = 30

    # Sort by fit_score descending, then domain
    internships_sorted = sorted(internships, key=lambda x: (-x.get("fit_score", 5), x.get("domain", "Z")))

    for i, job in enumerate(internships_sorted, 1):
        r       = i + 3
        domain  = job.get("domain", "Finance Operations")
        color   = DOMAIN_COLORS.get(domain, "404040")
        row_bg  = "F5F5F5" if i % 2 == 0 else "FFFFFF"
        score   = job.get("fit_score", 5)

        # Score badge color
        if score >= 9:   score_bg = "C6EFCE"; score_fc = "276221"   # green
        elif score >= 7: score_bg = "FFEB9C"; score_fc = "9C5700"   # yellow
        elif score >= 5: score_bg = "DDEBF7"; score_fc = "1F4E79"   # blue
        else:            score_bg = "FFE0E0"; score_fc = "9C0006"   # red

        values = [i, job["title"], job["company"], job["firm_type"], job["domain"],
                  job["location"], job["stipend"], job["duration"], job["posted"],
                  job["deadline"], job["platform"], job["link"],
                  job["status"], f"{score}/10", job.get("fit_reason", ""), ""]

        for col, val in enumerate(values, 1):
            cell = ws.cell(r, col, val)
            cell.border = thin
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            if col == 2:    # Title
                cell.font = Font(name="Arial", size=9, bold=True, color=color)
                cell.fill = PatternFill("solid", start_color=row_bg)
            elif col == 5:  # Domain
                cell.font = Font(name="Arial", size=9, color=color)
                cell.fill = PatternFill("solid", start_color=row_bg)
            elif col == 12: # Link
                cell.font = link_font
                cell.fill = PatternFill("solid", start_color=row_bg)
                cell.hyperlink = str(val)
            elif col == 13: # Status
                cell.font = Font(name="Arial", size=9, bold=True, color="375623")
                cell.fill = PatternFill("solid", start_color="E2EFDA")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 14: # AI Fit Score
                cell.font = Font(name="Arial", size=10, bold=True, color=score_fc)
                cell.fill = PatternFill("solid", start_color=score_bg)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 15: # Why It Fits
                cell.font = Font(name="Arial", size=8, italic=True, color="404040")
                cell.fill = PatternFill("solid", start_color=row_bg)
            else:
                cell.font = data_font
                cell.fill = PatternFill("solid", start_color=row_bg)
        ws.row_dimensions[r].height = 40

    col_widths = [4, 38, 26, 24, 24, 24, 20, 14, 12, 18, 14, 50, 16, 12, 38, 20]
    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:P{3 + len(internships)}"

    # ── Sheet 2: Top Picks (fit score >= 7) ──────────────────────────
    ws_top = wb.create_sheet("🌟 Top Picks")
    ws_top.merge_cells("A1:P1")
    ws_top["A1"] = f"Top Picks for Tanishq — AI Fit Score ≥ 7 — {today}"
    ws_top["A1"].font = Font(name="Arial", bold=True, size=12, color="276221")
    ws_top["A1"].fill = PatternFill("solid", start_color="C6EFCE")
    ws_top["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_top.row_dimensions[1].height = 28

    for col, h in enumerate(headers, 1):
        cell = ws_top.cell(2, col, h)
        cell.font = header_font
        cell.fill = PatternFill("solid", start_color="276221")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin
    ws_top.row_dimensions[2].height = 28

    top_jobs = [j for j in internships_sorted if j.get("fit_score", 0) >= 7]
    for i, job in enumerate(top_jobs, 1):
        r      = i + 2
        domain = job.get("domain", "Finance Operations")
        color  = DOMAIN_COLORS.get(domain, "404040")
        score  = job.get("fit_score", 7)
        row_bg = "F0FFF0" if i % 2 == 0 else "FFFFFF"
        if score >= 9:   score_bg = "C6EFCE"; score_fc = "276221"
        elif score >= 7: score_bg = "FFEB9C"; score_fc = "9C5700"
        else:            score_bg = "DDEBF7"; score_fc = "1F4E79"

        values = [i, job["title"], job["company"], job["firm_type"], job["domain"],
                  job["location"], job["stipend"], job["duration"], job["posted"],
                  job["deadline"], job["platform"], job["link"],
                  job["status"], f"{score}/10", job.get("fit_reason", ""), ""]
        for col, val in enumerate(values, 1):
            cell = ws_top.cell(r, col, val)
            cell.border = thin
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if col == 2:
                cell.font = Font(name="Arial", size=9, bold=True, color=color)
                cell.fill = PatternFill("solid", start_color=row_bg)
            elif col == 12:
                cell.font = link_font
                cell.fill = PatternFill("solid", start_color=row_bg)
                cell.hyperlink = str(val)
            elif col == 14:
                cell.font = Font(name="Arial", size=10, bold=True, color=score_fc)
                cell.fill = PatternFill("solid", start_color=score_bg)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 15:
                cell.font = Font(name="Arial", size=8, italic=True, color="404040")
                cell.fill = PatternFill("solid", start_color=row_bg)
            else:
                cell.font = data_font
                cell.fill = PatternFill("solid", start_color=row_bg)
        ws_top.row_dimensions[r].height = 40

    for col, w in enumerate(col_widths, 1):
        ws_top.column_dimensions[get_column_letter(col)].width = w
    ws_top.freeze_panes = "A3"


    # ── Sheet 3: 3-Month Internships ─────────────────────────────────
    def is_3_month(d_str):
        import re
        d = d_str.lower().strip()
        if any(x in d for x in ["not mentioned", "not disclosed", "n/a", "flexible"]):
            return False
        nums = list(map(int, re.findall(r"\d+", d)))
        if not nums:
            return False
        if "month" in d:
            return min(nums) <= 3 and max(nums) <= 3
        if "week" in d:
            return min(nums) <= 12 and max(nums) <= 12
        return False

    def is_4_month(d_str):
        import re
        d = d_str.lower().strip()
        if any(x in d for x in ["not mentioned", "not disclosed", "n/a", "flexible"]):
            return False
        nums = list(map(int, re.findall(r"\d+", d)))
        if not nums:
            return False
        if "month" in d:
            return 4 in nums or max(nums) == 4
        if "week" in d:
            return min(nums) >= 13 and max(nums) <= 16
        return False

    for sheet_label, sheet_color, header_color, filter_fn, sheet_title in [
        ("3-Month Internships", "FFF2CC", "BF8F00",
         is_3_month, f"3-Month Internships — {today}"),
        ("4-Month Internships", "DDEBF7", "1D6A96",
         is_4_month, f"4-Month Internships — {today}"),
    ]:
        ws_dur = wb.create_sheet(sheet_label)
        ws_dur.merge_cells("A1:P1")
        ws_dur["A1"] = sheet_title
        ws_dur["A1"].font = Font(name="Arial", bold=True, size=12, color=header_color)
        ws_dur["A1"].fill = PatternFill("solid", start_color=sheet_color)
        ws_dur["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_dur.row_dimensions[1].height = 28

        for col, h in enumerate(headers, 1):
            cell = ws_dur.cell(2, col, h)
            cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill("solid", start_color=header_color)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin
        ws_dur.row_dimensions[2].height = 28

        filtered = [j for j in internships_sorted if filter_fn(j.get("duration", ""))]

        if not filtered:
            ws_dur.merge_cells("A3:P3")
            ws_dur["A3"] = "No listings found for this duration in today's run."
            ws_dur["A3"].font = Font(name="Arial", italic=True, size=10, color="7F7F7F")
            ws_dur["A3"].alignment = Alignment(horizontal="center", vertical="center")
            ws_dur.row_dimensions[3].height = 30
        else:
            for i, job in enumerate(filtered, 1):
                r      = i + 2
                domain = job.get("domain", "Finance Operations")
                color  = DOMAIN_COLORS.get(domain, "404040")
                row_bg = "FFFEF0" if sheet_label == "3-Month Internships" else "F0F6FF"
                row_bg = row_bg if i % 2 != 0 else "FFFFFF"
                score  = job.get("fit_score", 5)
                if score >= 9:   score_bg = "C6EFCE"; score_fc = "276221"
                elif score >= 7: score_bg = "FFEB9C"; score_fc = "9C5700"
                elif score >= 5: score_bg = "DDEBF7"; score_fc = "1F4E79"
                else:            score_bg = "FFE0E0"; score_fc = "9C0006"

                values = [i, job["title"], job["company"], job["firm_type"], job["domain"],
                          job["location"], job["stipend"], job["duration"], job["posted"],
                          job["deadline"], job["platform"], job["link"],
                          job["status"], f"{score}/10", job.get("fit_reason", ""), ""]

                for col, val in enumerate(values, 1):
                    cell = ws_dur.cell(r, col, val)
                    cell.border = thin
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    if col == 2:
                        cell.font = Font(name="Arial", size=9, bold=True, color=color)
                        cell.fill = PatternFill("solid", start_color=row_bg)
                    elif col == 12:
                        cell.font = Font(name="Arial", size=9, color="1558BB", underline="single")
                        cell.fill = PatternFill("solid", start_color=row_bg)
                        cell.hyperlink = str(val)
                    elif col == 13:
                        cell.font = Font(name="Arial", size=9, bold=True, color="375623")
                        cell.fill = PatternFill("solid", start_color="E2EFDA")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col == 14:
                        cell.font = Font(name="Arial", size=10, bold=True, color=score_fc)
                        cell.fill = PatternFill("solid", start_color=score_bg)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col == 15:
                        cell.font = Font(name="Arial", size=8, italic=True, color="404040")
                        cell.fill = PatternFill("solid", start_color=row_bg)
                    else:
                        cell.font = Font(name="Arial", size=9)
                        cell.fill = PatternFill("solid", start_color=row_bg)
                ws_dur.row_dimensions[r].height = 40

        for col, w in enumerate(col_widths, 1):
            ws_dur.column_dimensions[get_column_letter(col)].width = w
        ws_dur.freeze_panes = "A3"
        ws_dur.auto_filter.ref = f"A2:P{2 + max(len(filtered), 1)}"

    # ── Sheet 5: Domain Summary ───────────────────────────────────────
    ws2 = wb.create_sheet("Domain Summary")
    ws2.merge_cells("A1:C1")
    ws2["A1"] = f"Internship Count by Domain — {today}"
    ws2["A1"].font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    ws2["A1"].fill = PatternFill("solid", start_color="D6E4F0")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28
    for col, h in enumerate(["Domain", "Count", "Source Platforms"], 1):
        cell = ws2.cell(2, col, h)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center"); cell.border = thin

    domain_counts, domain_platforms = {}, {}
    for job in internships:
        d = job.get("domain", "Finance Operations")
        p = job.get("platform", "")
        domain_counts[d] = domain_counts.get(d, 0) + 1
        domain_platforms.setdefault(d, set()).add(p)

    for row_idx, (domain, count) in enumerate(sorted(domain_counts.items(), key=lambda x: -x[1]), 3):
        color = DOMAIN_COLORS.get(domain, "404040")
        bg = "F5F5F5" if row_idx % 2 == 0 else "FFFFFF"
        c1 = ws2.cell(row_idx, 1, domain);  c1.font = Font(name="Arial", size=10, bold=True, color=color)
        c2 = ws2.cell(row_idx, 2, count);   c2.font = Font(name="Arial", size=10, bold=True); c2.alignment = Alignment(horizontal="center")
        c3 = ws2.cell(row_idx, 3, ", ".join(sorted(domain_platforms.get(domain, [])))); c3.font = Font(name="Arial", size=9)
        for c in [c1, c2, c3]:
            c.fill = PatternFill("solid", start_color=bg); c.border = thin

    last = len(domain_counts) + 3
    for col, val in enumerate(["TOTAL", f"=SUM(B3:B{last-1})", ""], 1):
        cell = ws2.cell(last, col, val)
        cell.font = Font(name="Arial", bold=True, size=10)
        cell.fill = PatternFill("solid", start_color="D6E4F0"); cell.border = thin
        if col == 2: cell.alignment = Alignment(horizontal="center")
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 40

    # ── Sheet 6: Application Tracker ─────────────────────────────────
    ws3 = wb.create_sheet("Application Tracker")
    ws3.merge_cells("A1:H1")
    ws3["A1"] = f"My Application Tracker — {today}"
    ws3["A1"].font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    ws3["A1"].fill = PatternFill("solid", start_color="D6E4F0")
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28
    for col, h in enumerate(["#", "Company", "Role", "Domain", "Applied Date", "Status", "Interview Date", "Notes"], 1):
        cell = ws3.cell(2, col, h)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center"); cell.border = thin
    ws3.row_dimensions[2].height = 28
    status_map = [
        ("Not Applied","F2F2F2"),("Applied","BDD7EE"),("In Review","FFF2CC"),
        ("Interview Scheduled","E2EFDA"),("Offer Received","C6EFCE"),("Rejected","FFE0E0"),
    ]
    for i in range(3, 53):
        row_bg = "FFFFFF" if i % 2 == 0 else "F9F9F9"
        for col in range(1, 9):
            cell = ws3.cell(i, col)
            cell.fill = PatternFill("solid", start_color=row_bg)
            cell.border = thin; cell.font = data_font
            if col == 1: cell.value = i - 2; cell.alignment = Alignment(horizontal="center")
    ws3.merge_cells("A55:H55")
    ws3["A55"] = "STATUS LEGEND"
    ws3["A55"].font = Font(name="Arial", bold=True, color="FFFFFF")
    ws3["A55"].fill = PatternFill("solid", start_color="1F4E79")
    ws3["A55"].alignment = Alignment(horizontal="center")
    for idx, (status, color) in enumerate(status_map, 56):
        cell = ws3.cell(idx, 1, status)
        cell.font = Font(name="Arial", bold=True, size=9)
        cell.fill = PatternFill("solid", start_color=color); cell.border = thin
        cell.alignment = Alignment(horizontal="center")
        ws3.merge_cells(f"B{idx}:H{idx}")
        ws3.cell(idx, 2).fill = PatternFill("solid", start_color=color)
        ws3.cell(idx, 2).border = thin
    for col, w in zip(range(1, 9), [4, 28, 36, 24, 16, 22, 18, 36]):
        ws3.column_dimensions[get_column_letter(col)].width = w
    ws3.freeze_panes = "A3"

    wb.save(filepath)
    print(f"Excel saved: {filepath}")


# ═════════════════════════════════════════════
# DURATION FILTER — max 4 months
# ═════════════════════════════════════════════
def filter_by_duration(jobs):
    import re

    def is_acceptable(d_str):
        d = d_str.lower().strip()
        # Keep if duration not specified — benefit of doubt
        if any(x in d for x in ["not mentioned", "not disclosed", "n/a",
                                  "rolling", "flexible", "part-time",
                                  "as per", "structured"]):
            return True
        nums = list(map(int, re.findall(r"\d+", d)))
        if not nums:
            return True
        min_num = min(nums)
        if "year" in d:
            return False          # 1 year+ always rejected
        if "month" in d:
            return min_num <= 4   # 3-6 months → min=3 → keep; 6-12 → min=6 → reject
        if "week" in d:
            return min_num <= 16  # 16 weeks = 4 months
        return min_num <= 4

    kept    = [j for j in jobs if is_acceptable(j.get("duration", "N/A"))]
    removed = len(jobs) - len(kept)
    print(f"Duration filter: removed {removed} listings over 4 months, {len(kept)} kept.")
    return kept

# ═════════════════════════════════════════════
# SEND EMAIL
# ═════════════════════════════════════════════
def send_email(filepath, internships, new_count):
    today_str = date.today().strftime("%B %d, %Y")
    total     = len(internships)
    top_picks = [j for j in internships if j.get("fit_score", 0) >= 9]

    domain_counts   = {}
    platform_counts = {}
    for job in internships:
        d = job.get("domain", "Finance Operations")
        p = job.get("platform", "Other")
        domain_counts[d]   = domain_counts.get(d, 0) + 1
        platform_counts[p] = platform_counts.get(p, 0) + 1

    domain_lines   = "\n".join([f"   {d}: {c}" for d, c in sorted(domain_counts.items(), key=lambda x: -x[1])])
    platform_lines = "\n".join([f"   {p}: {c}" for p, c in sorted(platform_counts.items(), key=lambda x: -x[1])])

    # Top picks summary for email body
    newline = "\n"
    top_lines = ""
    for j in sorted(top_picks, key=lambda x: -x.get("fit_score", 0))[:5]:
        score    = j["fit_score"]
        title    = j["title"]
        company  = j["company"]
        location = j["location"]
        reason   = j.get("fit_reason", "")
        link     = j["link"]
        top_lines += "   [" + str(score) + "/10] " + title + " @ " + company + " — " + location + newline
        top_lines += "          -> " + reason + newline
        top_lines += "          -> " + link + newline + newline

    msg = MIMEMultipart()
    msg["From"]    = YOUR_EMAIL
    msg["To"]      = SEND_TO_EMAIL
    msg["Subject"] = f"🎯 {new_count} NEW Internships for Tanishq — {today_str} — {total} Total"

    no_picks_msg = "   No 9+ listings this run — check Sheet 1 for 7+ scores."
    top_section  = top_lines if top_lines else no_picks_msg

    body = f"""Hi Tanishq,

Your Internship Bot v5.0 report is ready!

DATE      : {today_str}
NEW TODAY : {new_count} listings (not seen before)
TOTAL     : {total} internships (including recurring)
LOCATION  : Delhi / Gurugram / Noida / Greater Noida
SCHEDULE  : Runs every Monday, Thursday, Saturday & Sunday

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🌟 TOP PICKS FOR YOU (AI Score 9+/10)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{top_section}

SOURCES SCRAPED:
{platform_lines}

DOMAIN BREAKDOWN:
{domain_lines}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
EXCEL HAS 6 SHEETS:
  Sheet 1 — All Internships (sorted by AI Fit Score, color coded)
  Sheet 2 — Top Picks (Fit Score 7+, best matches for you)
  Sheet 3 — 3-Month Internships (short, focused internships)
  Sheet 4 — 4-Month Internships (slightly longer internships)
  Sheet 5 — Domain Summary (count per domain + platforms)
  Sheet 6 — Application Tracker (track your progress)

AI FIT SCORE is based on your resume:
  BBA Finance & Banking | SEBI NISM | McKinsey Forward
  Yhills Finance Intern | 23 Ventures Founder Fellow
  Skills: Excel, Power BI, Financial Modeling, Strategy

HOW TO USE:
  1. Start with Sheet 2 (Top Picks) — highest fit first
  2. Click Apply Link to go directly to the listing
  3. Update STATUS column after applying
  4. Use Sheet 4 to track interview progress

Good luck!
-- Internship Bot v5.0 (GitHub Actions)
   Sources: Internshala + Indeed + Glassdoor + Naukri + Foundit + Jobaaj + Wellfound
   Dedup: Google Sheets | AI Scoring: Claude API
   Schedule: Mon / Thu / Sat / Sun
"""
    msg.attach(MIMEText(body, "plain"))
    with open(filepath, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        fname = os.path.basename(filepath)
        part.add_header("Content-Disposition", f"attachment; filename={fname}")
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(YOUR_EMAIL, YOUR_PASSWORD)
        server.sendmail(YOUR_EMAIL, SEND_TO_EMAIL, msg.as_string())
    print("Email sent!")

# ═════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════
def run():
    # Schedule guard
    today_weekday = date.today().weekday()
    if today_weekday not in RUN_ON_WEEKDAYS:
        day_name = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"][today_weekday]
        print(f"Skipping — today is {day_name}. Bot runs Mon / Thu / Sat / Sun only.")
        return

    print("Starting Internship Bot v5.0...")

    # ── Step 1: Scrape all sources ──
    all_jobs = []
    all_jobs += scrape_internshala()
    all_jobs += scrape_indeed()
    all_jobs += scrape_glassdoor()
    all_jobs += scrape_naukri()
    all_jobs += scrape_foundit()
    all_jobs += scrape_jobaaj()
    all_jobs += scrape_wellfound()        # NEW
    all_jobs += get_quality_listings()

    # ── Step 2: Local dedup (same run) ──
    seen_local, unique_all = set(), []
    for j in all_jobs:
        key = (j["title"].strip().lower()[:40], j["company"].strip().lower()[:30])
        if key not in seen_local:
            seen_local.add(key)
            unique_all.append(j)
    print(f"After local dedup: {len(unique_all)} unique listings")

    # ── Step 2b: Duration filter (max 4 months) ──
    unique_all = filter_by_duration(unique_all)

    # ── Step 3: Cross-run dedup via Google Sheets ──
    gs_client              = get_gsheet_client()
    seen_keys, dedup_ws    = load_seen_keys(gs_client)
    new_jobs               = filter_new_only(unique_all, seen_keys)
    new_count              = len(new_jobs)

    # Save new keys back to Google Sheets
    save_new_keys(dedup_ws, new_jobs)

    # For the Excel we show ALL unique (not just new) so context isn't lost,
    # but email subject highlights new count
    jobs_to_report = unique_all

    # ── Step 4: AI Fit Scoring ──
    print("Running AI fit scoring...")
    jobs_to_report = score_jobs_with_ai(jobs_to_report)

    # ── Step 5: Build Excel & Send Email ──
    today_str = date.today().strftime("%Y-%m-%d")
    filepath  = f"/tmp/Internships_Tanishq_{today_str}.xlsx"
    build_excel(jobs_to_report, filepath)
    send_email(filepath, jobs_to_report, new_count)
    print(f"Done! {len(jobs_to_report)} total listings, {new_count} new this run.")

if __name__ == "__main__":
    run()
