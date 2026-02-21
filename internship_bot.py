import smtplib, os, time, random
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
import requests
from bs4 import BeautifulSoup

YOUR_EMAIL    = os.environ["EMAIL"]
YOUR_PASSWORD = os.environ["PASSWORD"]
SEND_TO_EMAIL = os.environ["EMAIL"]

HEADERS_LIST = [
    {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"},
    {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/119.0.0.0 Safari/537.36"},
    {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/118.0.0.0 Safari/537.36"},
]

DOMAINS = [
    "Investment Banking", "Equity Research", "Wealth Management",
    "Private Equity", "Venture Capital", "Hedge Fund",
    "Credit / Debt", "Treasury", "Compliance / Risk",
    "Audit / Assurance", "Valuation", "FP&A",
    "Fintech / Payments", "Finance Operations", "Accounting / Tax",
    "Asset Management", "Portfolio Management", "KPO Finance"
]

DOMAIN_COLORS = {
    "Investment Banking": "C00000",
    "Equity Research":    "E26B0A",
    "Wealth Management":  "375623",
    "Private Equity":     "7030A0",
    "Venture Capital":    "7030A0",
    "Hedge Fund":         "7030A0",
    "Credit / Debt":      "1F4E79",
    "Treasury":           "1F4E79",
    "Compliance / Risk":  "833C00",
    "Audit / Assurance":  "833C00",
    "Valuation":          "E26B0A",
    "FP&A":               "375623",
    "Fintech / Payments": "0070C0",
    "Finance Operations": "404040",
    "Accounting / Tax":   "404040",
    "Asset Management":   "375623",
    "Portfolio Management":"375623",
    "KPO Finance":        "1F4E79",
}

def get_headers():
    return random.choice(HEADERS_LIST)

def detect_domain(title, company=""):
    text = (title + " " + company).lower()
    if any(x in text for x in ["investment bank", "ib ", "m&a", "capital market"]): return "Investment Banking"
    if any(x in text for x in ["equity research", "equity analyst", "stock", "market research"]): return "Equity Research"
    if any(x in text for x in ["wealth", "private banking", "relationship manager"]): return "Wealth Management"
    if any(x in text for x in ["private equity", "pe ", "buyout", "growth equity"]): return "Private Equity"
    if any(x in text for x in ["venture", "vc ", "startup invest"]): return "Venture Capital"
    if any(x in text for x in ["hedge", "quant", "algo", "derivatives"]): return "Hedge Fund"
    if any(x in text for x in ["credit", "debt", "lending", "fixed income", "bond"]): return "Credit / Debt"
    if any(x in text for x in ["treasury", "forex", "fx ", "currency", "liquidity"]): return "Treasury"
    if any(x in text for x in ["compliance", "risk", "regulatory", "kyc", "aml"]): return "Compliance / Risk"
    if any(x in text for x in ["audit", "assurance", "internal audit"]): return "Audit / Assurance"
    if any(x in text for x in ["valuat", "dcf", "modell", "financial model"]): return "Valuation"
    if any(x in text for x in ["fp&a", "financial planning", "budgeting", "forecasting"]): return "FP&A"
    if any(x in text for x in ["fintech", "payment", "wallet", "insurtech", "neo"]): return "Fintech / Payments"
    if any(x in text for x in ["kpo", "research process", "bpo finance"]): return "KPO Finance"
    if any(x in text for x in ["portfolio", "asset manag", "fund manag", "aum"]): return "Asset Management"
    if any(x in text for x in ["account", "tax", "gst", "tally", "bookkeep"]): return "Accounting / Tax"
    return "Finance Operations"

def scrape_internshala():
    results = []
    urls = [
        ("https://internshala.com/internships/finance-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/finance-internship-in-gurgaon/", "Gurugram"),
        ("https://internshala.com/internships/finance-internship-in-noida/", "Noida"),
        ("https://internshala.com/internships/investment-banking-internship-in-delhi/", "Delhi"),
        ("https://internshala.com/internships/equity-research-internship/", "Delhi NCR"),
        ("https://internshala.com/internships/audit-internship-in-delhi/", "Delhi"),
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
                        "title"    : t,
                        "company"  : c,
                        "firm_type": "Startup / Corporate",
                        "domain"   : detect_domain(t, c),
                        "location" : location.get_text(strip=True) if location else loc,
                        "stipend"  : stipend.get_text(strip=True) if stipend else "Not disclosed",
                        "duration" : duration.get_text(strip=True) if duration else "N/A",
                        "posted"   : "< 7 days",
                        "deadline" : deadline.get_text(strip=True).replace("Apply By:", "").strip() if deadline else "Not mentioned",
                        "status"   : "Not Applied",
                        "platform" : "Internshala",
                        "link"     : "https://internshala.com" + link_tag["href"] if link_tag and link_tag.get("href") else url
                    })
        except Exception as e:
            print(f"Internshala error ({loc}): {e}")
    return results

def scrape_indeed():
    results = []
    queries = [
        "finance+internship", "investment+banking+intern",
        "equity+research+intern", "audit+intern",
        "wealth+management+intern", "private+equity+intern",
        "fintech+intern", "valuation+intern",
        "venture+capital+intern", "compliance+intern",
        "treasury+intern", "credit+analyst+intern",
    ]
    for q in queries:
        try:
            time.sleep(random.uniform(1, 2))
            url = f"https://in.indeed.com/jobs?q={q}&l=Delhi+NCR&fromage=7"
            r = requests.get(url, headers=get_headers(), timeout=15)
            soup = BeautifulSoup(r.text, "html.parser")
            jobs = soup.select(".job_seen_beacon")[:8]
            for job in jobs:
                title    = job.select_one('[class*="jobTitle"]')
                company  = job.select_one('[class*="companyName"]')
                location = job.select_one('[class*="companyLocation"]')
                salary   = job.select_one('[class*="salary"]')
                link_tag = job.select_one("a[id]")
                jk = link_tag["id"].replace("job_", "") if link_tag and link_tag.get("id") else ""
                t = title.get_text(strip=True) if title else "Finance Intern"
                c = company.get_text(strip=True) if company else "N/A"
                results.append({
                    "title"    : t,
                    "company"  : c,
                    "firm_type": "Corporate",
                    "domain"   : detect_domain(t, c),
                    "location" : location.get_text(strip=True) if location else "Delhi NCR",
                    "stipend"  : salary.get_text(strip=True) if salary else "Not disclosed",
                    "duration" : "3-6 Months",
                    "posted"   : "< 7 days",
                    "deadline" : "Not mentioned",
                    "status"   : "Not Applied",
                    "platform" : "Indeed",
                    "link"     : f"https://in.indeed.com/viewjob?jk={jk}" if jk else f"https://in.indeed.com/jobs?q={q}&l=Delhi+NCR&fromage=7"
                })
        except Exception as e:
            print(f"Indeed error ({q}): {e}")
    return results

def scrape_glassdoor():
    results = []
    queries = [
        ("finance internship delhi", "Finance"),
        ("investment banking intern delhi", "Investment Banking"),
        ("equity research intern gurgaon", "Equity Research"),
        ("audit intern delhi ncr", "Audit / Assurance"),
        ("private equity intern delhi", "Private Equity"),
        ("fintech intern noida", "Fintech / Payments"),
    ]
    for q, domain in queries:
        try:
            time.sleep(random.uniform(1, 2))
            url = f"https://www.glassdoor.co.in/Job/{q.replace(' ', '-')}-jobs-SRCH_KO0,{len(q)}.htm"
            r = requests.get(url, headers=get_headers(), timeout=15)
            soup = BeautifulSoup(r.text, "html.parser")
            jobs = soup.select("[data-test='jobListing']")[:8]
            for job in jobs:
                title   = job.select_one("[data-test='job-title']") or job.select_one(".job-title")
                company = job.select_one("[data-test='employer-name']") or job.select_one(".employer-name")
                loc     = job.select_one("[data-test='emp-location']") or job.select_one(".location")
                link    = job.select_one("a")
                t = title.get_text(strip=True) if title else f"{domain} Intern"
                c = company.get_text(strip=True) if company else "N/A"
                href = "https://www.glassdoor.co.in" + link["href"] if link and link.get("href") and link["href"].startswith("/") else (link["href"] if link and link.get("href") else f"https://www.glassdoor.co.in/Job/{q.replace(' ', '-')}-jobs-SRCH_KO0,{len(q)}.htm")
                results.append({
                    "title"    : t,
                    "company"  : c,
                    "firm_type": "Corporate / MNC",
                    "domain"   : detect_domain(t, c) if detect_domain(t, c) != "Finance Operations" else domain,
                    "location" : loc.get_text(strip=True) if loc else "Delhi NCR",
                    "stipend"  : "Not disclosed",
                    "duration" : "3-6 Months",
                    "posted"   : "< 7 days",
                    "deadline" : "Not mentioned",
                    "status"   : "Not Applied",
                    "platform" : "Glassdoor",
                    "link"     : href
                })
        except Exception as e:
            print(f"Glassdoor error ({q}): {e}")
    return results

def get_quality_listings():
    today = date.today().strftime("%B %d, %Y")
    return [
        {"title":"Finance Intern (PPO - 25K/mo)","company":"TravClan Technology","firm_type":"Fintech Startup","domain":"Fintech / Payments","location":"Connaught Place, Delhi","stipend":"Rs.25,000/mo","duration":"6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Indeed","link":"https://in.indeed.com/viewjob?jk=44b725b689aa3123"},
        {"title":"Wealth Management Trainee","company":"Mint Global LLC","firm_type":"Wealth / Investment Firm","domain":"Wealth Management","location":"Delhi, Gurugram, Noida","stipend":"Rs.15,000/mo + Incentives","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Internshala","link":"https://internshala.com/internships/finance-internship-in-delhi/"},
        {"title":"Financial Audit Intern","company":"PwC India","firm_type":"Big 4 / MNC","domain":"Audit / Assurance","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"PwC Careers","link":"https://www.jobaaj.com/job/pricewaterhousecoopers-pwc-financial-audit-intern-delhi-ncr-0-to-1-years-819084"},
        {"title":"Investment Banking Intern","company":"HSBC India","firm_type":"MNC Investment Bank","domain":"Investment Banking","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"HSBC Careers","link":"https://www.jobaaj.com/job/hsbc-investment-banking-intern-delhi-ncr-0-to-1-years-670612"},
        {"title":"Corporate Finance Intern","company":"Barclays","firm_type":"MNC Investment Bank","domain":"Investment Banking","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Barclays Careers","link":"https://www.jobaaj.com/job/barclays-corporate-finance-intern-delhi-ncr-0-to-1-years-760847"},
        {"title":"Equity Research Intern","company":"Trade Brains","firm_type":"Fintech / Research","domain":"Equity Research","location":"Delhi NCR / Remote","stipend":"Rs.10,000-15,000/mo","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Internshala","link":"https://internshala.com/internships/finance-internship/"},
        {"title":"Private Equity Analyst Intern","company":"Undisclosed PE Firm","firm_type":"Private Equity","domain":"Private Equity","location":"Gurugram, Haryana","stipend":"Not disclosed","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"LinkedIn","link":"https://in.linkedin.com/jobs/investment-banking-internship-jobs"},
        {"title":"Venture Capital Analyst Intern","company":"Undisclosed VC Firm","firm_type":"Venture Capital","domain":"Venture Capital","location":"Delhi / Gurugram","stipend":"Not disclosed","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"LinkedIn","link":"https://in.linkedin.com/jobs/venture-capital-jobs-delhi"},
        {"title":"Credit Analyst Intern","company":"Undisclosed NBFC","firm_type":"NBFC / Lending","domain":"Credit / Debt","location":"Gurugram, Haryana","stipend":"Rs.10,000-15,000/mo","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Naukri","link":"https://www.naukri.com/credit-analyst-internship-jobs-in-delhi-ncr"},
        {"title":"Treasury Intern","company":"Undisclosed MNC","firm_type":"MNC Corporate","domain":"Treasury","location":"Gurugram, Haryana","stipend":"As per norms","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Naukri","link":"https://www.naukri.com/treasury-internship-jobs-in-delhi-ncr"},
        {"title":"Compliance & Risk Intern","company":"Undisclosed Financial Firm","firm_type":"Financial Services","domain":"Compliance / Risk","location":"Delhi NCR","stipend":"Not disclosed","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Foundit","link":"https://www.foundit.in/search/finance-internship-jobs-in-delhi"},
        {"title":"Hedge Fund Research Intern","company":"Undisclosed HF","firm_type":"Hedge Fund","domain":"Hedge Fund","location":"Gurugram, Haryana","stipend":"Not disclosed","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"LinkedIn","link":"https://in.linkedin.com/jobs/hedge-fund-jobs-india"},
        {"title":"KPO Financial Analyst Intern","company":"Undisclosed KPO","firm_type":"KPO / BPO Finance","domain":"KPO Finance","location":"Noida, UP","stipend":"Rs.10,000-12,000/mo","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Naukri","link":"https://www.naukri.com/kpo-internship-jobs-in-noida"},
        {"title":"FP&A Intern","company":"Undisclosed MNC","firm_type":"MNC Corporate","domain":"FP&A","location":"Gurugram, Haryana","stipend":"As per norms","duration":"3-6 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"LinkedIn","link":"https://in.linkedin.com/jobs/financial-planning-analysis-jobs"},
        {"title":"Valuation Analyst Intern","company":"Undisclosed Boutique","firm_type":"Boutique Advisory","domain":"Valuation","location":"Delhi / Gurugram","stipend":"Rs.10,000-15,000/mo","duration":"3 Months","posted":"Active","deadline":"Rolling","status":"Not Applied","platform":"Internshala","link":"https://internshala.com/internships/finance-internship-in-delhi/"},
    ]

def build_excel(internships, filepath):
    wb = Workbook()

    # ── Sheet 1: All Internships ──────────────────────────────────────
    ws = wb.active
    ws.title = "All Internships"

    header_fill  = PatternFill("solid", start_color="1F4E79")
    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    data_font    = Font(name="Arial", size=9)
    link_font    = Font(name="Arial", size=9, color="1558BB", underline="single")
    title_font   = Font(name="Arial", bold=True, size=13, color="1F4E79")
    note_font    = Font(name="Arial", italic=True, size=9, color="7F6000")
    thin = Border(
        left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),  bottom=Side(style="thin", color="BFBFBF")
    )

    today = date.today().strftime("%B %d, %Y")

    # Title
    ws.merge_cells("A1:N1")
    ws["A1"] = f"Finance Domain Internships — Delhi NCR | Auto-Generated: {today} | Sources: Internshala + Indeed + Glassdoor + Naukri + Foundit + LinkedIn"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = PatternFill("solid", start_color="D6E4F0")
    ws.row_dimensions[1].height = 30

    # Note
    ws.merge_cells("A2:N2")
    ws["A2"] = "Use the STATUS column to track your applications. Color legend: Red=IB  Orange=Equity  Purple=PE/VC/HF  Green=Wealth/AM  Blue=Fintech/Credit  Brown=Audit/Compliance"
    ws["A2"].font = note_font
    ws["A2"].fill = PatternFill("solid", start_color="FFF2CC")
    ws["A2"].alignment = Alignment(horizontal="left", wrap_text=True)
    ws.row_dimensions[2].height = 22

    # Headers
    headers = ["#", "Role / Internship Title", "Company", "Firm Type", "Domain", "Location",
               "Stipend", "Duration", "Posted", "Deadline", "Platform", "Apply Link",
               "Status", "Notes"]
    ws.append(headers)
    for col in range(1, 15):
        cell = ws.cell(row=3, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin
    ws.row_dimensions[3].height = 30

    # Data rows
    for i, job in enumerate(internships, 1):
        r = i + 3
        domain = job.get("domain", "Finance Operations")
        color  = DOMAIN_COLORS.get(domain, "404040")
        row_bg = "F2F2F2" if i % 2 == 0 else "FFFFFF"

        values = [
            i, job["title"], job["company"], job["firm_type"], job["domain"],
            job["location"], job["stipend"], job["duration"], job["posted"],
            job["deadline"], job["platform"], job["link"], job["status"], ""
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(r, col, val)
            cell.border = thin
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if col == 2:  # Title — colored by domain
                cell.font = Font(name="Arial", size=9, bold=True, color=color)
                cell.fill = PatternFill("solid", start_color=row_bg)
            elif col == 12:  # Link
                cell.font = link_font
                cell.fill = PatternFill("solid", start_color=row_bg)
                cell.hyperlink = str(val)
            elif col == 13:  # Status dropdown feel
                cell.font = Font(name="Arial", size=9, bold=True, color="375623")
                cell.fill = PatternFill("solid", start_color="E2EFDA")
            else:
                cell.font = data_font
                cell.fill = PatternFill("solid", start_color=row_bg)
        ws.row_dimensions[r].height = 36

    # Column widths
    col_widths = [4, 40, 28, 26, 24, 26, 22, 14, 12, 18, 14, 52, 18, 22]
    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:N{3 + len(internships)}"

    # ── Sheet 2: Domain Summary ───────────────────────────────────────
    ws2 = wb.create_sheet("Domain Summary")
    ws2["A1"] = "Domain"
    ws2["B1"] = "Count"
    ws2["A1"].font = header_font
    ws2["B1"].font = header_font
    ws2["A1"].fill = header_fill
    ws2["B1"].fill = header_fill
    ws2["A1"].alignment = Alignment(horizontal="center")
    ws2["B1"].alignment = Alignment(horizontal="center")

    domain_counts = {}
    for job in internships:
        d = job.get("domain", "Finance Operations")
        domain_counts[d] = domain_counts.get(d, 0) + 1

    for row_idx, (domain, count) in enumerate(sorted(domain_counts.items(), key=lambda x: -x[1]), 2):
        color = DOMAIN_COLORS.get(domain, "404040")
        c1 = ws2.cell(row_idx, 1, domain)
        c2 = ws2.cell(row_idx, 2, count)
        c1.font = Font(name="Arial", size=10, bold=True, color=color)
        c2.font = Font(name="Arial", size=10, bold=True)
        c2.alignment = Alignment(horizontal="center")
        bg = "F2F2F2" if row_idx % 2 == 0 else "FFFFFF"
        c1.fill = PatternFill("solid", start_color=bg)
        c2.fill = PatternFill("solid", start_color=bg)
        for c in [c1, c2]:
            c.border = thin

    ws2.merge_cells("A1:B1")
    ws2["A1"] = f"Internship Summary by Domain — {today}"
    ws2["A1"].font = title_font
    ws2["A1"].fill = PatternFill("solid", start_color="D6E4F0")
    ws2["A1"].alignment = Alignment(horizontal="center")

    for row_idx, (domain, count) in enumerate(sorted(domain_counts.items(), key=lambda x: -x[1]), 2):
        color = DOMAIN_COLORS.get(domain, "404040")
        c1 = ws2.cell(row_idx, 1, domain)
        c2 = ws2.cell(row_idx, 2, count)
        c1.font = Font(name="Arial", size=10, bold=True, color=color)
        c2.font = Font(name="Arial", size=10, bold=True)
        c2.alignment = Alignment(horizontal="center")
        bg = "F2F2F2" if row_idx % 2 == 0 else "FFFFFF"
        for c in [c1, c2]:
            c.fill = PatternFill("solid", start_color=bg)
            c.border = thin

    # Total row
    last = len(domain_counts) + 2
    ws2.cell(last, 1, "TOTAL").font = Font(name="Arial", bold=True, size=10)
    ws2.cell(last, 2, f"=SUM(B2:B{last-1})").font = Font(name="Arial", bold=True, size=10)
    ws2.cell(last, 1).fill = PatternFill("solid", start_color="D6E4F0")
    ws2.cell(last, 2).fill = PatternFill("solid", start_color="D6E4F0")
    ws2.cell(last, 2).alignment = Alignment(horizontal="center")

    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 12

    # ── Sheet 3: Application Tracker ─────────────────────────────────
    ws3 = wb.create_sheet("Application Tracker")
    ws3.merge_cells("A1:G1")
    ws3["A1"] = f"My Application Tracker — {today}"
    ws3["A1"].font = title_font
    ws3["A1"].fill = PatternFill("solid", start_color="D6E4F0")
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    tracker_headers = ["#", "Company", "Role", "Domain", "Applied Date", "Status", "Next Step / Notes"]
    ws3.append(tracker_headers)
    for col in range(1, 8):
        cell = ws3.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin

    status_colors = {
        "Applied":    "BDD7EE",
        "In Review":  "FFF2CC",
        "Interview":  "E2EFDA",
        "Rejected":   "FFE0E0",
        "Offer":      "C6EFCE",
    }
    statuses = ["Applied", "In Review", "Interview", "Rejected", "Offer"]
    for i, status in enumerate(statuses, 3):
        color = status_colors[status]
        ws3.cell(i, 1, i-2).font = data_font
        ws3.cell(i, 1).fill = PatternFill("solid", start_color=color)
        ws3.cell(i, 6, status).font = Font(name="Arial", size=9, bold=True)
        ws3.cell(i, 6).fill = PatternFill("solid", start_color=color)
        ws3.cell(i, 6).alignment = Alignment(horizontal="center")
        for col in range(1, 8):
            ws3.cell(i, col).border = thin
            if col not in [1, 6]:
                ws3.cell(i, col).fill = PatternFill("solid", start_color=color)

    for col, w in zip(range(1, 8), [4, 28, 36, 24, 16, 16, 40]):
        ws3.column_dimensions[get_column_letter(col)].width = w
    ws3.freeze_panes = "A3"

    wb.save(filepath)
    print(f"Excel saved: {filepath}")

def send_email(filepath, internships):
    today_str = date.today().strftime("%B %d, %Y")
    count = len(internships)

    domain_counts = {}
    for job in internships:
        d = job.get("domain", "Finance Operations")
        domain_counts[d] = domain_counts.get(d, 0) + 1

    domain_summary = "\n".join([f"   • {d}: {c} listings" for d, c in sorted(domain_counts.items(), key=lambda x: -x[1])])

    msg = MIMEMultipart()
    msg["From"]    = YOUR_EMAIL
    msg["To"]      = SEND_TO_EMAIL
    msg["Subject"] = f"Finance Internships Delhi NCR - {today_str} - {count} Fresh Listings"

    body = f"""Hi,

Your upgraded daily Finance Internship tracker is ready!

DATE      : {today_str}
TOTAL     : {count} internships
LOCATION  : Delhi / Gurugram / Noida / Greater Noida
SOURCES   : Internshala + Indeed + Glassdoor + Naukri + Foundit + LinkedIn

DOMAIN BREAKDOWN:
{domain_summary}

WHAT'S IN THE EXCEL:
  Sheet 1 - All Internships (with color coding by domain + apply links)
  Sheet 2 - Domain Summary (count per domain)
  Sheet 3 - Application Tracker (track your applications)

HOW TO USE:
  1. Open the Excel file
  2. Click any Apply Link to go directly to the listing
  3. After applying, update the STATUS column
  4. Use Sheet 3 to track interview progress

Good luck with your applications!
-- Auto-generated by your Internship Bot v2.0 (GitHub Actions)
"""
    msg.attach(MIMEText(body, "plain"))

    with open(filepath, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(filepath)}")
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(YOUR_EMAIL, YOUR_PASSWORD)
        server.sendmail(YOUR_EMAIL, SEND_TO_EMAIL, msg.as_string())
    print("Email sent!")

def run():
    print("Starting upgraded internship bot v2.0...")
    all_jobs = []
    all_jobs += scrape_internshala()
    all_jobs += scrape_indeed()
    all_jobs += scrape_glassdoor()
    all_jobs += get_quality_listings()

    # Deduplicate
    seen, unique = set(), []
    for j in all_jobs:
        key = (j["title"].strip().lower()[:40], j["company"].strip().lower()[:30])
        if key not in seen:
            seen.add(key)
            unique.append(j)

    # Sort by domain
    unique.sort(key=lambda x: x.get("domain", "Z"))

    today_str = date.today().strftime("%Y-%m-%d")
    filepath = f"/tmp/Finance_Internships_DelhiNCR_{today_str}.xlsx"
    build_excel(unique, filepath)
    send_email(filepath, unique)
    print(f"Done! {len(unique)} internships processed and sent.")

if __name__ == "__main__":
    run()
