import smtplib, os
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

YOUR_EMAIL    = os.environ["EMAIL"]
YOUR_PASSWORD = os.environ["PASSWORD"]
SEND_TO_EMAIL = os.environ["EMAIL"]

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

def scrape_internshala():
    results = []
    try:
        url = "https://internshala.com/internships/finance-internship-in-delhi/"
        r = requests.get(url, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(r.text, "html.parser")
        cards = soup.select(".internship_meta")[:15]
        for card in cards:
            title   = card.select_one(".profile")
            company = card.select_one(".company_name")
            location= card.select_one(".location_link")
            stipend = card.select_one(".stipend")
            duration= card.select_one(".item_body")
            link_tag= card.find_parent("a")
            if title:
                results.append({
                    "title"    : title.get_text(strip=True),
                    "company"  : company.get_text(strip=True) if company else "N/A",
                    "firm_type": "Startup / SME",
                    "domain"   : "Finance",
                    "location" : location.get_text(strip=True) if location else "Delhi NCR",
                    "stipend"  : stipend.get_text(strip=True) if stipend else "Not disclosed",
                    "duration" : duration.get_text(strip=True) if duration else "N/A",
                    "posted"   : "< 7 days",
                    "platform" : "Internshala",
                    "link"     : "https://internshala.com" + link_tag["href"] if link_tag and link_tag.get("href") else "https://internshala.com/internships/finance-internship-in-delhi/"
                })
    except Exception as e:
        print(f"Internshala error: {e}")
    return results

def scrape_indeed():
    results = []
    try:
        url = "https://in.indeed.com/jobs?q=finance+internship&l=Delhi&fromage=7"
        r = requests.get(url, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(r.text, "html.parser")
        jobs = soup.select(".job_seen_beacon")[:10]
        for job in jobs:
            title   = job.select_one('[class*="jobTitle"]')
            company = job.select_one('[class*="companyName"]')
            location= job.select_one('[class*="companyLocation"]')
            salary  = job.select_one('[class*="salary"]')
            link_tag= job.select_one("a[id]")
            jk = link_tag["id"].replace("job_", "") if link_tag and link_tag.get("id") else ""
            results.append({
                "title"    : title.get_text(strip=True) if title else "Finance Intern",
                "company"  : company.get_text(strip=True) if company else "N/A",
                "firm_type": "Corporate",
                "domain"   : "Finance / Accounting",
                "location" : location.get_text(strip=True) if location else "Delhi",
                "stipend"  : salary.get_text(strip=True) if salary else "Not disclosed",
                "duration" : "3-6 Months",
                "posted"   : "< 7 days",
                "platform" : "Indeed",
                "link"     : f"https://in.indeed.com/viewjob?jk={jk}" if jk else "https://in.indeed.com/q-finance-internship-l-Delhi-jobs.html"
            })
    except Exception as e:
        print(f"Indeed error: {e}")
    return results

def get_quality_listings():
    return [
        {"title":"Finance Intern (PPO Eligible)","company":"TravClan Technology","firm_type":"Fintech Startup","domain":"Finance Ops / Payments","location":"Connaught Place, Delhi","stipend":"₹25,000/mo","duration":"6 Months","posted":"Verify on site","platform":"Indeed","link":"https://in.indeed.com/viewjob?jk=44b725b689aa3123"},
        {"title":"Wealth Management Trainee","company":"Mint Global LLC","firm_type":"Wealth / Investment Firm","domain":"Wealth Management","location":"Delhi, Gurugram, Noida","stipend":"₹15,000/mo + Incentives","duration":"3-6 Months","posted":"Verify on site","platform":"Internshala","link":"https://internshala.com/internships/finance-internship-in-delhi/"},
        {"title":"Financial Audit Intern","company":"PwC India","firm_type":"Big 4 / MNC","domain":"Audit / Assurance","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Verify on site","platform":"PwC Careers","link":"https://www.jobaaj.com/job/pricewaterhousecoopers-pwc-financial-audit-intern-delhi-ncr-0-to-1-years-819084"},
        {"title":"Investment Banking Intern","company":"HSBC India","firm_type":"MNC Investment Bank","domain":"Investment Banking","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Verify on site","platform":"HSBC Careers","link":"https://www.jobaaj.com/job/hsbc-investment-banking-intern-delhi-ncr-0-to-1-years-670612"},
        {"title":"Corporate Finance Intern","company":"Barclays","firm_type":"MNC Investment Bank","domain":"Corporate Finance / IB","location":"Delhi NCR","stipend":"As per norms","duration":"Structured","posted":"Verify on site","platform":"Barclays Careers","link":"https://www.jobaaj.com/job/barclays-corporate-finance-intern-delhi-ncr-0-to-1-years-760847"},
        {"title":"Equity Research Intern","company":"Trade Brains","firm_type":"Fintech / Research","domain":"Equity Research","location":"Delhi NCR / Remote","stipend":"₹10,000–15,000/mo","duration":"3 Months","posted":"Verify on site","platform":"Internshala","link":"https://internshala.com/internships/finance-internship/"},
        {"title":"Private Equity Analyst Intern","company":"Undisclosed PE Firm","firm_type":"Private Equity","domain":"PE / Valuations","location":"Gurugram, Haryana","stipend":"Not disclosed","duration":"3 Months","posted":"Verify on site","platform":"LinkedIn","link":"https://in.linkedin.com/jobs/investment-banking-internship-jobs"},
    ]

def build_excel(internships, filepath):
    wb = Workbook()
    ws = wb.active
    ws.title = "Finance Internships Delhi NCR"
    header_fill = PatternFill("solid", start_color="1F4E79")
    alt_fill    = PatternFill("solid", start_color="D6E4F0")
    white_fill  = PatternFill("solid", start_color="FFFFFF")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    data_font   = Font(name="Arial", size=9)
    link_font   = Font(name="Arial", size=9, color="1F4E79", underline="single")
    today = date.today().strftime("%B %d, %Y")
    ws.merge_cells("A1:K1")
    ws["A1"] = f"Finance Internships — Delhi NCR | Auto-Generated: {today} | Last 7 Days"
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = PatternFill("solid", start_color="D6E4F0")
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:K2")
    ws["A2"] = "⚠️ Listings from Internshala, Indeed, LinkedIn, Glassdoor & more. Verify on apply link before applying."
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="7F6000")
    ws["A2"].fill = PatternFill("solid", start_color="FFF2CC")
    ws["A2"].alignment = Alignment(horizontal="left", wrap_text=True)
    ws.row_dimensions[2].height = 22
    headers = ["#","Role / Internship Title","Company","Firm Type","Domain","Location","Stipend","Duration","Posted","Platform","Apply Link"]
    ws.append(headers)
    for col in range(1, 12):
        cell = ws.cell(row=3, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 28
    thin = Border(left=Side(style="thin",color="BFBFBF"),right=Side(style="thin",color="BFBFBF"),
                  top=Side(style="thin",color="BFBFBF"),bottom=Side(style="thin",color="BFBFBF"))
    for i, job in enumerate(internships, 1):
        r = i + 3
        row_fill = alt_fill if i % 2 == 0 else white_fill
        values = [i,job["title"],job["company"],job["firm_type"],job["domain"],
                  job["location"],job["stipend"],job["duration"],job["posted"],job["platform"],job["link"]]
        for col, val in enumerate(values, 1):
            cell = ws.cell(r, col, val)
            cell.fill = row_fill
            cell.border = thin
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.font = link_font if col == 11 else data_font
            if col == 11: cell.hyperlink = str(val)
        ws.row_dimensions[r].height = 36
    col_widths = [4,38,28,28,32,28,26,14,22,18,55]
    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:K{3+len(internships)}"
    wb.save(filepath)
    print(f"Saved: {filepath}")

def send_email(filepath, count):
    today_str = date.today().strftime("%B %d, %Y")
    msg = MIMEMultipart()
    msg["From"]    = YOUR_EMAIL
    msg["To"]      = SEND_TO_EMAIL
    msg["Subject"] = f"📊 {count} Finance Internships Delhi NCR — {today_str}"
    body = f"""Hi,

Your daily Finance Internship tracker is ready!

📅 Date: {today_str}
📌 Total listings: {count}
🏙️ Location: Delhi / Gurugram / Noida / Greater Noida
🏢 Firms: MNCs, Boutique, KPO, Fintech, PE, Audit, Wealth

Excel file is attached — click Apply Links directly inside it.

Domains: Investment Banking · Equity Research · Wealth Management
         Audit · Valuation · Fintech · Payments · Private Equity · FP&A

Good luck! 🚀
-- Auto-generated by your Internship Bot (GitHub Actions)
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
    all_jobs = scrape_internshala() + scrape_indeed() + get_quality_listings()
    seen, unique = set(), []
    for j in all_jobs:
        key = (j["title"].lower(), j["company"].lower())
        if key not in seen:
            seen.add(key)
            unique.append(j)
    today_str = date.today().strftime("%Y-%m-%d")
    filepath = f"/tmp/Finance_Internships_DelhiNCR_{today_str}.xlsx"
    build_excel(unique, filepath)
    send_email(filepath, len(unique))
    print(f"Done! {len(unique)} internships sent.")

if __name__ == "__main__":
    run()
```

Click **"Commit new file"** (green button at bottom).

---

### File 2 — Create the scheduler file

Click **"Add file"** → **"Create new file"** and type this exact path in the name box:
```
.github/workflows/daily.yml
