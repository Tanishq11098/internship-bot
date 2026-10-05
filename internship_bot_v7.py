"""
Internship Bot v7  (Phase 1 + Phase 2)

Phase 1: PPO / role-type tagging, detail-page fetch, full-time queries,
         relevance scoring, new-only digest, weekly full-time digest,
         batched Google Sheets writes, html escaping, scraper health alert.
Phase 2: Application tracker tab with auto-added top listings, deadline
         extraction, and follow-up / deadline reminders by email.
"""
import os
import re
import json
import time
import random
import html
import smtplib
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import urlparse, urlunparse, parse_qsl, urlencode, quote_plus

import requests
import urllib3
import pandas as pd
import gspread
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ------------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------------
IST_NOW = datetime.now(timezone.utc).replace(tzinfo=None) + timedelta(hours=5, minutes=30)
TODAY = IST_NOW
TODAY_D = TODAY.date()

EMAIL = os.environ.get("EMAIL")
EMAIL_PASS = os.environ.get("PASSWORD")
TO_EMAIL = os.environ.get("EMAIL_TO") or EMAIL
GSHEET_ID = os.environ.get("GSHEET_ID")
SA_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

FORCE_FULLTIME_DIGEST = os.environ.get("FORCE_FULLTIME_DIGEST", "").lower() in ("1", "true", "yes")
STRICT_LOCATION = os.environ.get("STRICT_LOCATION", "").lower() in ("1", "true", "yes")

MAX_DETAIL_FETCH = 40        # detail pages fetched per run
TRACKER_MIN_SCORE = 60       # auto-add to tracker at or above this score
TRACKER_MAX_ADD = 10         # max auto-added rows per run
DEADLINE_WARN_DAYS = 3
FOLLOWUP_AFTER_DAYS = 7
GHOSTED_AFTER_DAYS = 21
STALE_TO_APPLY_DAYS = 5

OUTPUT_FILE = "internships.xlsx"
UNVERIFIED_HOSTS = {"www.timesjobs.com"}   # broken cert chain on CI runners

print(f"[CONFIG] EMAIL={'set' if EMAIL else 'MISSING'} PASSWORD={'set' if EMAIL_PASS else 'MISSING'} "
      f"GSHEET={'set' if GSHEET_ID else 'not set'} SA_JSON={'set' if SA_JSON else 'not set'}")

# (shine slug, timesjobs keyword)
QUERIES = [
    ("finance-intern", "finance intern"),
    ("investment-banking-intern", "investment banking intern"),
    ("equity-research-intern", "equity research intern"),
    ("venture-capital-intern", "venture capital intern"),
    ("financial-analyst-fresher", "financial analyst fresher"),
    ("investment-analyst-fresher", "investment analyst fresher"),
    ("management-trainee-finance", "management trainee finance"),
    ("graduate-trainee-finance", "graduate trainee finance"),
]

TIER1 = [
    "goldman sachs", "morgan stanley", "jpmorgan", "j.p. morgan", "hsbc", "barclays", "citi",
    "citibank", "blackrock", "bcg", "bain", "mckinsey", "kotak", "axis bank", "axis capital",
    "hdfc", "icici", "avendus", "ambit", "motilal oswal", "edelweiss", "jm financial",
    "nomura", "deutsche bank", "peak xv", "sequoia", "accel", "blume", "elevation capital",
    "3one4", "kedaara", "chryscapital", "a.t. kearney", "oliver wyman",
]
TIER1_RE = re.compile(r"\b(?:" + "|".join(re.escape(t) for t in TIER1) + r")\b", re.I)

EXCLUDED_LINKS = set()  # add normalized links you never want to see

# ------------------------------------------------------------------
# TEXT RULES
# ------------------------------------------------------------------
NOISE_RE = re.compile(
    r"software|developer|devops|data scientist|sales|marketing|recruit|\bhr\b|human resource|"
    r"telecall|\bbpo\b|customer|graphic|content|business development|\bbde\b|logistics|"
    r"senior|\bsr\.?\b|manager|\bvp\b|vice president|director|head of|\blead\b|principal|chief|\bavp\b",
    re.I)

DOMAINS = [  # (name, weight, regex) - first match wins
    ("VC / PE", 28, re.compile(r"venture|\bvc\b|private equity|growth equity|deal sourcing", re.I)),
    ("Investment Banking", 28, re.compile(r"investment bank|\bibd\b|m&a|mergers|capital markets|corporate finance|transaction advisory|valuation", re.I)),
    ("Equity Research", 24, re.compile(r"equity research|research analyst|sell[- ]side|buy[- ]side|credit research|financial research|macro", re.I)),
    ("Portfolio / Asset Mgmt", 22, re.compile(r"portfolio|asset management|wealth|fund management|investment analyst|investment management|treasury", re.I)),
    ("Consulting / Strategy", 20, re.compile(r"consult|strategy|business analyst", re.I)),
    ("Finance (General)", 12, re.compile(r"financ|banking|fp&a|credit analyst|risk|analyst", re.I)),
]

PPO_RE = re.compile(
    r"\bppo\b|pre[- ]?placement|(?:conversion|convert\w*)\s+(?:to|into)\s+(?:a\s+)?(?:full[- ]?time|permanent)|"
    r"(?:full[- ]?time|permanent)\s+(?:job\s+)?(?:offer|conversion)", re.I)
INTERN_RE = re.compile(r"\b(?:intern|interns|internship|summer analyst|summer associate|apprentice)\b", re.I)
GRAD_RE = re.compile(
    r"graduate\s+(?:program|programme|trainee|analyst|scheme)|management trainee|campus|"
    r"2027\s+(?:batch|graduate)|class of 2027", re.I)
CREDENTIAL_RE = re.compile(
    r"(?:mba|pgdm)\s+(?:only|mandatory|required|students only)|pursuing\s+(?:an?\s+)?(?:mba|pgdm|cfa)\b|"
    r"\bca\s+(?:inter|final|qualified|only|required|mandatory)|cfa\s+level", re.I)
EXP_RE = re.compile(r"(\d{1,2})\s*(?:\+|-|–|to)?\s*\d{0,2}\s*(?:years?|yrs?)\s*(?:of\s+)?(?:work\s+)?(?:experience|exp)\b", re.I)
EXP_CARD_RE = re.compile(r"(\d{1,2})\s*(?:-|–|to)\s*\d{1,2}\s*(?:years?|yrs?)\b", re.I)
NCR_RE = re.compile(r"delhi|\bncr\b|noida|gurgaon|gurugram|ghaziabad|faridabad", re.I)
REMOTE_RE = re.compile(r"remote|work from home|\bwfh\b", re.I)
DEADLINE_RE = re.compile(
    r"(?:last date|apply by|apply before|deadline|closing date|applications? clos\w*(?:\s+on)?)\D{0,25}"
    r"(\d{1,2}(?:st|nd|rd|th)?[\s\-/](?:[A-Za-z]{3,9}|\d{1,2})[\s\-/,]+\d{2,4})", re.I)


def parse_date(v):
    """Parse strings, Sheets serial numbers, or datetimes into a datetime (or None)."""
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        if 30000 < v < 80000:
            return datetime(1899, 12, 30) + timedelta(days=float(v))
        return None
    s = re.sub(r"(?<=\d)(st|nd|rd|th)\b", "", str(v).strip(), flags=re.I)
    s = re.sub(r"[\s\-/,]+", " ", s).strip()
    for fmt in ("%Y %m %d", "%d %m %Y", "%d %b %Y", "%d %B %Y", "%d %m %y", "%d %b %y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def normalize_link(link):
    link = (link or "").strip()
    if not link:
        return ""
    p = urlparse(link)
    q = [(k, v) for k, v in parse_qsl(p.query) if not k.lower().startswith(("utm_", "ref", "src"))]
    return urlunparse((p.scheme, p.netloc, p.path.rstrip("/"), "", urlencode(q), ""))


def full_text(j):
    return " ".join([j.get("title", ""), j.get("card_text", ""), j.get("detail_text", "")])


def min_experience(j):
    vals = [int(m) for m in EXP_RE.findall(full_text(j))]
    vals += [int(m) for m in EXP_CARD_RE.findall(j.get("card_text", ""))]
    return min(vals) if vals else None


def classify_domain(title):
    for name, weight, rx in DOMAINS:
        if rx.search(title):
            return name, weight
    return "Other", 0


def title_gate(j):
    """Cheap check before spending a detail request."""
    t = j["title"]
    if NOISE_RE.search(t):
        return False
    return classify_domain(t)[0] != "Other"


def quality_gate(j):
    text = full_text(j)
    if CREDENTIAL_RE.search(text):
        return False
    mn = min_experience(j)
    is_intern = bool(INTERN_RE.search(j["title"]))
    if not is_intern and mn is not None and mn >= 2:
        return False
    if is_intern and mn is not None and mn >= 3:
        return False
    if STRICT_LOCATION and not j.get("ncr_or_remote"):
        return False
    return True


def role_type(j):
    if INTERN_RE.search(j["title"]):
        return "PPO Internship" if PPO_RE.search(full_text(j)) else "Internship"
    if GRAD_RE.search(j["title"] + " " + j.get("card_text", "")[:600] + " " + j.get("detail_text", "")[:600]):
        return "Graduate Program"
    return "Full-time (Fresher)"


def extract_deadline(j):
    m = DEADLINE_RE.search(j.get("detail_text", "") + " " + j.get("card_text", ""))
    if not m:
        return ""
    d = parse_date(m.group(1))
    if not d or d.date() < TODAY_D - timedelta(days=1):
        return ""
    return d.strftime("%Y-%m-%d")


def score_job(j):
    s = j["domain_weight"]
    s += {"PPO Internship": 30, "Graduate Program": 20, "Full-time (Fresher)": 10, "Internship": 8}.get(j["type"], 0)
    if j["tier1"]:
        s += 15
    if j["ncr_or_remote"]:
        s += 10
    if j.get("deadline"):
        d = parse_date(j["deadline"])
        if d and 0 <= (d.date() - TODAY_D).days <= 7:
            s += 5
    return min(s, 100)


def enrich_basic(j):
    j["domain"], j["domain_weight"] = classify_domain(j["title"])
    j["tier1"] = bool(TIER1_RE.search(j.get("company", "")))
    loc_blob = j.get("location") or j.get("card_text", "")[:300]
    j["ncr_or_remote"] = bool(NCR_RE.search(loc_blob) or REMOTE_RE.search(loc_blob) or REMOTE_RE.search(j["title"]))


def finalize(j):
    j["type"] = role_type(j)
    j["ppo"] = j["type"] == "PPO Internship"
    j["deadline"] = extract_deadline(j)
    j["score"] = score_job(j)


# ------------------------------------------------------------------
# HTTP
# ------------------------------------------------------------------
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
           "Accept-Language": "en-IN,en;q=0.9"}


def safe_request(url, retries=2):
    verify = urlparse(url).hostname not in UNVERIFIED_HOSTS
    for i in range(retries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=20, verify=verify)
            if r.status_code == 200:
                return r
            print(f"[WARN] {url} -> {r.status_code}")
        except Exception as e:
            print(f"[ERROR] {url}: {e}")
        if i < retries - 1:
            time.sleep(6)
    return None


def polite_sleep(a=1.5, b=3.5):
    time.sleep(random.uniform(a, b))


# ------------------------------------------------------------------
# SCRAPERS
# ------------------------------------------------------------------
def make_job(title, company, location, link, source, card):
    return {
        "title": title, "company": company or "Unknown", "location": location,
        "link": link, "source": source,
        "card_text": card.get_text(" ", strip=True)[:800],
        "detail_text": "", "found_on": TODAY.strftime("%Y-%m-%d"),
    }


def scrape_shine(slug):
    jobs = []
    r = safe_request(f"https://www.shine.com/job-search/{slug}-jobs")
    if not r:
        return jobs
    soup = BeautifulSoup(r.text, "html.parser")
    cards = soup.select("li.job-listing") or soup.select("div[class*='jobCard']") or soup.select("li[class*='job']")
    for c in cards:
        try:
            t = c.select_one("a[class*='title'], h2, h3")
            a = c.select_one("a[href]")
            if not t or not a:
                continue
            comp = c.select_one("[class*='company'], [class*='employer']")
            loc = c.select_one("[class*='loc']")
            href = a["href"]
            link = href if href.startswith("http") else "https://www.shine.com" + href
            jobs.append(make_job(t.text.strip(), comp.text.strip() if comp else "",
                                 loc.text.strip() if loc else "", link, "Shine", c))
        except Exception as e:
            print(f"[WARN] Shine parse: {e}")
    print(f"[INFO] Shine '{slug}': {len(jobs)}")
    return jobs


def scrape_timesjobs(keyword):
    jobs = []
    r = safe_request(f"https://www.timesjobs.com/candidate/job-search.html?txtKeywords={quote_plus(keyword)}")
    if not r:
        return jobs
    soup = BeautifulSoup(r.text, "html.parser")
    cards = soup.select("li.clearfix.job-bx.wht-shd-bx") or soup.select("li[class*='job-bx']")
    for c in cards:
        try:
            a = c.select_one("h2 a")
            if not a or not a.get("href"):
                continue  # v6 bug: fallback to search URL made distinct jobs collapse
            comp = c.select_one("h3.joblist-comp-name") or c.select_one("[class*='company']")
            loc = c.select_one("[class*='loc']") or c.select_one(".top-jd-dtl li")
            jobs.append(make_job(a.text.strip(), comp.text.strip() if comp else "",
                                 loc.text.strip() if loc else "", a["href"], "TimesJobs", c))
        except Exception as e:
            print(f"[WARN] TimesJobs parse: {e}")
    print(f"[INFO] TimesJobs '{keyword}': {len(jobs)}")
    return jobs


def scrape_all():
    jobs = []
    for slug, kw in QUERIES:
        jobs += scrape_shine(slug)
        polite_sleep()
        jobs += scrape_timesjobs(kw)
        polite_sleep()
    return jobs


def fetch_detail(j):
    r = safe_request(j["link"], retries=1)
    if r:
        text = BeautifulSoup(r.text, "html.parser").get_text(" ", strip=True)
        j["detail_text"] = text[:5000]
    polite_sleep(1.0, 2.5)


# ------------------------------------------------------------------
# GOOGLE SHEETS
# ------------------------------------------------------------------
SEEN_HEADERS = ["link", "first_seen"]
PENDING_HEADERS = ["link", "title", "company", "location", "type", "domain", "score", "deadline", "source", "tier1"]
TRACKER_HEADERS = ["Company", "Role", "Type", "Link", "Score", "Status", "Applied Date", "Deadline",
                   "Follow-up Date", "Contact", "PPO", "Notes", "Added On"]
STATUSES = ["To Apply", "Applied", "Interview", "Offer", "Rejected", "Ghosted", "Withdrawn"]


def get_spreadsheet():
    if not SA_JSON or not GSHEET_ID:
        print("[WARN] Sheets not configured: no cross-run dedup, tracker, or weekly digest storage")
        return None
    try:
        creds = Credentials.from_service_account_info(
            json.loads(SA_JSON), scopes=["https://www.googleapis.com/auth/spreadsheets"])
        return gspread.authorize(creds).open_by_key(GSHEET_ID)
    except Exception as e:
        print(f"[ERROR] Sheets auth/open failed: {e}")
        return None


def get_ws(sh, title, headers):
    """Return (worksheet, created)."""
    try:
        return sh.worksheet(title), False
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=len(headers))
        ws.append_row(headers)
        return ws, True


def load_seen(sh):
    try:
        ws, _ = get_ws(sh, "seen_listings", SEEN_HEADERS)
        return {normalize_link(v) for v in ws.col_values(1)[1:] if v.strip()}
    except Exception as e:
        print(f"[ERROR] load_seen: {e}")
        return set()


def save_seen(sh, jobs):
    if not jobs:
        return
    try:
        ws, _ = get_ws(sh, "seen_listings", SEEN_HEADERS)
        ws.append_rows([[normalize_link(j["link"]), j["found_on"]] for j in jobs], value_input_option="RAW")
    except Exception as e:
        print(f"[ERROR] save_seen: {e}")


def save_pending(sh, jobs):
    if not jobs:
        return
    try:
        ws, _ = get_ws(sh, "pending_fulltime", PENDING_HEADERS)
        ws.append_rows([[j["link"], j["title"], j["company"], j["location"], j["type"], j["domain"],
                         j["score"], j["deadline"], j["source"], "Y" if j["tier1"] else ""] for j in jobs],
                       value_input_option="RAW")
    except Exception as e:
        print(f"[ERROR] save_pending: {e}")


def load_pending(sh):
    try:
        ws, _ = get_ws(sh, "pending_fulltime", PENDING_HEADERS)
        rows = ws.get_all_values()[1:]
        out = []
        for r in rows:
            r = r + [""] * (len(PENDING_HEADERS) - len(r))
            d = dict(zip(PENDING_HEADERS, r))
            if not d["link"]:
                continue
            d["score"] = int(d["score"]) if str(d["score"]).isdigit() else 0
            d["tier1"] = d["tier1"] == "Y"
            d["ppo"] = False
            out.append(d)
        return out
    except Exception as e:
        print(f"[ERROR] load_pending: {e}")
        return []


def clear_pending(sh):
    try:
        ws, _ = get_ws(sh, "pending_fulltime", PENDING_HEADERS)
        ws.clear()
        ws.append_row(PENDING_HEADERS)
    except Exception as e:
        print(f"[ERROR] clear_pending: {e}")


def get_tracker(sh):
    ws, created = get_ws(sh, "tracker", TRACKER_HEADERS)
    if created:
        try:
            sh.batch_update({"requests": [{"setDataValidation": {
                "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 1000,
                          "startColumnIndex": 5, "endColumnIndex": 6},
                "rule": {"condition": {"type": "ONE_OF_LIST",
                                       "values": [{"userEnteredValue": s} for s in STATUSES]},
                         "showCustomUi": True, "strict": False}}}]})
        except Exception as e:
            print(f"[WARN] Could not add Status dropdown: {e}")
    return ws


def read_tracker_rows(ws):
    try:
        values = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
    except TypeError:
        values = ws.get_all_values()
    if not values:
        return []
    headers = values[0]
    return [dict(zip(headers, r + [""] * (len(headers) - len(r)))) for r in values[1:] if any(r)]


def add_to_tracker(sh, jobs):
    """Append top-scoring new listings as 'To Apply'. Returns number added."""
    try:
        ws = get_tracker(sh)
        existing = {normalize_link(str(r.get("Link", ""))) for r in read_tracker_rows(ws)}
        picks = [j for j in sorted(jobs, key=lambda x: -x["score"])
                 if j["score"] >= TRACKER_MIN_SCORE and normalize_link(j["link"]) not in existing][:TRACKER_MAX_ADD]
        if not picks:
            return 0
        ws.append_rows([[j["company"], j["title"], j["type"], j["link"], j["score"], "To Apply", "",
                         j["deadline"], "", "", "Y" if j["ppo"] else "N", "", TODAY.strftime("%Y-%m-%d")]
                        for j in picks], value_input_option="RAW")
        print(f"[INFO] Added {len(picks)} listings to tracker")
        return len(picks)
    except Exception as e:
        print(f"[ERROR] add_to_tracker: {e}")
        return 0


def build_reminders(sh):
    try:
        rows = read_tracker_rows(get_tracker(sh))
    except Exception as e:
        print(f"[ERROR] tracker read: {e}")
        return []
    out = []
    for r in rows:
        status = str(r.get("Status", "")).strip().lower()
        who = f"{r.get('Company', '')} - {r.get('Role', '')}".strip(" -")
        link = str(r.get("Link", ""))
        applied, deadline = parse_date(r.get("Applied Date")), parse_date(r.get("Deadline"))
        follow, added = parse_date(r.get("Follow-up Date")), parse_date(r.get("Added On"))

        def add(level, text):
            out.append({"level": level, "text": text, "link": link})

        if status == "to apply":
            if deadline:
                d = (deadline.date() - TODAY_D).days
                if 0 <= d <= DEADLINE_WARN_DAYS:
                    add(0, f"Deadline in {d} day(s): {who}")
                    continue
            if added and (TODAY_D - added.date()).days >= STALE_TO_APPLY_DAYS:
                add(2, f"In 'To Apply' for {(TODAY_D - added.date()).days} days. Apply or drop: {who}")
        elif status in ("applied", "interview"):
            if follow and follow.date() <= TODAY_D:
                add(1, f"Follow-up due: {who}")
            elif not follow and applied:
                age = (TODAY_D - applied.date()).days
                if age >= GHOSTED_AFTER_DAYS:
                    add(2, f"No reply in {age} days. Send a last nudge or mark Ghosted: {who}")
                elif age >= FOLLOWUP_AFTER_DAYS:
                    add(1, f"Follow up, applied {age} days ago: {who}")
    return sorted(out, key=lambda x: x["level"])


# ------------------------------------------------------------------
# EMAIL
# ------------------------------------------------------------------
def esc(x):
    return html.escape(str(x), quote=True)


def send_mail(subject, body):
    if not all([EMAIL, EMAIL_PASS, TO_EMAIL]):
        print("[ERROR] Email credentials missing; skipping send")
        return
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"], msg["From"], msg["To"] = subject, EMAIL, TO_EMAIL
        msg.attach(MIMEText(body, "html"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL, EMAIL_PASS)
            server.sendmail(EMAIL, TO_EMAIL, msg.as_string())
        print(f"[INFO] Email sent: {subject}")
    except Exception as e:
        print(f"[ERROR] Email failed: {e}")


def job_html(j):
    badges = []
    if j.get("ppo"):
        badges.append("<span style='background:#0a7d34;color:#fff;padding:1px 6px;border-radius:4px'>PPO</span>")
    if j.get("tier1"):
        badges.append("<span style='background:#b8860b;color:#fff;padding:1px 6px;border-radius:4px'>Tier 1</span>")
    badges.append(f"<span style='color:#555'>{esc(j.get('type', ''))} | {esc(j.get('domain', ''))}</span>")
    dl = f"<br>Deadline: <b>{esc(j['deadline'])}</b>" if j.get("deadline") else ""
    return (f"<p style='margin:0 0 14px'><b>{esc(j['title'])}</b> ({esc(j.get('score', 0))})<br>"
            f"{esc(j['company'])} | {esc(j.get('location') or 'India')}<br>{' '.join(badges)}{dl}<br>"
            f"<a href=\"{esc(j['link'])}\" style='background:#0073e6;color:#fff;padding:5px 12px;"
            f"border-radius:6px;text-decoration:none;display:inline-block;margin-top:4px'>Apply</a></p>")


def reminders_html(reminders):
    if not reminders:
        return ""
    icons = {0: "&#128680;", 1: "&#9200;", 2: "&#128204;"}
    items = "".join(
        f"<li>{icons.get(r['level'], '')} "
        + (f"<a href=\"{esc(r['link'])}\">{esc(r['text'])}</a>" if r["link"].startswith("http") else esc(r["text"]))
        + "</li>" for r in reminders)
    return f"<h2>Action items ({len(reminders)})</h2><ul>{items}</ul><hr>"


def send_daily(internships, reminders, tracker_added):
    ppo = sum(1 for j in internships if j["ppo"])
    body = reminders_html(reminders)
    if internships:
        body += f"<h2>{len(internships)} new internships ({ppo} PPO)</h2>"
        grouped = defaultdict(list)
        for j in internships:
            grouped[j["domain"]].append(j)
        for d, items in sorted(grouped.items(), key=lambda kv: -max(i["score"] for i in kv[1])):
            body += f"<h3>{esc(d)}</h3>" + "".join(job_html(j) for j in sorted(items, key=lambda x: -x["score"]))
    if tracker_added:
        body += f"<p><i>{tracker_added} top listings were added to your tracker tab as 'To Apply'.</i></p>"
    parts = []
    if internships:
        parts.append(f"{len(internships)} new ({ppo} PPO)")
    if reminders:
        parts.append(f"{len(reminders)} action items")
    send_mail(f"[Internship Bot] {' | '.join(parts)} - {TODAY.strftime('%d %b')}", body)


def send_weekly_fulltime(jobs):
    jobs = sorted(jobs, key=lambda x: -x["score"])
    body = f"<h2>Weekly full-time digest: {len(jobs)} roles</h2>" + "".join(job_html(j) for j in jobs)
    send_mail(f"[Internship Bot] Weekly full-time digest: {len(jobs)} roles - {TODAY.strftime('%d %b')}", body)


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def main():
    sh = get_spreadsheet()
    reminders = build_reminders(sh) if sh else []

    jobs = scrape_all()
    print(f"[INFO] Total scraped: {len(jobs)}")

    if not jobs:
        print("[WARN] All scrapers returned 0. Selectors likely broke or the runner is blocked.")
        send_mail("[Internship Bot] ALERT: scrapers returned 0 results",
                  reminders_html(reminders) +
                  "<p>Every scraper returned zero listings. Check the Actions log; "
                  "site markup may have changed or the runner IP is blocked.</p>")
        return

    # in-run dedup + exclusions
    seen_run, unique = set(), []
    for j in jobs:
        n = normalize_link(j["link"])
        if n and n not in seen_run and n not in EXCLUDED_LINKS:
            seen_run.add(n)
            unique.append(j)
    print(f"[INFO] After dedup: {len(unique)}")

    for j in unique:
        enrich_basic(j)
    candidates = [j for j in unique if title_gate(j)]
    print(f"[INFO] After title gate: {len(candidates)}")

    seen = load_seen(sh) if sh else set()
    new = [j for j in candidates if normalize_link(j["link"]) not in seen]
    new.sort(key=lambda j: -(j["domain_weight"] + (15 if j["tier1"] else 0) + (10 if j["ncr_or_remote"] else 0)))
    to_process = new[:MAX_DETAIL_FETCH]
    print(f"[INFO] New: {len(new)}; fetching details for {len(to_process)}")

    for j in to_process:
        fetch_detail(j)

    kept = []
    for j in to_process:
        if quality_gate(j):
            finalize(j)
            kept.append(j)
    print(f"[INFO] Kept after quality gate: {len(kept)}")

    internships = sorted([j for j in kept if "Internship" in j["type"]], key=lambda x: -x["score"])
    fulltime = sorted([j for j in kept if "Internship" not in j["type"]], key=lambda x: -x["score"])

    # persist state
    if sh:
        save_seen(sh, to_process)          # includes rejected ones so they are not re-fetched
        save_pending(sh, fulltime)
        tracker_added = add_to_tracker(sh, kept)
    else:
        tracker_added = 0

    # Excel
    cols = ["score", "type", "domain", "title", "company", "location", "deadline", "tier1", "ppo", "source", "link", "found_on"]
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        for name, subset in (("Internships", internships), ("Full-time", fulltime)):
            pd.DataFrame(subset, columns=cols).to_excel(w, sheet_name=name, index=False)

    # emails
    if internships or reminders:
        send_daily(internships, reminders, tracker_added)
    else:
        print("[INFO] Nothing new and no reminders; no daily email")

    if FORCE_FULLTIME_DIGEST or TODAY.weekday() == 0:
        due = load_pending(sh) if sh else fulltime
        if due:
            send_weekly_fulltime(due)
            if sh:
                clear_pending(sh)
        else:
            print("[INFO] Weekly digest due but no full-time roles pending")


if __name__ == "__main__":
    main()
