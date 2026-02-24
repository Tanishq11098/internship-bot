# Google Sheets Setup Guide
# Internship Bot v5.0 — Deduplication Feature

This guide walks you through creating a Google Service Account so the bot
can read/write to Google Sheets for cross-run deduplication.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 1 — Create a Google Cloud Project
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Go to https://console.cloud.google.com
2. Click "Select a project" → "New Project"
3. Name it: internship-bot  (or anything)
4. Click "Create"


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 2 — Enable Google Sheets API
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. In your project, go to:
   APIs & Services → Library
2. Search: "Google Sheets API"
3. Click it → Click "Enable"


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 3 — Create a Service Account
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Go to: APIs & Services → Credentials
2. Click "+ Create Credentials" → "Service account"
3. Fill in:
     Service account name:  internship-bot-sa
     Service account ID:    internship-bot-sa   (auto-filled)
4. Click "Create and Continue"
5. Role: skip (click "Continue")
6. Click "Done"


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 4 — Download the JSON Key
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Click your newly created service account
2. Go to the "Keys" tab
3. Click "Add Key" → "Create new key"
4. Select: JSON → Click "Create"
5. A file like internship-bot-sa-xxxx.json will download
6. KEEP THIS FILE SAFE — treat it like a password


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 5 — Create the Google Sheet
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Go to https://sheets.google.com
2. Create a new blank spreadsheet
3. Name it: Internship Bot — Seen Listings
4. Copy the Spreadsheet ID from the URL:
   https://docs.google.com/spreadsheets/d/  <<<THIS_PART>>>  /edit
   Example ID: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 6 — Share the Sheet with the Service Account
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Open your Google Sheet
2. Click "Share" (top right)
3. In the "Add people" field, paste the service account email.
   It looks like:
     internship-bot-sa@internship-bot-xxxxx.iam.gserviceaccount.com
   (Find it in the JSON file under "client_email")
4. Set permission: Editor
5. Uncheck "Notify people"
6. Click "Share"


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 7 — Add GitHub Secrets
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Go to your GitHub repo → Settings → Secrets and variables → Actions
Add the following secrets:

SECRET NAME                   VALUE
─────────────────────────────────────────────────────────────────────
EMAIL                         your Gmail address
PASSWORD                      your Gmail App Password (not main password)
ANTHROPIC_API_KEY             your Anthropic API key (from console.anthropic.com)
GSHEET_ID                     the Spreadsheet ID from Step 5
GOOGLE_SERVICE_ACCOUNT_JSON   the ENTIRE contents of the JSON file from Step 4
                              (open it in Notepad, select all, paste)

How to get a Gmail App Password:
  1. Go to myaccount.google.com/security
  2. Enable 2-Step Verification (if not already)
  3. Search "App Passwords" → Create one for "Mail"
  4. Use that 16-character password as your PASSWORD secret

How to get an Anthropic API Key:
  1. Go to https://console.anthropic.com
  2. API Keys → Create Key
  3. Copy and save it immediately (shown only once)


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 8 — Deploy Your Files
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Put files in your repo like this:

  your-repo/
  ├── internship_bot_v5.py
  └── .github/
      └── workflows/
          └── internship_bot_v5.yml

Commit and push. The bot will run automatically on:
  Mon / Thu / Sat / Sun at 8:00 AM IST

To test immediately:
  GitHub → Actions tab → "Internship Bot v5.0" → "Run workflow"


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
HOW DEDUPLICATION WORKS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
First run:
  → All listings scraped → All saved to Google Sheet tab "seen_listings"
  → Email subject says "X NEW Internships"

Next runs:
  → Bot loads seen_listings from Google Sheet
  → Filters out anything already seen
  → Only NEW listings count in "NEW TODAY" email subject
  → All listings (old + new) still appear in Excel for reference
  → New keys appended to Google Sheet for future dedup

The Google Sheet will grow over time and acts as your permanent
listing memory across all runs.


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TROUBLESHOOTING
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Error: "PERMISSION_DENIED"
  → You forgot to share the Google Sheet with the service account email (Step 6)

Error: "invalid_grant"
  → The JSON key file is malformed or expired. Re-download from Step 4.

Error: "GOOGLE_SERVICE_ACCOUNT_JSON not set"
  → The GitHub secret name must match exactly (case-sensitive)

Error: "anthropic.AuthenticationError"
  → Check your ANTHROPIC_API_KEY secret is correct

Bot runs but no email received:
  → Check Gmail "Less secure apps" or use App Password
  → Check spam folder
