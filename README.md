# Internship Bot 🚀

A proprietary Python automation system designed to streamline internship
tracking and management using Google Sheets integration.

------------------------------------------------------------------------

## 📌 Overview

Internship Bot is an automation tool built to:

-   Track internship-related data
-   Automatically update Google Sheets
-   Reduce manual tracking effort
-   Streamline internship workflow processes

This project is built using Python and integrates with the Google Sheets
API for structured data management.

------------------------------------------------------------------------

## 🗂 Project Structure

internship-bot/ │ ├── internship_bot_v5.py \# Main automation script ├──
GOOGLE_SHEETS_SETUP.md \# Google Sheets API setup guide └── README.md \#
Project documentation

------------------------------------------------------------------------

## ⚙️ Requirements

Before running the bot, ensure you have:

-   Python 3.7+
-   Google Cloud Project with Sheets API enabled
-   Service Account credentials (JSON file)
-   Required Python libraries

------------------------------------------------------------------------

## 📦 Installation

### 1️⃣ Clone the Repository

git clone https://github.com/Tanishq11098/internship-bot.git\
cd internship-bot

### 2️⃣ Install Dependencies

If you create a requirements.txt:

pip install -r requirements.txt

Example manual installation:

pip install gspread oauth2client pandas

------------------------------------------------------------------------

## 🔐 Google Sheets Setup

Follow the instructions inside:

GOOGLE_SHEETS_SETUP.md

This includes:

-   Enabling Google Sheets API
-   Creating a Service Account
-   Downloading credentials JSON
-   Sharing your sheet with the service account email

------------------------------------------------------------------------

## ▶️ Running the Bot

python internship_bot_v5.py

Make sure:

-   Your credentials file path is correct
-   Your Google Sheet is properly configured
-   Required permissions are granted

------------------------------------------------------------------------

## 🛡 Security Notice

Do NOT upload: - credentials.json - .env - API keys

Add them to .gitignore

Example:

credentials.json\
.env

------------------------------------------------------------------------

## ❗ License & Usage

© 2026 Tanishq Singhal\
All Rights Reserved.

This software is proprietary and confidential.\
Unauthorized copying, distribution, modification, or commercial usage is
strictly prohibited without written permission from the author.

------------------------------------------------------------------------

## 📬 Contact

GitHub: https://github.com/Tanishq11098

------------------------------------------------------------------------

## ⭐ Future Improvements

-   Add logging system
-   Add automated scheduling
-   Improve error handling
-   Add notification system
-   Deploy as cloud-based automation

------------------------------------------------------------------------

Built with focus, automation, and ownership.
