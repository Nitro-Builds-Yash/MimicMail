<div align="center">

# ✉️ MimicMail — Humanized Outreach Engine

**A clinical-grade, stealth outreach automation platform for Zoho Mail with Playwright and modern Tkinter UI.**

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Playwright](https://img.shields.io/badge/engine-Playwright-2EAD33.svg)](https://playwright.dev/)
[![Tkinter Modern GUI](https://img.shields.io/badge/GUI-Tkinter-1A62C6.svg)](https://docs.python.org/3/library/tkinter.html)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

<p align="center">
  <img src="assets/pulsus_logo.png" alt="Pulsus Group Official Logo" width="280"/>
</p>

</div>

---

## 📌 Overview

**MimicMail** is engineered for high-deliverability clinical and enterprise outreach. Rather than relying on rigid SMTP connections that frequently trigger spam and bulk-sender penalties, MimicMail automates genuine browser interactions through **Playwright** directly inside Zoho Mail's rich webmail interface.

It features persistent browser sessions, anti-bot behavioral jitter, randomized HTML template rotation across 5 customizable slots, real-time visual deliverability analytics, and automatic spreadsheet status synchronization.

---

## ✨ Key Features

- **🌐 Single-Browser Persistent Session:** Authenticate once in your standard browser; the background automation engine attaches directly to the authenticated session without credential leakage or repeated MFA prompts.
- **🎲 5-Slot HTML Template Pool:** Define up to 5 custom HTML templates. MimicMail randomly picks one template per recipient to avoid content fingerprinting by spam filters.
- **✒️ Standardized Typography:** Built-in Times New Roman formatting with zero leading whitespace, and automatic **bolding** for `{name}` and `{title}` variables.
- **⏱️ Anti-Bot Behavioral Jitter:** Configurable resting cadence (1.0 to 5.0 minutes) with organic humanized delay jitter ($\pm15-35\text{s}$) to safeguard sender reputation.
- **📊 Real-time Mission Control:**
  - Dynamic Determinate Progress Bar (0% to 100%).
  - Live Ratio Breakdown Canvas (Visual color segments for Sent, Pending, and Errors).
  - Countdown timer display and live terminal audit logging.
- **🛡️ Fault-Tolerant Excel Tracking:** Non-destructive spreadsheet updating; saves progress in real-time as `Status = SENT` alongside timestamps (`Sent_At`) directly onto disk.

---

## 📂 Repository Structure

```text
MimicMail/
├── assets/                          # Application assets & default templates
│   ├── pulsus_logo.png              # Official branding logo
│   └── templates/                   # Reusable HTML template drafts
│       ├── template_1_editorial.html
│       ├── template_2_speaker.html
│       ├── template_3_advisory.html
│       ├── template_4_trials.html
│       └── template_5_outreach.html
├── data/                            # Spreadsheet workspaces
│   ├── input/                       # Registry files to process
│   │   ├── .gitkeep
│   │   └── recipients.example.csv   # Reference schema format
│   └── processed/                   # Output storage for completed runs
│       └── .gitkeep
├── .env.example                     # Environment configuration reference
├── .gitignore                       # Git ignore rules for virtual environments & sessions
├── README.md                        # Project documentation
├── requirements.txt                 # Python package dependencies
└── zoho_playwright_mailer.py        # Main application entry point & GUI engine
```

---

## 🚀 Quick Start Guide

### 1. Prerequisites
Ensure you have Python 3.10 or higher installed.

```bash
python --version
```

### 2. Clone Repository
```bash
git clone https://github.com/Nitro-Builds-Yash/MimicMail.git
cd MimicMail
```

### 3. Setup Virtual Environment & Install Dependencies
```bash
# Create virtual environment
python -m venv .venv

# Activate virtual environment
# Windows (PowerShell):
.venv\Scripts\Activate.ps1
# Linux / macOS:
source .venv/bin/activate

# Install required dependencies
pip install -r requirements.txt

# Install Playwright Chromium browser binary
playwright install chromium
```

---

## 🖥️ Usage Workflow

### 1. Start the Application
Run the launcher:
```bash
python zoho_playwright_mailer.py
```

### 2. Connect Zoho Webmail
1. Click **`🌐 Launch Zoho Webmail`** at the top-left of the Session card.
2. An automated Chromium browser window will open.
3. Sign into your Zoho account (complete any 2FA/MFA if enabled).
4. Keep this browser window open; the status indicator in the UI will update to **`Connected`**.

### 3. Ingest Recipient Registry
1. Prepare an Excel spreadsheet (`.xlsx`) containing at least the following column headers:
   - `name` — Full recipient name.
   - `email` — Valid recipient email address.
   - `title` — Recipient role, conference title, or publication subject.
2. Click **`📁 Browse Excel (.xlsx)`** at the top-right of Card 1 and choose your file.

### 4. Configure Subject & HTML Templates
1. Set your subject line format in the Subject entry (default: `{title}`).
2. In Card 2, inspect or edit the **5 HTML Template Slots**. You can load `.html` files using the **`📂 Load File`** button on each tab or paste HTML directly.
3. The engine will pick a random template from the active slots for each contact.

### 5. Launch Automated Outreach
1. Adjust the **Rest Interval** slider (default: `2.0 mins`) and toggle **Human Behavioral Jitter**.
2. Click **`🚀 START AUTOMATION`**.
3. Watch live metrics, the ratio breakdown graph, and terminal logs update in real time.
4. If needed, click **`🛑 EMERGENCY STOP`** to safely pause after the current dispatch finishes.

---

## 📋 Registry Column Schema

| Column Name | Required | Description | Example |
|:---|:---:|:---|:---|
| `name` | **Yes** | Full name of recipient | `Dr. Sarah Jenkins` |
| `email` | **Yes** | Destination email address | `s.jenkins@hospital.org` |
| `title` | **Yes** | Dynamic title / designation used in subject and body | `Chief of Oncology` |
| `Status` | *Auto* | Populated by engine upon completion | `SENT` |
| `Sent_At` | *Auto* | Timestamp recorded when sent | `2026-10-07 14:32:05` |

---

## 🔒 Security & Privacy

- **No Stored Passwords:** MimicMail does not capture or store Zoho credentials in plain text or configuration files. Authentication happens natively in the Chromium session.
- **Ignored Session Cache:** The `.gitignore` is pre-configured to strictly ignore `zoho_browser_profile/`, local spreadsheet logs, and `.env` files to prevent accidental leakage to version control.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
