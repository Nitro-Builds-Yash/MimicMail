# ⚕️ Pulsus MedScout — MimicMail 🩺

A clinical-grade, humanized outreach automation engine built for Zoho Mail with Playwright and a dark medical interface.

---

## 🔬 Features
- **🌐 Single-Browser Persistent Session:** Authenticate once in your standard browser; the background queue operates directly in that tab without repeated logins.
- **⚕️ Clinical Dark Mode Interface:** Professional dark UI with clinical symbols (`⚕️`, `🩺`, `🔬`, `🧬`, `🏥`), live countdown pacing, and console audit logs.
- **📝 Times New Roman HTML Rendering:** Automatic formatting in Times New Roman (12pt) with zero leading spaces before `Dear {name},`.
- **📌 Dynamic Subject Resolution:** Automatically sets each recipient's subject line to their exact `title` from Excel.
- **🛡️ Resilient Batch Validation:** Automatically detects and skips invalid email addresses without crashing the queue.
- **⏱️ Anti-Bot Behavioral Jitter:** Staggers dispatches with 1–5 minute configurable intervals + randomized jitter to maintain high deliverability and account safety.

---

## 🚀 Getting Started

### 1. Requirements
- Python 3.10+
- Install dependencies:
```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Launch
```bash
python zoho_playwright_mailer.py
```

### 3. Usage Flow
1. Click **`🌐 Launch Webmail & Sign In`** $\rightarrow$ log in to Zoho Mail in the browser window (leave the browser open).
2. Click **`📁 Browse Registry (.xlsx)`** and select your Excel spreadsheet (`name`, `email`, `title`).
3. Click **`🚀 START AUTOMATED OUTREACH`** $\rightarrow$ dispatches in the background while updating rows to `Status = SENT` on disk.
