# Project structure

```text
zoho-mail-automation/
├── config/                     # Reserved for non-secret configuration
├── data/
│   ├── input/                  # Example/source workbooks (keep personal data private)
│   └── processed/              # Optional completed workbook copies
├── docs/                       # Project notes and decisions
├── logs/                       # Reserved for local app logs
├── src/zoho_mail_automation/
│   ├── app.py                  # Tkinter drag/drop review interface
│   └── data/                   # name/email/title validation and sent-state persistence
└── tests/                      # Reserved for future tests
```

Run from the repository root after installation with `zoho-mail-merge`.
