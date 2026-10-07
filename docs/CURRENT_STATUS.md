# Current status

Updated: 2026-10-07

## Project state

The desktop interface has a styled header and review card, Excel/CSV drop area with browse fallback, optional Zoho login ID, and a visible progress bar. The subject comes from each row's `title`; the body template supports `{name}`, `{email}`, and `{title}`. Zoho Mail opens in the default browser. Progress advances only when the user confirms a message was manually sent. Workbook headers must include lowercase `name`, `email`, and `title`; `Status` and `Sent_At` are added if absent. `SENT` rows are skipped. The 1–5 minute pacing control disables the app's Open/Copy actions between confirmations, but cannot block direct browser actions.

The app deliberately does not accept a Zoho password, automate browser login, or click Send. The user logs in and sends each message in the visible Zoho page. No SMTP dispatch code remains active in the application.

## Verification and limits

- Source syntax compilation was performed after edits.
- The application has not been launched in a desktop session or tested with Zoho.
- Runtime packages (`openpyxl`, `tkinterdnd2`) must be installed with `python -m pip install -e .`.
- No automated tests were added or run.
