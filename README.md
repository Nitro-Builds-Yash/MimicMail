# Zoho Mail Merge — Review and Send

A desktop helper for preparing personalized messages from a local Excel or CSV workbook. It opens Zoho Mail in the default browser, previews one recipient at a time, copies the message for review, and records rows that the user confirms were manually sent.

The app does not collect or store a Zoho password, log in to Zoho, click Send, or dispatch messages through SMTP. After a workbook loads it opens Zoho Mail in the default browser; the user signs in and sends each message in the visible Zoho page. The selected 1–5 minute interval locks the app's Open/Copy controls between confirmed sends. It cannot prevent a user from sending directly in the browser during that interval. Use only with recipients who agreed to receive the messages and follow Zoho's current policies.

## Run

1. Install Python 3.10 or newer.
2. From this folder, run `python -m pip install -e .`.
3. Run `zoho-mail-merge` or `python -m zoho_mail_automation.app`.
4. Drop an `.xlsx` or `.csv` file in the drop area (or browse). The first worksheet is used for Excel files.
5. Enter a Zoho login ID if desired and choose a 1–5 minute interval. The app opens Zoho Mail; enter your password on Zoho's page.
6. Review the prepared message, copy it to Zoho Mail, and send it yourself. Confirm **Mark sent & next** to update the workbook and progress bar.

Required workbook headers are exactly `name`, `email`, and `title` (lowercase). `Status` and `Sent_At` tracking columns are added when missing. Rows already marked `SENT` and invalid/incomplete rows are skipped. The `title` value is used as each recipient's subject; use **Copy subject** to copy it separately, and **Copy message** to copy the recipient, subject, and body preview. Edit the body template in the app; supported placeholders are `{name}`, `{email}`, and `{title}`. Progress counts rows you confirm were sent; it does not observe Zoho delivery.

## Privacy

The app never asks for your Zoho password. Enter credentials only on Zoho's own page. Recipient data remains in the selected workbook. Do not commit recipient files or credentials.

See [`docs/PROJECT_STRUCTURE.md`](docs/PROJECT_STRUCTURE.md), [`docs/CURRENT_STATUS.md`](docs/CURRENT_STATUS.md), and [`docs/DECISIONS.md`](docs/DECISIONS.md).
