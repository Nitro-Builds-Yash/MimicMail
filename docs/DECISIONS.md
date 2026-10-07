# Decisions

Updated: 2026-10-07

1. **Visible, user-controlled sending:** Open Zoho Mail in the default browser; the operator logs in, reviews, and clicks Send manually. Do not automate browser login or sending.
2. **No password collection:** The desktop tool only accepts an optional login ID for convenience. Password entry belongs on Zoho's own sign-in page.
3. **Pacing:** Provide a 1–5 minute interval and temporarily disable the app's Open/Copy actions between confirmed sends. The app cannot block direct browser actions; do not randomize timing.
4. **Transparent workflow:** Do not spoof desktop-client headers or conceal automation. No SMTP transport is used by this UI.
5. **Resume state:** After the user confirms a send, persist `SENT` and a local-time ISO timestamp; skip those rows on the next load.
6. **Workbook schema:** Support `.xlsx` and `.csv`, require exact lowercase `name`, `email`, and `title` headers, and add `Status`/`Sent_At` tracking columns when absent.
7. **Drag/drop fallback:** Use tkinterdnd2 for drop support and retain a normal file picker for environments where drag/drop support is unavailable.
8. **Workbook title as subject:** Use each recipient's required `title` value as the message subject. Show it in the preview and provide a separate action to copy it into Zoho's subject field.
