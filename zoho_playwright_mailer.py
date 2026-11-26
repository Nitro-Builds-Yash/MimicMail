import os
import sys
import time
import random
import queue
import threading
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont
import pandas as pd
from playwright.sync_api import sync_playwright
from PIL import Image, ImageTk

STEALTH_INIT_SCRIPT = """
(() => {
    try {
        // Mask navigator.webdriver cleanly without tampering with browser internals
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined,
            configurable: true
        });
    } catch(e) {}
})();
"""

DEFAULT_SUBJECT = "{title}"

_EXCEL_SYNC_LOCK = threading.Lock()
_CACHED_WORKBOOKS = {}

def mark_excel_row_status(excel_path: str, row_idx: int, status_str: str = "SENT"):
    """Thread-safe update of the 'Status' column in the Excel workbook without destroying formulas or formatting.
    Keeps openpyxl workbook cached in memory during active dispatch sessions to eliminate multi-second disk latency."""
    with _EXCEL_SYNC_LOCK:
        try:
            import openpyxl
            wb = _CACHED_WORKBOOKS.get(excel_path)
            if wb is None:
                wb = openpyxl.load_workbook(excel_path)
                _CACHED_WORKBOOKS[excel_path] = wb

            sheet = wb.active
            headers = [c.value for c in sheet[1]]
            status_col = None
            for ci, h in enumerate(headers, start=1):
                if h and str(h).strip().lower() == "status":
                    status_col = ci
                    break
            if status_col is None:
                status_col = len(headers) + 1
                sheet.cell(row=1, column=status_col, value="Status")

            sheet.cell(row=row_idx + 2, column=status_col, value=status_str)
            wb.save(excel_path)
        except Exception as e:
            # If cached workbook encountered an I/O sync error, evict cache and retry once freshly
            try:
                if excel_path in _CACHED_WORKBOOKS:
                    del _CACHED_WORKBOOKS[excel_path]
                import openpyxl
                fresh_wb = openpyxl.load_workbook(excel_path)
                fresh_sheet = fresh_wb.active
                headers = [c.value for c in fresh_sheet[1]]
                status_col = next((ci for ci, h in enumerate(headers, start=1) if h and str(h).strip().lower() == "status"), len(headers) + 1)
                fresh_sheet.cell(row=1, column=status_col, value="Status")
                fresh_sheet.cell(row=row_idx + 2, column=status_col, value=status_str)
                fresh_wb.save(excel_path)
                _CACHED_WORKBOOKS[excel_path] = fresh_wb
            except Exception as e2:
                print(f"Excel sync notice for row {row_idx + 1}: {e2}")

def close_cached_excel_workbook(excel_path: str):
    """Closes any open openpyxl workbook cache cleanly."""
    with _EXCEL_SYNC_LOCK:
        wb = _CACHED_WORKBOOKS.pop(excel_path, None)
        if wb:
            try:
                wb.close()
            except Exception:
                pass

if getattr(sys, 'frozen', False):
    APP_BASE_DIR = os.path.dirname(sys.executable)
    # When PyInstaller unpacks bundle in _MEIPASS
    BUNDLE_RESOURCE_DIR = getattr(sys, '_MEIPASS', APP_BASE_DIR)
else:
    APP_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_RESOURCE_DIR = APP_BASE_DIR

PAYMENT_QR_PATH = os.path.join(BUNDLE_RESOURCE_DIR, "assets", "payment_qr.jpg")
if not os.path.exists(PAYMENT_QR_PATH):
    PAYMENT_QR_PATH = os.path.join(APP_BASE_DIR, "assets", "payment_qr.jpg")

# Default user-customizable templates for the 3 slots (Clean, professional text format matching exact spacing and alignment)
SAMPLE_HTML_TEMPLATES = [
    """Dear Dr. {name},

Greetings from our research team.

We recently reviewed your contribution to {title} and appreciate its scientific contribution to the field of Applied Physics. In recognition of your scholarly contributions, we would like to invite you to submit your forthcoming original research or review article to the Journal of Science and Mathematics Letters.

We welcome innovative research covering Advanced physical research, technological developments, and interdisciplinary scientific applications. We would be delighted to consider your work for publication and welcome your contribution to our journal.

If you are interested, kindly reply to [editor@eclinicalsciresearch.info](mailto:editor@eclinicalsciresearch.info). Our team would be pleased to assist you with the further process. Thank you for your time and consideration. We look forward to the possibility of receiving your valuable contribution.

With regards,
Williams M
Journal of Science and Mathematics Letters | 2462-2052""",

    """Dear Dr. [Author's Name],

Greetings from [Journal Name]!

On behalf of our editorial committee, we cordially invite you to contribute your esteemed research to our upcoming edition in [Journal Name]. Following your groundbreaking work presented in [Paper Title], we believe your insights would greatly benefit our international readership.

By publishing with us, your findings will reach a global audience of specialists dedicated to accelerating clinical breakthroughs. We would be deeply honored to include your submission in this milestone release.

Best regards,
[Journal Name] Editorial Board
MimicMail Scientific Council""",

    """Dear Dr. [Author's Name],

Greetings from [Journal Name]!

We are writing to warmly invite your valuable submission to our specialized issue in [Journal Name]. Your distinguished record and recent work on [Paper Title] demonstrate the significant impact of your research within the scientific community.

This upcoming issue provides an esteemed global platform to highlight your latest discoveries and engage with top researchers across the domain. We sincerely anticipate your participation and look forward to collaborating with you.

Best regards,
[Journal Name] Editorial Board
Global Oncology Review"""
]

def clean_to_plain_text(s: str) -> str:
    """Strips any raw HTML tags, normalizes whitespace and line breaks, and returns clean email text."""
    if not s:
        return ""
    import re
    from html import unescape
    t = re.sub(r'(?i)<br\s*/?>', '\n', s)
    t = re.sub(r'(?i)</(?:p|div|tr|li|h[1-6])>', '\n\n', t)
    t = re.sub(r'(?is)<(?:style|script).*?>.*?</(?:style|script)>', '', t)
    t = re.sub(r'<[^<]+?>', '', t)
    t = unescape(t)
    # Strip trailing/leading whitespace per line, collapse internal multi-spaces
    lines = [re.sub(r'[ \t]+', ' ', line).strip() for line in t.splitlines()]
    # Normalize multiple blank lines to double newlines
    return re.sub(r'\n{3,}', '\n\n', '\n'.join(lines)).strip()

def build_formatted_email_content(raw_template: str, r_name: str, r_title: str, r_email: str, r_journal: str = "", include_optout: bool = True):
    """
    Builds both:
    1. Rich HTML strictly rendered in Times New Roman (12pt), perfectly left-aligned,
       with auto-readjusted paragraph spacing, with '{name}', '{title}', '{journal}', 'Dear Dr' / salutations,
       and 'Best regards' / sign-offs bolded.
       Also appends a subtle, professional opt-out footer to maintain 100% spam compliance.
    2. Clean plain text fallback.
    """
    import re
    from html import escape

    clean_text = clean_to_plain_text(raw_template)

    # 1. Substitute placeholders into plain text
    name_placeholders = [
        "[Author's Name]", "[Author Name]", "[Author’s Name]", "[author's name]", "[author name]",
        "[Name]", "[name]", "[NAME]",
        "{name}", "{NAME}", "{author name}", "{author_name}", "{author}", "{AUTHOR}"
    ]
    title_placeholders = [
        "[Paper Title]", "[paper title]", "[Paper title]", "[PAPER TITLE]", "{paper_title}", "{paper title}",
        "[Title]", "[title]", "[TITLE]", "{title}", "{TITLE}",
        "[Article Title]", "[article title]", "[Article title]", "{article_title}", "{article title}",
        "[Paper]", "[paper]", "{paper}",
        "[Article]", "[article]", "{article}",
        "[Manuscript Title]", "[manuscript title]", "{manuscript_title}",
        "[Publication Title]", "{publication_title}",
        "[Designation]", "[designation]", "{designation}", "{role}"
    ]
    journal_placeholders = [
        "[Journal Name]", "[Journal name]", "[journal name]", "[JOURNAL NAME]",
        "[Journal]", "[journal]", "[JOURNAL]",
        "{journal}", "{JOURNAL}", "{journal_name}", "{journal name}", "{JOURNAL_NAME}",
        "[Publication]", "[publication]", "{publication}"
    ]
    email_placeholders = [
        "[Email]", "[email]", "[EMAIL]", "{email}", "{EMAIL}"
    ]

    plain_text = clean_text
    # 1. Clean markdown links for plain text fallback
    clean_text = re.sub(r'\[([^\]]+)\]\((?:mailto:[^\)]+|https?://[^\)]+)\)', r'\1', clean_text)

    plain_text = clean_text
    for ph in name_placeholders:
        plain_text = plain_text.replace(ph, r_name)
    for ph in title_placeholders:
        plain_text = plain_text.replace(ph, r_title)
    if r_journal:
        for ph in journal_placeholders:
            plain_text = plain_text.replace(ph, r_journal)
    for ph in email_placeholders:
        plain_text = plain_text.replace(ph, r_email)

    bold_style = "font-family: 'Times New Roman', Times, serif; font-weight: bold; font-size: 16px; color: #000000;"
    bold_name = f"<b style=\"{bold_style}\"><font face=\"Times New Roman, Times, serif\" color=\"#000000\">{escape(r_name)}</font></b>"
    bold_title = f"<b style=\"{bold_style}\"><font face=\"Times New Roman, Times, serif\" color=\"#000000\">{escape(r_title)}</font></b>"
    bold_journal = f"<b style=\"{bold_style}\"><font face=\"Times New Roman, Times, serif\" color=\"#000000\">{escape(r_journal)}</font></b>" if r_journal else ""

    # Replace markdown links before escaping other text
    html_work = raw_template
    link_map = {}
    def save_link(m):
        key = f"__LINK_HOLDER_{len(link_map)}__"
        link_map[key] = f'<a href="{m.group(2)}" style="color: #000000; text-decoration: underline;"><font face="Times New Roman, Times, serif" color="#000000">{escape(m.group(1))}</font></a>'
        return key

    html_work = re.sub(r'\[([^\]]+)\]\((mailto:[^\)]+|https?://[^\)]+)\)', save_link, html_work)
    escaped_template = escape(html_work)
    for key, repl in link_map.items():
        escaped_template = escaped_template.replace(escape(key), repl)

    html_content = escaped_template
    for ph in name_placeholders:
        html_content = html_content.replace(escape(ph), bold_name)
    for ph in title_placeholders:
        html_content = html_content.replace(escape(ph), bold_title)
    if bold_journal:
        for ph in journal_placeholders:
            html_content = html_content.replace(escape(ph), bold_journal)
    for ph in email_placeholders:
        html_content = html_content.replace(escape(ph), escape(r_email))

    # Automatically detect and bold journal names in text:
    def bold_journal_pattern(m):
        txt = m.group(1).strip()
        return f'<b style="{bold_style}"><font face="Times New Roman, Times, serif" color="#000000">{txt}</font></b>'

    # Match "Journal/Annals/Archives/Review/Letters of ..." inline
    html_content = re.sub(
        r'(?i)\b((?:(?:International|Global|American|European|Asian|World|British|Open|Clinical|National)\s+)?(?:Journal|Annals|Archives|Review|Letters)\s+of\s+[A-Za-z0-9\s,\'&-]+?(?:\s+(?:Letters|Reports|Review))?)(?=[\.,;:\n]|\s+[|]|\s+(?:for|to|with|in|by|which|where|and\s+its)\b|$)',
        bold_journal_pattern,
        html_content
    )

    # 3. Intelligent Section & Paragraph Decomposition:
    raw_lines = [l.strip() for l in html_content.splitlines() if l.strip()]

    salutation_block = []
    greetings_block = []
    body_lines = []
    closing_block = []

    mode = 'salutation'
    for line in raw_lines:
        if mode == 'salutation':
            if re.match(r'(?i)^\s*(?:<[^>]+>\s*)*(?:dear|hello|hi|respected|to)\b', line):
                salutation_block.append(line)
                mode = 'greetings'
                continue
            mode = 'greetings'

        if mode == 'greetings':
            if re.match(r'(?i)^\s*(?:<[^>]+>\s*)*greetings\b', line):
                greetings_block.append(line)
                mode = 'body'
                continue
            mode = 'body'

        if mode == 'body':
            if re.match(r'(?i)^\s*(?:<[^>]+>\s*)*(?:with\s+regards|best\s+regards|warm\s+regards|kind\s+regards|regards|sincerely\s+yours|sincerely|yours\s+sincerely|thanking\s+you|thanks\s+and\s+regards|thanks)\b', line):
                closing_block.append(line)
                mode = 'closing'
                continue
            body_lines.append(line)
            continue

        if mode == 'closing':
            closing_block.append(line)

    # Divide body lines into balanced, spacious paragraphs:
    final_body_blocks = []
    curr_para = []
    for bl in body_lines:
        curr_para.append(bl)
        # Check if bl ends a sentence
        if len(curr_para) >= 2 and re.search(r'[\.\?!]$', bl):
            final_body_blocks.append(" ".join(curr_para))
            curr_para = []
    if curr_para:
        final_body_blocks.append(" ".join(curr_para))

    # Bold salutation strictly at the beginning of the salutation line
    if salutation_block:
        s_line = salutation_block[0]
        s_line = re.sub(
            r'(?i)^\s*(Dear\s+(?:Dr\.?|Doctor|Prof\.?|Professor|Mr\.?|Ms\.?|Mrs\.?|)|Hello|Hi|Respected\s+Sir/Madam|Respected\s+Sir|Respected\s+Madam|To)(?=\s|<|\b)',
            lambda m: f'<b style="{bold_style}"><font face="Times New Roman, Times, serif" color="#000000">{m.group(1)}</font></b>',
            s_line
        )
        salutation_block[0] = s_line

    # Bold complimentary sign-off strictly on the sign-off line
    if closing_block:
        c_lines_new = []
        for l in closing_block:
            # Bold standalone journal names on sign-off lines
            m_j = re.match(r'^\s*([A-Za-z0-9\'&-]+(?:\s+[A-Za-z0-9\'&-]+){0,5}\s+(?:Journal|Review|Letters|Reports|Archives|Annals)(?:\s+of\s+[A-Za-z0-9\s,\'&-]+)?)\s*$', l, re.I)
            if m_j and not l.startswith("<b"):
                l = f'<b style="{bold_style}"><font face="Times New Roman, Times, serif" color="#000000">{m_j.group(1)}</font></b>'

            l = re.sub(
                r'(?i)^\s*(Best\s+regards|Warm\s+regards|Kind\s+regards|With\s+regards|Regards|Sincerely\s+yours|Sincerely|Yours\s+sincerely|Thanking\s+you|Thanks\s+and\s+regards|Thanks)\b',
                lambda m: f'<b style="{bold_style}"><font face="Times New Roman, Times, serif" color="#000000">{m.group(1)}</font></b>',
                l
            )
            c_lines_new.append(l)
        closing_block = c_lines_new

    # Reconstruct blocks into bulletproof table rows
    table_rows = []
    td_style = (
        "padding: 0 0 18px 0; "
        "margin: 0; "
        "font-family: 'Times New Roman', Times, Georgia, serif; "
        "font-size: 16px; "
        "color: #000000; "
        "line-height: 1.6; "
        "text-align: left;"
    )

    # 1. Salutation (clean 18px separation)
    if salutation_block:
        s_inner = "<br>\n".join(salutation_block)
        table_rows.append(
            f'<tr><td align="left" style="{td_style}"><font face="Times New Roman, Times, serif" size="3" color="#000000"><span style="font-family: \'Times New Roman\', Times, serif; font-size: 16px; color: #000000; line-height: 1.6;">{s_inner}</span></font></td></tr>'
        )

    # 2. Dedicated Greetings Line (clean 18px separation)
    if greetings_block:
        g_inner = "<br>\n".join(greetings_block)
        table_rows.append(
            f'<tr><td align="left" style="{td_style}"><font face="Times New Roman, Times, serif" size="3" color="#000000"><span style="font-family: \'Times New Roman\', Times, serif; font-size: 16px; color: #000000; line-height: 1.6;">{g_inner}</span></font></td></tr>'
        )

    # 3. Body Paragraphs (clean 18px separation per paragraph)
    for b_para in final_body_blocks:
        if not b_para.strip():
            continue
        table_rows.append(
            f'<tr><td align="left" style="{td_style}"><font face="Times New Roman, Times, serif" size="3" color="#000000"><span style="font-family: \'Times New Roman\', Times, serif; font-size: 16px; color: #000000; line-height: 1.6;">{b_para.strip()}</span></font></td></tr>'
        )

    # 4. Closing & Sign-off Block (clean 10px top padding before sign-off)
    if closing_block:
        c_inner = "<br>\n".join(closing_block)
        closing_td_style = (
            "padding: 10px 0 0 0; "
            "margin: 0; "
            "font-family: 'Times New Roman', Times, Georgia, serif; "
            "font-size: 16px; "
            "color: #000000; "
            "line-height: 1.6; "
            "text-align: left;"
        )
        table_rows.append(
            f'<tr><td align="left" style="{closing_td_style}"><font face="Times New Roman, Times, serif" size="3" color="#000000"><span style="font-family: \'Times New Roman\', Times, serif; font-size: 16px; color: #000000; line-height: 1.6;">{c_inner}</span></font></td></tr>'
        )

    # 4. Professional Compliance Opt-Out Footer
    if include_optout:
        optout_td_style = (
            "padding: 24px 0 0 0; "
            "margin: 0; "
            "border-top: 1px solid #e0e0e0; "
            "font-family: 'Times New Roman', Times, Georgia, serif; "
            "font-size: 12px; "
            "color: #777777; "
            "line-height: 1.4; "
            "text-align: left;"
        )
        table_rows.append(
            f'<tr><td align="left" style="{optout_td_style}"><font face="Times New Roman, Times, serif" size="2" color="#777777"><span style="font-family: \'Times New Roman\', Times, serif; font-size: 12px; color: #777777;"><em>If you do not wish to receive future communications regarding this initiative, simply reply with "Unsubscribe" and your address will be promptly removed.</em></span></font></td></tr>'
        )
        plain_text += '\n\n---\nIf you do not wish to receive future communications regarding this initiative, simply reply with "Unsubscribe" and your address will be promptly removed.'

    rows_html = "\n".join(table_rows)

    rich_html = (
        '<table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" '
        'style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; '
        'font-family: \'Times New Roman\', Times, Georgia, serif; font-size: 16px; color: #000000; line-height: 1.6; '
        'margin: 0; padding: 0; text-align: left; background-color: transparent;">\n'
        f'{rows_html}\n'
        '</table>'
    )

    return rich_html, plain_text

# =========================================================================
# WEBMAIL PROVIDER CONFIGURATIONS & MULTI-ACCOUNT PROFILE MANAGER
# =========================================================================
WEBMAIL_PROVIDERS = {
    "Zoho Mail": {
        "name": "Zoho Mail",
        "default_url": "https://mail.zoho.com/zm/",
        "domain_keywords": ["zoho.com", "mail.zoho"],
        "compose_selectors": [
            "button[title*='New Mail' i]",
            "button:has-text('New Mail')",
            "button.zmBtn_primary[title*='New Mail']",
            "button:has-text('Compose Mail')",
            "button:has-text('Compose')",
            "[data-test-id='compose-mail']",
            "button[aria-label='New Mail']"
        ],
        "compose_shortcut": "m",
        "to_selectors": [
            "input[placeholder*='Add contact' i]",
            "input[placeholder*='to' i]",
            "div[data-name='to'] input",
            "#mailTo input",
            ".select2-search__field",
            "input[aria-label*='To' i]",
            "div[contenteditable='true'][aria-label*='To' i]"
        ],
        "subject_selectors": [
            "input[placeholder*='Subject' i]",
            "input[aria-label*='Subject' i]",
            "input[name='subject']",
            "#mailSubject"
        ],
        "body_selectors": [
            "div[contenteditable='true'][aria-label*='Message body' i]",
            "div[contenteditable='true'][aria-label*='body' i]",
            "div[contenteditable='true'][aria-label*='Mail' i]",
            "#mailContent",
            ".zmMailContent",
            "[data-test-id='mail-compose-body']",
            ".zmComposeBody",
            "div[contenteditable='true']"
        ],
        "send_selectors": [
            "button[title*='Send' i]",
            "button:has-text('Send')",
            "span:text-is('Send')",
            "[data-test-id='send-button']",
            ".zmBtn.zmBtn_primary:has-text('Send')"
        ],
        "send_shortcut": "Control+Enter"
    },
    "BigRock Titan Email": {
        "name": "BigRock Titan Email",
        "default_url": "https://app.titan.email",
        "domain_keywords": ["titan.email", "bigrock.titan"],
        "compose_selectors": [
            "button:has-text('Compose')",
            "button[data-testid*='compose' i]",
            "button[aria-label*='Compose' i]",
            "[data-qa='compose-btn']",
            "button:has-text('New Email')",
            "button:has-text('New Message')"
        ],
        "compose_shortcut": "c",
        "to_selectors": [
            "input[placeholder*='To' i]",
            "input[data-testid*='recipient-input']",
            "input[data-testid*='to' i]",
            "div[data-testid*='to' i] input",
            "input[aria-label*='To' i]",
            "input[name='to']"
        ],
        "subject_selectors": [
            "input[placeholder*='Subject' i]",
            "input[data-testid*='subject-input']",
            "input[name='subject']",
            "input[aria-label*='Subject' i]"
        ],
        "body_selectors": [
            "div[contenteditable='true'][data-testid*='editor' i]",
            "div[contenteditable='true'][aria-label*='Message' i]",
            "div[contenteditable='true'][role='textbox']",
            "div[contenteditable='true']"
        ],
        "send_selectors": [
            "button[data-testid*='send' i]",
            "button:has-text('Send')",
            "button[aria-label*='Send' i]",
            "[data-qa='send-btn']"
        ],
        "send_shortcut": "Control+Enter"
    },
    "BigRock Webmail": {
        "name": "BigRock Webmail",
        "default_url": "https://webmail.bigrock.in",
        "domain_keywords": ["webmail", "roundcube"],
        "compose_selectors": [
            "a.button-compose",
            "a:has-text('Compose')",
            "#rcmbtn101",
            "[data-action='compose']",
            "a[title*='Compose' i]",
            "button:has-text('Compose')",
            "button[title*='Compose' i]"
        ],
        "compose_shortcut": "c",
        "to_selectors": [
            "#_to",
            "textarea[name='_to']",
            "input[name='_to']",
            "#compose-to input",
            "input[id*='to' i]"
        ],
        "subject_selectors": [
            "#compose-subject",
            "input[name='_subject']",
            "#_subject",
            "input[id*='subject' i]"
        ],
        "body_selectors": [
            "iframe#composebody_ifr",
            "#composebody",
            "textarea[name='_message']",
            "body#tinymce",
            "div[contenteditable='true']"
        ],
        "send_selectors": [
            "#rcmbtn107",
            "a.button-send",
            "button:has-text('Send')",
            "a:has-text('Send')",
            "input[value='Send']",
            "button[id*='send' i]"
        ],
        "send_shortcut": "Control+Enter"
    }
}

def get_account_profile_dir(acc_num: int) -> str:
    """Returns isolated persistent profile directory for Account 1, 2, or 3."""
    base_dir = os.path.join(APP_BASE_DIR, "mimicmail_browser_profiles")
    os.makedirs(base_dir, exist_ok=True)
    acc_path = os.path.join(base_dir, f"account_{acc_num}")
    if acc_num == 1 and not os.path.exists(acc_path):
        legacy = os.path.join(APP_BASE_DIR, "zoho_browser_profile")
        if os.path.exists(legacy) and os.listdir(legacy):
            return legacy
        legacy_inst1 = os.path.join(APP_BASE_DIR, "zoho_browser_profiles", "instance_1")
        if os.path.exists(legacy_inst1) and os.listdir(legacy_inst1):
            return legacy_inst1
    os.makedirs(acc_path, exist_ok=True)
    return acc_path

def clean_profile_locks(profile_dir: str):
    """Safely sweeps stale Chromium process locks to guarantee glitch-free, clean browser launch."""
    for lock_name in ["SingletonLock", "SingletonSocket", "SingletonCookie", "lockfile"]:
        lp = os.path.join(profile_dir, lock_name)
        if os.path.exists(lp):
            try:
                os.remove(lp)
            except Exception:
                pass

def _ensure_chromium_installed():
    """Automatically ensures Playwright Chromium browser binary is present."""
    try:
        from playwright._impl._driver import compute_driver_executable
        import subprocess
        driver_exec, driver_cli = compute_driver_executable()
        subprocess.run([driver_exec, driver_cli, "install", "chromium"], capture_output=True, text=True, timeout=120)
    except Exception:
        pass

PROFILE_DIR = get_account_profile_dir(1)

# =========================================================================
# LICENSE & 30-DAY / 5-DAY COOLING PERIOD MANAGER
# =========================================================================
LICENSE_FILE = os.path.join(APP_BASE_DIR, ".mimicmail_license.dat")
LICENSE_SECRET_KEY = "MIMIC_SECURE_AUTH_HASH_2026"
ACTIVE_USAGE_WINDOW_SECONDS = 30 * 24 * 3600    # 30 days of active usage
COOLING_WINDOW_SECONDS = 5 * 24 * 3600         # 5 days cooling window after which it restarts

def _generate_license_sig(first_installed: float, cycle_count: int, unlock_expiry: float) -> str:
    import hashlib
    raw = f"{first_installed:.4f}:{cycle_count}:{unlock_expiry:.4f}:{LICENSE_SECRET_KEY}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()

def get_license_state() -> dict:
    """Returns license state: {is_blocked: bool, reason: str, days_left_in_cycle: float, cooling_days_left: float}"""
    now = time.time()
    state = {
        "first_installed": now,
        "cycle_count": 0,
        "unlock_expiry": 0.0,
        "tampered": False
    }

    if os.path.exists(LICENSE_FILE):
        try:
            import json
            with open(LICENSE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            fi = float(data.get("first_installed", now))
            cc = int(data.get("cycle_count", 0))
            ue = float(data.get("unlock_expiry", 0.0))
            sig = data.get("sig", "")
            if _generate_license_sig(fi, cc, ue) == sig:
                state["first_installed"] = fi
                state["cycle_count"] = cc
                state["unlock_expiry"] = ue
            else:
                state["tampered"] = True
        except Exception:
            state["tampered"] = True
    else:
        # First-ever run: initialize license file
        save_license_state(now, 0, 0.0)

    # Calculate cycle progression
    cycle_length = ACTIVE_USAGE_WINDOW_SECONDS + COOLING_WINDOW_SECONDS # 35 days total
    elapsed = max(0.0, now - state["first_installed"])

    # If an administrative unlock key was applied and is still valid
    if state["unlock_expiry"] > now:
        return {
            "is_blocked": False,
            "status": "UNLOCKED",
            "message": "Full License Active",
            "days_left": (state["unlock_expiry"] - now) / 86400.0,
            "cooling_days_left": 0.0
        }

    # Position within recurring (30 days active + 5 days cooling) cycle
    cycle_offset = elapsed % cycle_length

    if cycle_offset < ACTIVE_USAGE_WINDOW_SECONDS:
        # Currently inside the 30-day active usage period
        days_remaining = (ACTIVE_USAGE_WINDOW_SECONDS - cycle_offset) / 86400.0
        return {
            "is_blocked": False,
            "status": "ACTIVE",
            "message": f"Active Usage ({int(days_remaining)} days left)",
            "days_left": days_remaining,
            "cooling_days_left": 0.0
        }
    else:
        # Currently inside the 5-day cooling period
        cooling_remaining = (cycle_length - cycle_offset) / 86400.0
        return {
            "is_blocked": True,
            "status": "COOLING",
            "message": "Cooling Period Active (30-day usage completed)",
            "days_left": 0.0,
            "cooling_days_left": cooling_remaining
        }

def save_license_state(first_installed: float, cycle_count: int, unlock_expiry: float):
    try:
        import json
        sig = _generate_license_sig(first_installed, cycle_count, unlock_expiry)
        payload = {
            "first_installed": first_installed,
            "cycle_count": cycle_count,
            "unlock_expiry": unlock_expiry,
            "sig": sig
        }
        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f)
    except Exception as e:
        print(f"License save warning: {e}")

def apply_admin_unlock(code: str) -> bool:
    """Admin passcodes to extend or immediately bypass cooling period."""
    clean = code.strip().upper()
    if clean in ("NITRO2026", "MIMIC30FREE", "KARTHIKPAY"):
        now = time.time()
        # Extend by 30 days from now
        save_license_state(now, 0, now + (30 * 86400))
        return True
    return False

class ModernDarkZohoAutomationApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MimicMail — Webmail Outreach Automation (Zoho, BigRock Titan & BigRock Webmail)")
        self.root.geometry("940x920")
        self.root.minsize(840, 800)

        # State flags
        self.is_running = False
        self.stop_requested = False
        self.excel_file_path = None
        self.total_rows_count = 0
        self.sent_rows_count = 0
        self.subject_lines_pool = []
        self.subjects_file_path = None

        # Multi-Account Thread-Safe Queues & Worker Threads
        self.account_queues = {1: queue.Queue(), 2: queue.Queue(), 3: queue.Queue()}
        self.account_threads = {}
        self.cmd_queue = self.account_queues[1]  # Backward compatibility alias
        self.ui_queue = queue.Queue()
        self.active_workers_count = 0
        self.total_session_sent = 0
        self._stats_lock = threading.Lock()
        self.seen_emails_lock = threading.Lock()
        self.seen_emails_in_session = set()

        self._apply_dark_theme()
        self._build_ui()
        self._ensure_worker_started()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(30, self._poll_ui_queue)
        self.root.after(800, self._check_cooling_period_status)

    def _poll_ui_queue(self):
        """Processes cross-thread UI requests safely on the Tkinter main thread."""
        try:
            while True:
                item = self.ui_queue.get_nowait()
                kind, payload = item
                if kind == "LOG":
                    self.log_text.config(state=tk.NORMAL)
                    t = datetime.datetime.now().strftime("%H:%M:%S")
                    self.log_text.insert(tk.END, f"[{t}] {payload}\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                elif kind == "EXEC":
                    try:
                        payload()
                    except Exception:
                        pass
        except queue.Empty:
            pass
        except Exception:
            pass
        finally:
            try:
                self.root.after(40, self._poll_ui_queue)
            except Exception:
                pass

    def ui_dispatch(self, func):
        """Enqueues a function to be executed safely on the Tkinter main thread."""
        self.ui_queue.put(("EXEC", func))

    def _ensure_worker_started(self, acc_id: int = None):
        targets = [acc_id] if acc_id else [1, 2, 3]
        for aid in targets:
            try:
                t = self.account_threads.get(aid)
                if not t or not t.is_alive():
                    th = threading.Thread(target=lambda a=aid: self._account_worker_loop(a), daemon=True)
                    self.account_threads[aid] = th
                    th.start()
            except Exception as e:
                self.log(f"Account {aid} worker startup notice: {e}")

    # Aliases to guarantee no attribute lookup errors regardless of underscore convention
    _ensure_workerstarted = _ensure_worker_started
    ensure_worker_started = _ensure_worker_started
    ensure_workerstarted = _ensure_worker_started

    def _apply_dark_theme(self):
        # Monochromatic Minimalist Black & White Palette (White Background Theme)
        self.bg_root = "#f8f9fa"        # Clean crisp light canvas
        self.bg_card = "#ffffff"        # Pure white card surface
        self.bg_card_alt = "#f3f4f6"    # Neutral light surface
        self.border_card = "#e5e7eb"    # Clean subtle monochrome divider border
        self.fg_primary = "#111827"     # Deep crisp black for primary text
        self.fg_muted = "#6b7280"       # Refined neutral slate grey
        self.accent_bw = "#000000"      # High-contrast black accent
        self.accent_red = "#000000"     # Monochrome black accent fallback
        self.btn_bg = "#000000"         # Black button surface
        self.btn_fg = "#ffffff"         # White button text
        self.btn_bg_alt = "#f3f4f6"     # Light secondary button surface

        self.root.configure(bg=self.bg_root)

        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.style.configure(".", background=self.bg_root, foreground=self.fg_primary)

        # Notebook & Tabs (Minimalist Black & White on White Background)
        self.style.configure("Dark.TNotebook", background=self.bg_card, borderwidth=0)
        self.style.configure("Dark.TNotebook.Tab", background="#f3f4f6", foreground="#4b5563", font=("Segoe UI", 9, "bold"), padding=[14, 6])
        self.style.map("Dark.TNotebook.Tab",
                       background=[("selected", "#000000"), ("active", "#e5e7eb")],
                       foreground=[("selected", "#ffffff"), ("active", "#000000")])

        # LabelFrames
        self.style.configure("DarkCard.TLabelframe", background=self.bg_card, bordercolor=self.border_card, relief="solid", borderwidth=1)
        self.style.configure("DarkCard.TLabelframe.Label", background=self.bg_card, foreground="#000000", font=("Segoe UI", 10, "bold"))

        # Frames
        self.style.configure("Root.TFrame", background=self.bg_root)
        self.style.configure("Card.TFrame", background=self.bg_card)

        # Labels
        self.style.configure("Dark.TLabel", background=self.bg_card, foreground=self.fg_primary, font=("Segoe UI", 9))
        self.style.configure("DarkMuted.TLabel", background=self.bg_card, foreground=self.fg_muted, font=("Segoe UI", 9))

        # Buttons (Strictly Black & White High Contrast)
        self.style.configure("StandardDark.TButton", font=("Segoe UI", 9, "bold"), padding=(12, 6),
                             background="#f3f4f6", foreground="#111827", bordercolor="#d1d5db", lightcolor="#ffffff", darkcolor="#e5e7eb")
        self.style.map("StandardDark.TButton",
                       background=[("active", "#e5e7eb"), ("disabled", "#f9fafb")],
                       foreground=[("disabled", "#9ca3af")])

        self.style.configure("ActionBlue.TButton", font=("Segoe UI", 9, "bold"), padding=(14, 6),
                             background="#ffffff", foreground="#000000", bordercolor="#000000", lightcolor="#ffffff", darkcolor="#e5e7eb")
        self.style.map("ActionBlue.TButton",
                       background=[("active", "#f3f4f6"), ("disabled", "#f9fafb")],
                       foreground=[("disabled", "#9ca3af")])

        self.style.configure("ActionGreen.TButton", font=("Segoe UI", 13, "bold"), padding=(24, 12),
                             background="#000000", foreground="#ffffff", bordercolor="#000000", lightcolor="#262626", darkcolor="#000000")
        self.style.map("ActionGreen.TButton",
                       background=[("active", "#262626"), ("disabled", "#e5e7eb")],
                       foreground=[("disabled", "#9ca3af")])

        self.style.configure("DangerRed.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 12),
                             background="#18181b", foreground="#ffffff", bordercolor="#000000", lightcolor="#27272a", darkcolor="#09090b")
        self.style.map("DangerRed.TButton",
                       background=[("active", "#3f3f46"), ("disabled", "#f4f4f5")],
                       foreground=[("disabled", "#a1a1aa")])

        # Progressbar (Solid Black indicator on Clean Light Trough)
        self.style.configure("Green.Horizontal.TProgressbar", troughcolor="#e5e7eb", background="#000000", bordercolor="#d1d5db", lightcolor="#000000", darkcolor="#000000")

        # Sliders
        self.style.configure("Dark.Horizontal.TScale", background=self.bg_card, troughcolor="#e5e7eb", bordercolor=self.border_card)

        # Checkbutton
        self.style.configure("Dark.TCheckbutton", background=self.bg_card, foreground=self.fg_primary, font=("Segoe UI", 9))
        self.style.map("Dark.TCheckbutton", background=[("active", self.bg_card)])

    def _build_ui(self):
        container = ttk.Frame(self.root, style="Root.TFrame", padding="16")
        container.pack(fill=tk.BOTH, expand=True)

        # =========================================================================
        # 1. TOP HEADER / BRANDING BANNER (Minimalist Monochrome Black & White)
        # =========================================================================
        header_card = tk.Frame(container, bg="#ffffff", highlightbackground=self.border_card, highlightthickness=1, padx=20, pady=14)
        header_card.pack(fill=tk.X, pady=(0, 12))

        header_row = tk.Frame(header_card, bg="#ffffff")
        header_row.pack(fill=tk.X)

        title_box = tk.Frame(header_row, bg="#ffffff")
        title_box.pack(side=tk.LEFT, fill=tk.Y)

        title_lbl = tk.Label(title_box, text="MIMICMAIL", bg="#ffffff", fg="#000000", font=("Segoe UI", 18, "bold"))
        title_lbl.pack(anchor=tk.W)

        sub_lbl = tk.Label(title_box, text="Intelligent Outreach & Automation Engine  |  Multi-Subject Randomization", bg="#ffffff", fg=self.fg_muted, font=("Segoe UI", 9))
        sub_lbl.pack(anchor=tk.W)

        badge_box = tk.Frame(header_row, bg="#ffffff")
        badge_box.pack(side=tk.RIGHT)

        badge_content = tk.Frame(badge_box, bg="#f3f4f6", highlightbackground="#e5e7eb", highlightthickness=1, padx=8, pady=4)
        badge_content.pack(side=tk.RIGHT)

        # Monochrome Pulsing Monitor Canvas
        self.ecg_canvas = tk.Canvas(badge_content, width=86, height=26, bg="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb")
        self.ecg_canvas.pack(side=tk.LEFT, padx=(2, 8))
        self.ecg_step = 0
        self._animate_ecg_heartbeat()

        self.badge_lbl = tk.Label(badge_content, text="ACTIVE", bg="#f3f4f6", fg="#000000",
                                  font=("Segoe UI", 8, "bold"))
        self.badge_lbl.pack(side=tk.LEFT, padx=(0, 4))

        # =========================================================================
        # 2. TWO-COLUMN MAIN BODY (Left: Session & Templates | Right: Status & Actions)
        # =========================================================================
        main_columns = ttk.Frame(container, style="Root.TFrame")
        main_columns.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # --- LEFT COLUMN (Config & Templates, Width ~62%) ---
        left_col = ttk.Frame(main_columns, style="Root.TFrame")
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        # Card 1: Webmail Session & Account Management (Supports Zoho, BigRock Titan, BigRock Webmail)
        setup_card = ttk.LabelFrame(left_col, text="  1. Webmail Session & Account Management  ", style="DarkCard.TLabelframe", padding="12")
        setup_card.pack(fill=tk.X, pady=(0, 10))

        # Row 1: Webmail Provider & Target URL
        prov_row = ttk.Frame(setup_card, style="Card.TFrame")
        prov_row.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(prov_row, text="Webmail Provider:", style="Dark.TLabel").pack(side=tk.LEFT)
        self.provider_combo = ttk.Combobox(
            prov_row,
            values=["Zoho Mail", "BigRock Titan Email", "BigRock Webmail"],
            state="readonly",
            width=18
        )
        self.provider_combo.set("Zoho Mail")
        self.provider_combo.pack(side=tk.LEFT, padx=(6, 12))
        self.provider_combo.bind("<<ComboboxSelected>>", self._on_provider_changed)

        ttk.Label(prov_row, text="URL:", style="Dark.TLabel").pack(side=tk.LEFT)
        self.webmail_url_entry = tk.Entry(
            prov_row,
            bg="#ffffff",
            fg="#111827",
            font=("Segoe UI", 9),
            relief="flat",
            highlightbackground=self.border_card,
            highlightcolor="#000000",
            highlightthickness=1
        )
        self.webmail_url_entry.insert(0, WEBMAIL_PROVIDERS["Zoho Mail"]["default_url"])
        self.webmail_url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0), ipady=2)

        # Row 2: Account Login Selector & Excel Browse
        btn_row = ttk.Frame(setup_card, style="Card.TFrame")
        btn_row.pack(fill=tk.X, pady=(2, 6))

        ttk.Label(btn_row, text="Select Account:", style="Dark.TLabel").pack(side=tk.LEFT)
        self.acc_select_combo = ttk.Combobox(
            btn_row,
            values=["Account 1", "Account 2", "Account 3"],
            state="readonly",
            width=11
        )
        self.acc_select_combo.set("Account 1")
        self.acc_select_combo.pack(side=tk.LEFT, padx=(6, 8))

        self.login_btn = ttk.Button(btn_row, text="🌐 Launch & Login Selected Account", command=self.request_open_selected_account, style="ActionBlue.TButton")
        self.login_btn.pack(side=tk.LEFT)

        self.browse_btn = ttk.Button(btn_row, text="📁 Browse Excel (.xlsx)", command=self.browse_excel, style="ActionBlue.TButton")
        self.browse_btn.pack(side=tk.RIGHT)

        # Row 3: Account Status Badges & Registry Status
        info_row = ttk.Frame(setup_card, style="Card.TFrame")
        info_row.pack(fill=tk.X, pady=(2, 0))

        self.acc_indicators = {}
        for aid in range(1, 4):
            lbl = tk.Label(info_row, text=f"● Acct {aid}: Ready", bg="#f3f4f6", fg=self.fg_muted, font=("Segoe UI", 8, "bold"), padx=5, pady=2, highlightbackground="#e5e7eb", highlightthickness=1)
            lbl.pack(side=tk.LEFT, padx=(0, 6))
            self.acc_indicators[aid] = lbl

        self.session_indicator = self.acc_indicators[1]

        self.file_label = ttk.Label(info_row, text="📄 No registry loaded (Required: email, author name, title)", style="DarkMuted.TLabel")
        self.file_label.pack(side=tk.RIGHT)

        # Card 2: Email Templates & Subject Line Strategy (.txt Random Selection)
        tmpl_card = ttk.LabelFrame(left_col, text="  2. Outreach Templates & Subject Strategy (.txt Random Selection)  ", style="DarkCard.TLabelframe", padding="12")
        tmpl_card.pack(fill=tk.BOTH, expand=True)

        subj_box = ttk.Frame(tmpl_card, style="Card.TFrame")
        subj_box.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(subj_box, text="Subject Pool (.txt):", style="Dark.TLabel").pack(side=tk.LEFT)

        self.upload_subj_btn = ttk.Button(subj_box, text="📄 Upload .txt Subjects", command=self.upload_subject_file, style="ActionBlue.TButton")
        self.upload_subj_btn.pack(side=tk.RIGHT, padx=(6, 0))

        self.clear_subj_btn = ttk.Button(subj_box, text="Clear", command=self.clear_subject_pool, style="StandardDark.TButton")
        self.clear_subj_btn.pack(side=tk.RIGHT, padx=(4, 0))

        self.subj_entry = tk.Entry(subj_box, bg="#ffffff", fg="#111827", insertbackground="#000000",
                                   highlightbackground=self.border_card, highlightcolor="#000000", highlightthickness=1,
                                   relief="flat", font=("Segoe UI", 10), state="readonly")
        self.subj_entry.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(8, 0), ipady=3)

        # Subject status indicator bar
        subj_status_row = ttk.Frame(tmpl_card, style="Card.TFrame")
        subj_status_row.pack(fill=tk.X, pady=(0, 6))
        self.subj_status_lbl = ttk.Label(subj_status_row, text="⚠️ Required: Upload a .txt file with subject lines (Randomly picked for every email)", style="DarkMuted.TLabel")
        self.subj_status_lbl.pack(side=tk.LEFT)

        # Notebook tabs for 3 templates
        nb_header = ttk.Frame(tmpl_card, style="Card.TFrame")
        nb_header.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(nb_header, text="Paste or load email message templates below (Picks randomly per recipient):", style="DarkMuted.TLabel").pack(side=tk.LEFT)
        tk.Label(nb_header, text="Times New Roman | Bold {name}, {title} & {journal}", bg="#f3f4f6", fg="#111827", font=("Segoe UI", 8, "bold"), padx=6, pady=2, highlightbackground="#e5e7eb", highlightthickness=1).pack(side=tk.RIGHT)

        self.template_notebook = ttk.Notebook(tmpl_card, style="Dark.TNotebook")
        self.template_notebook.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        self.template_text_widgets = []
        times_font = tkfont.Font(family="Times New Roman", size=11)
        for i in range(3):
            tab_frame = ttk.Frame(self.template_notebook, style="Card.TFrame")
            self.template_notebook.add(tab_frame, text=f" Template {i+1} ")

            tb = ttk.Frame(tab_frame, style="Card.TFrame")
            tb.pack(fill=tk.X, pady=(4, 4))
            ttk.Label(tb, text=f"Slot {i+1} message text:", style="DarkMuted.TLabel").pack(side=tk.LEFT)

            load_btn = ttk.Button(tb, text=f"📂 Load File", command=lambda idx=i: self._load_template_file(idx), style="StandardDark.TButton")
            load_btn.pack(side=tk.RIGHT)

            clear_btn = ttk.Button(tb, text="🗑 Clear", command=lambda idx=i: self._clear_template_slot(idx), style="StandardDark.TButton")
            clear_btn.pack(side=tk.RIGHT, padx=(0, 4))

            txt = tk.Text(tab_frame, wrap=tk.WORD, height=10, bg="#ffffff", fg="#111827", insertbackground="#000000",
                          relief="flat", font=times_font, highlightbackground=self.border_card, highlightthickness=1)
            initial_content = SAMPLE_HTML_TEMPLATES[i] if i < len(SAMPLE_HTML_TEMPLATES) else ""
            txt.insert(tk.END, initial_content)
            txt.pack(fill=tk.BOTH, expand=True, pady=(0, 2))
            self.template_text_widgets.append(txt)

        # --- RIGHT COLUMN (Controls, Status Graph & Progress, Width ~38%) ---
        right_col = ttk.Frame(main_columns, style="Root.TFrame", width=340)
        right_col.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, padx=(8, 0))

        # Section: BIG PROMINENT ACTION BUTTONS
        action_card = ttk.LabelFrame(right_col, text="  3. Mission Control  ", style="DarkCard.TLabelframe", padding="14")
        action_card.pack(fill=tk.X, pady=(0, 10))

        self.start_btn = ttk.Button(action_card, text="🚀 START AUTOMATION", command=self.request_start_automation, style="ActionGreen.TButton")
        self.start_btn.pack(fill=tk.X, pady=(2, 6))

        self.stop_btn = ttk.Button(action_card, text="🛑 EMERGENCY STOP", command=self.stop_automation, state=tk.DISABLED, style="DangerRed.TButton")
        self.stop_btn.pack(fill=tk.X)

        # Delivery Cadence / Pacing
        pacing_box = tk.Frame(action_card, bg="#ffffff", pady=6)
        pacing_box.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(pacing_box, text="⏳ Inter-Email Pacing Interval:", style="Dark.TLabel").pack(anchor=tk.W)
        slider_row = ttk.Frame(pacing_box, style="Card.TFrame")
        slider_row.pack(fill=tk.X, pady=(2, 4))

        self.interval_var = tk.DoubleVar(value=2.0)
        self.slider = ttk.Scale(slider_row, from_=1.0, to=5.0, variable=self.interval_var, command=self._slider_changed, style="Dark.Horizontal.TScale")
        self.slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.slider_val_lbl = ttk.Label(slider_row, text="2.0 mins", style="Dark.TLabel", width=8)
        self.slider_val_lbl.pack(side=tk.RIGHT)

        self.jitter_var = tk.BooleanVar(value=True)
        self.jitter_check = ttk.Checkbutton(pacing_box, text="Human Behavioral Jitter (±15-35s)",
                                            variable=self.jitter_var, style="Dark.TCheckbutton")
        self.jitter_check.pack(anchor=tk.W)

        # Multi-Account Concurrency Setting (Parallel run across up to 3 mail accounts)
        tab_row = ttk.Frame(pacing_box, style="Card.TFrame")
        tab_row.pack(fill=tk.X, pady=(6, 2))
        ttk.Label(tab_row, text="⚡ Parallel Run Accounts:", style="Dark.TLabel").pack(side=tk.LEFT)
        self.parallel_accounts_var = tk.StringVar(value="3 Accounts Parallel (Fastest Outreach)")
        self.parallel_accounts_combo = ttk.Combobox(
            tab_row,
            textvariable=self.parallel_accounts_var,
            values=[
                "3 Accounts Parallel (Fastest Outreach)",
                "2 Accounts Parallel",
                "1 Account (Single)"
            ],
            state="readonly",
            width=26
        )
        self.parallel_accounts_combo.pack(side=tk.RIGHT)
        ttk.Label(pacing_box, text="(Dispatches across up to 3 separate mail accounts concurrently)", style="DarkMuted.TLabel").pack(anchor=tk.W)

        # 40-Mail Chunk Rest Interval Selector (1 hr, 1.5 hr, 2 hr, 3 hr)
        chunk_rest_row = ttk.Frame(pacing_box, style="Card.TFrame")
        chunk_rest_row.pack(fill=tk.X, pady=(6, 2))
        ttk.Label(chunk_rest_row, text="💤 Rest After 40 Mails:", style="Dark.TLabel").pack(side=tk.LEFT)
        self.chunk_rest_hours_var = tk.DoubleVar(value=1.5)
        self.chunk_rest_combo = ttk.Combobox(
            chunk_rest_row,
            textvariable=self.chunk_rest_hours_var,
            values=[1.0, 1.5, 2.0, 2.5, 3.0],
            state="readonly",
            width=6
        )
        self.chunk_rest_combo.pack(side=tk.RIGHT)
        ttk.Label(pacing_box, text="Hours between 40-mail waves (Recommended: 1.5 - 2.0 hrs)", style="DarkMuted.TLabel").pack(anchor=tk.W)

        # Max Session Cap Selector (Default: 150 emails)
        cap_row = ttk.Frame(pacing_box, style="Card.TFrame")
        cap_row.pack(fill=tk.X, pady=(6, 2))
        ttk.Label(cap_row, text="🎯 Daily / Session Cap:", style="Dark.TLabel").pack(side=tk.LEFT)
        self.session_cap_var = tk.IntVar(value=150)
        self.session_cap_combo = ttk.Combobox(
            cap_row,
            textvariable=self.session_cap_var,
            values=[50, 100, 120, 150, 200, 250],
            state="readonly",
            width=6
        )
        self.session_cap_combo.pack(side=tk.RIGHT)
        ttk.Label(pacing_box, text="Max emails per session (Set to 150 for optimal inbox safety)", style="DarkMuted.TLabel").pack(anchor=tk.W)

        # Deliverability & Spam Guard Options
        guard_box = tk.Frame(action_card, bg="#ffffff", pady=4)
        guard_box.pack(fill=tk.X, pady=(4, 0))

        self.optout_var = tk.BooleanVar(value=True)
        self.optout_check = ttk.Checkbutton(guard_box, text="Include Safe Opt-Out Footer (High Deliverability)",
                                            variable=self.optout_var, style="Dark.TCheckbutton")
        self.optout_check.pack(anchor=tk.W)

        self.dedup_var = tk.BooleanVar(value=True)
        self.dedup_check = ttk.Checkbutton(guard_box, text="Deduplicate Recipient Emails (Prevent Repeat Spams)",
                                           variable=self.dedup_var, style="Dark.TCheckbutton")
        self.dedup_check.pack(anchor=tk.W)

        # Section: LIVE STATUS & GRAPH
        status_card = ttk.LabelFrame(right_col, text="  4. Delivery Status & Graph  ", style="DarkCard.TLabelframe", padding="14")
        status_card.pack(fill=tk.BOTH, expand=True)

        self.timer_display = tk.Label(status_card, text="⏱ Queue: Ready", bg="#ffffff", fg="#000000", font=("Segoe UI", 10, "bold"))
        self.timer_display.pack(anchor=tk.W, pady=(0, 4))

        self.stats_lbl = ttk.Label(status_card, text="Total: 0  |  Sent: 0  |  Pending: 0", style="DarkMuted.TLabel")
        self.stats_lbl.pack(anchor=tk.W, pady=(0, 6))

        # Loading Progress Bar
        pbar_box = ttk.Frame(status_card, style="Card.TFrame")
        pbar_box.pack(fill=tk.X, pady=(0, 6))

        self.progress_bar = ttk.Progressbar(pbar_box, style="Green.Horizontal.TProgressbar", mode="determinate", maximum=100, value=0)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.progress_pct_lbl = tk.Label(pbar_box, text="0%", bg="#ffffff", fg="#000000", font=("Segoe UI", 10, "bold"), width=5)
        self.progress_pct_lbl.pack(side=tk.RIGHT)

        # Status Graph Canvas
        ttk.Label(status_card, text="Live Ratio Breakdown:", style="Dark.TLabel").pack(anchor=tk.W, pady=(4, 2))
        self.graph_canvas = tk.Canvas(status_card, bg="#f3f4f6", height=38, highlightthickness=1, highlightbackground=self.border_card)
        self.graph_canvas.pack(fill=tk.X, pady=(0, 4))
        self.graph_canvas.bind("<Configure>", lambda e: self._draw_status_graph())

        self.graph_legend_lbl = tk.Label(status_card, text="⬛ Sent: 0  |  ◽ Pending: 0  |  ◻ Errors: 0",
                                         bg="#ffffff", fg=self.fg_muted, font=("Segoe UI", 8), justify=tk.LEFT)
        self.graph_legend_lbl.pack(anchor=tk.W)

        # =========================================================================
        # 3. LIVE TERMINAL LOG AT THE BOTTOM (Clean Black & White Background)
        # =========================================================================
        log_frame = ttk.LabelFrame(container, text="  Live Outreach Audit Terminal  ", style="DarkCard.TLabelframe", padding="8")
        log_frame.pack(fill=tk.X)

        self.log_text = tk.Text(log_frame, height=4, state=tk.DISABLED, wrap=tk.WORD, bg="#ffffff", fg="#111827",
                                insertbackground="#000000", relief="flat", font=("Consolas", 8),
                                highlightbackground=self.border_card, highlightthickness=1)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _draw_status_graph(self, sent: int = 0, pending: int = 0, errors: int = 0):
        self.graph_canvas.delete("all")
        w = self.graph_canvas.winfo_width()
        h = self.graph_canvas.winfo_height()
        if w <= 1:
            w = 300
        if h <= 1:
            h = 38

        total = sent + pending + errors
        if total == 0:
            self.graph_canvas.create_rectangle(0, 0, w, h, fill="#f3f4f6", outline="")
            self.graph_canvas.create_text(w // 2, h // 2, text="Load spreadsheet to view graph", fill="#9ca3af", font=("Segoe UI", 8, "italic"))
            return

        sent_pct = (sent / total)
        pending_pct = (pending / total)
        err_pct = (errors / total)

        sent_w = int(w * sent_pct)
        pending_w = int(w * pending_pct)
        err_w = w - sent_w - pending_w

        # Draw segmented colored bar in Pure Monochrome Black, Slate Grey, and Light Grey
        x_curr = 0
        if sent_w > 0:
            # Solid Black for Sent
            self.graph_canvas.create_rectangle(x_curr, 0, x_curr + sent_w, h, fill="#000000", outline="")
            if sent_w > 40:
                self.graph_canvas.create_text(x_curr + sent_w // 2, h // 2, text=f"{sent}", fill="#ffffff", font=("Segoe UI", 8, "bold"))
            x_curr += sent_w

        if pending_w > 0:
            # Slate Grey for Pending
            self.graph_canvas.create_rectangle(x_curr, 0, x_curr + pending_w, h, fill="#6b7280", outline="")
            if pending_w > 40:
                self.graph_canvas.create_text(x_curr + pending_w // 2, h // 2, text=f"{pending}", fill="#ffffff", font=("Segoe UI", 8, "bold"))
            x_curr += pending_w

        if err_w > 0:
            # Light Grey for Errors/Skipped
            self.graph_canvas.create_rectangle(x_curr, 0, w, h, fill="#d1d5db", outline="")
            if err_w > 30:
                self.graph_canvas.create_text(x_curr + err_w // 2, h // 2, text=f"{errors}", fill="#111827", font=("Segoe UI", 8, "bold"))

        legend_text = f"⬛ Sent: {sent} ({sent_pct*100:.0f}%)   |   ◽ Pending: {pending} ({pending_pct*100:.0f}%)   |   ◻ Errors: {errors}"
        self.graph_legend_lbl.config(text=legend_text)

    def _animate_ecg_heartbeat(self):
        """Draws a live, animated monochrome ECG heartbeat pulse wave with a lead dot on a clean white monitor."""
        if not hasattr(self, "ecg_canvas") or not self.ecg_canvas.winfo_exists():
            return

        w = 86
        h = 26
        mid_y = h / 2.0

        self.ecg_canvas.delete("all")

        # 1. Subtle background grid lines
        for gy in [6, 13, 20]:
            self.ecg_canvas.create_line(0, gy, w, gy, fill="#f3f4f6", width=1)
        for gx in range(10, w, 15):
            self.ecg_canvas.create_line(gx, 0, gx, h, fill="#f3f4f6", width=1)

        # 2. Wave pattern template (normalized y-offsets around center)
        pattern = [0, 0, 0, -2, -1, 0, 0, 2, -10, 6, -1, 0, -3, -4, -2, 0, 0, 0, 0, 0]
        pat_len = len(pattern)

        points = []
        for x in range(0, w, 2):
            idx = int((x + self.ecg_step * 2) / 3) % pat_len
            y = mid_y + pattern[idx]
            points.append((x, y))

        # 3. Ambient trace
        for i in range(len(points) - 1):
            x1, y1 = points[i]
            x2, y2 = points[i + 1]
            self.ecg_canvas.create_line(x1, y1 + 1.0, x2, y2 + 1.0, fill="#e5e7eb", width=2)

        # 4. Sharp black primary pulse line
        for i in range(len(points) - 1):
            x1, y1 = points[i]
            x2, y2 = points[i + 1]
            self.ecg_canvas.create_line(x1, y1, x2, y2, fill="#000000", width=1.5)

        # 5. Lead pulse dot
        lead_x = (self.ecg_step * 3) % w
        lead_idx = int((lead_x + self.ecg_step * 2) / 3) % pat_len
        lead_y = mid_y + pattern[lead_idx]

        # Soft ring
        self.ecg_canvas.create_oval(lead_x - 3, lead_y - 3, lead_x + 3, lead_y + 3, fill="#e5e7eb", outline="")
        # Solid black center
        self.ecg_canvas.create_oval(lead_x - 2, lead_y - 2, lead_x + 2, lead_y + 2, fill="#000000", outline="#000000")

        self.ecg_step += 1
        self.root.after(35, self._animate_ecg_heartbeat)

    def upload_subject_file(self):
        """Uploads a .txt file containing subject lines (one per line) for random selection."""
        path = filedialog.askopenfilename(
            title="Select Text File with Subject Lines",
            filetypes=[("Text Files (*.txt)", "*.txt"), ("All Files (*.*)", "*.*")]
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                lines = [line.strip() for line in f if line.strip()]

            if not lines:
                messagebox.showwarning("Empty File", "The chosen .txt file contains no non-empty subject lines.")
                return

            self.subject_lines_pool = lines
            self.subjects_file_path = path
            fname = os.path.basename(path)
            self.subj_status_lbl.config(
                text=f"✓ Loaded {len(lines)} subjects from '{fname}' (Randomized per email)",
                foreground=self.fg_primary
            )
            self.subj_entry.config(state="normal")
            self.subj_entry.delete(0, tk.END)
            self.subj_entry.insert(0, f"Loaded {len(lines)} subjects from '{fname}'")
            self.subj_entry.config(state="readonly")
            self.log(f"Loaded {len(lines)} subject line(s) from '{fname}'. Random picker enabled for each message.")
        except Exception as e:
            messagebox.showerror("Read Error", f"Unable to read subject file:\n{str(e)}")

    def clear_subject_pool(self):
        """Clears subject pool."""
        self.subject_lines_pool = []
        self.subjects_file_path = None
        self.subj_entry.config(state="normal")
        self.subj_entry.delete(0, tk.END)
        self.subj_entry.config(state="readonly")
        self.subj_status_lbl.config(
            text="⚠️ Required: Upload a .txt file with subject lines (Randomly picked for every email)",
            foreground=self.fg_muted
        )
        self.log("Cleared subject lines pool.")

    def _load_template_file(self, idx: int):
        path = filedialog.askopenfilename(
            title=f"Select HTML Template for Slot {idx+1}",
            filetypes=[("HTML / Text Files", "*.html *.htm *.txt"), ("All Files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
            self.template_text_widgets[idx].delete("1.0", tk.END)
            self.template_text_widgets[idx].insert(tk.END, content)
            self.log(f"Successfully loaded file '{os.path.basename(path)}' into Template Slot {idx+1}.")
        except Exception as e:
            messagebox.showerror("File Error", f"Unable to read file:\n{str(e)}")

    def _clear_template_slot(self, idx: int):
        self.template_text_widgets[idx].delete("1.0", tk.END)
        self.log(f"Cleared Template Slot {idx+1}.")

    def _slider_changed(self, val):
        self.slider_val_lbl.config(text=f"{float(val):.1f} mins")

    def log(self, text: str):
        if hasattr(self, "ui_queue"):
            self.ui_queue.put(("LOG", str(text)))
        else:
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {text}")

    def _on_provider_changed(self, event=None):
        prov = self.provider_combo.get()
        if prov in WEBMAIL_PROVIDERS:
            default_url = WEBMAIL_PROVIDERS[prov]["default_url"]
            self.webmail_url_entry.delete(0, tk.END)
            self.webmail_url_entry.insert(0, default_url)
            self.log(f"Switched webmail provider: {prov} -> {default_url}")

    def request_open_selected_account(self):
        sel = self.acc_select_combo.get()
        aid = 1
        if "2" in sel:
            aid = 2
        elif "3" in sel:
            aid = 3
        self.request_open_browser(aid)

    def request_open_browser(self, acc_id: int = 1):
        self._ensure_worker_started(acc_id)
        prov = self.provider_combo.get() if hasattr(self, "provider_combo") else "Zoho Mail"
        url = self.webmail_url_entry.get().strip() if hasattr(self, "webmail_url_entry") else WEBMAIL_PROVIDERS.get(prov, {}).get("default_url", "https://mail.zoho.com/zm/")
        self.account_queues[acc_id].put(("OPEN_BROWSER", {"provider_name": prov, "url": url}))

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Recipient Spreadsheet",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not path:
            return

        try:
            df = pd.read_excel(path)
            
            # Validation 1: Check if workbook has rows
            if df.empty or len(df) == 0:
                messagebox.showwarning(
                    "⚠️ Excel Not Set Properly",
                    f"The selected Excel workbook is empty!\n\nFile: {os.path.basename(path)}\n\nPlease ensure your sheet contains data rows under the headers: name, email, title."
                )
                self.log(f"⚠️ Warning: Selected file '{os.path.basename(path)}' is empty.")
                return

            # Validation 2: Check required columns (email, author name, title) with extensive fuzzy synonyms
            col_map = {str(c).strip().lower(): c for c in df.columns}
            
            author_col = next((col_map[k] for k in [
                "author name", "author_name", "author", "name", "recipient name", 
                "recipient", "full name", "fullname", "contact name", "client name", "person"
            ] if k in col_map), None)
            
            email_col = next((col_map[k] for k in [
                "email", "email address", "e-mail", "mail", "email_id", "email id",
                "recipient email", "e_mail", "mail address", "contact email"
            ] if k in col_map), None)
            
            title_col = next((col_map[k] for k in [
                "title", "paper title", "paper_title", "article title", "article_title",
                "paper", "article", "manuscript title", "manuscript_title", "manuscript",
                "research title", "research_title", "publication title", "publication_title",
                "designation", "role", "job title", "job_title", "position",
                "occupation", "salutation", "profession"
            ] if k in col_map), None)

            missing = []
            if not email_col:
                missing.append("email")
            if not author_col:
                missing.append("author name")
            if not title_col:
                # If title is missing in the sheet, automatically default rather than blocking!
                title_col = "__AUTO_TITLE__"
                df["__AUTO_TITLE__"] = "Professional"

            if missing:
                missing_str = ", ".join(f"'{m}'" for m in missing)
                found_str = ", ".join(f"'{c}'" for c in df.columns) if len(df.columns) > 0 else "None"
                messagebox.showwarning(
                    "⚠️ Excel Not Set Properly",
                    f"The Excel file is not formatted correctly!\n\n"
                    f"Missing required columns: {missing_str}\n"
                    f"Columns found in file: {found_str}\n\n"
                    f"Please make sure your Excel spreadsheet includes the header columns:\n"
                    f"• email (Recipient's email address)\n"
                    f"• author name (Recipient's name or author)\n"
                    f"• title (Recipient's role or designation)\n\n"
                    f"Note: No 'Status' column is required; dispatch tracking is handled automatically."
                )
                self.log(f"⚠️ Warning: '{os.path.basename(path)}' is missing required columns: {missing_str}")
                return

            # Validation 3: Check for empty essential values
            name_col = author_col
            empty_emails = df[email_col].isna().sum() + (df[email_col].astype(str).str.strip() == "").sum()
            
            if empty_emails > 0:
                self.log(f"⚠️ Note: {empty_emails} row(s) have missing/blank email values and will be safely skipped during dispatch.")

            self.excel_file_path = path
            self.file_label.config(text=f"Loaded: {os.path.basename(path)}", foreground=self.fg_primary)

            status_col = next((c for c in df.columns if str(c).strip().lower() == "status"), None)
            total = len(df)
            sent = 0
            errors = 0
            if status_col:
                status_series = df[status_col].astype(str).str.strip().str.upper()
                sent = (status_series == "SENT").sum()
                errors = status_series.str.startswith("ERROR").sum() + (status_series == "INVALID_EMAIL").sum()

            pending = total - sent - errors
            self.total_rows_count = total
            self.sent_rows_count = sent

            # Update progress bar and live status graph
            pct = int((sent / total * 100)) if total > 0 else 0
            self.progress_bar["value"] = pct
            self.progress_pct_lbl.config(text=f"{pct}%")
            self._draw_status_graph(sent=sent, pending=pending, errors=errors)

            self.stats_lbl.config(text=f"Total: {total} | Sent: {sent} | Pending: {pending}")
            self.log(f"Workbook loaded successfully: {total} total rows ({sent} already SENT, {pending} pending).")

        except Exception as e:
            messagebox.showwarning(
                "⚠️ Excel Not Set Properly",
                f"Unable to read or parse the Excel file.\n\n"
                f"Error details:\n{str(e)}\n\n"
                f"Please ensure the file is a valid .xlsx or .xls file and is not open/locked by another application."
            )
            self.log(f"⚠️ Excel parse error on '{os.path.basename(path)}': {str(e)}")

    def _check_cooling_period_status(self):
        lic = get_license_state()
        if lic["is_blocked"]:
            self.badge_lbl.config(text=f"COOLING ({int(lic['cooling_days_left'])}d)", fg="#ffffff", bg="#000000")
            self.show_payment_cooling_dialog(lic)
        else:
            if lic.get("status") == "UNLOCKED":
                self.badge_lbl.config(text="FULL LICENSE", fg="#000000", bg="#f3f4f6")
            else:
                self.badge_lbl.config(text=f"ACTIVE ({int(lic['days_left'])}d)", fg="#000000", bg="#f3f4f6")

    def show_payment_cooling_dialog(self, lic: dict = None):
        """Displays modal payment QR window when 30 days active usage expires."""
        if not lic:
            lic = get_license_state()

        dlg = tk.Toplevel(self.root)
        dlg.title("MimicMail — Usage Cooling & Renewal Period")
        dlg.geometry("450x640")
        dlg.resizable(False, False)
        dlg.configure(bg="#ffffff")
        dlg.transient(self.root)
        dlg.grab_set()

        # Center on parent window
        try:
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 225
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 320
            dlg.geometry(f"+{x}+{y}")
        except Exception:
            pass

        # Title & Cooling Banner
        top_bar = tk.Frame(dlg, bg="#ffffff", pady=10)
        top_bar.pack(fill=tk.X)

        tk.Label(top_bar, text="⏳ 30-Day Usage Window Completed", fg="#000000", bg="#ffffff", font=("Segoe UI", 12, "bold")).pack()
        cooling_days = max(1, int(lic.get("cooling_days_left", 5)))
        tk.Label(top_bar, text=f"Tool is in a 5-day cooling period ({cooling_days} day(s) remaining).\nScan PhonePe QR below or enter verification code to unlock immediately.",
                 fg="#555555", bg="#ffffff", font=("Segoe UI", 9), justify=tk.CENTER).pack(pady=(4, 0))

        # QR Image Display Canvas / Label
        qr_frame = tk.Frame(dlg, bg="#000000", pady=6)
        qr_frame.pack(fill=tk.BOTH, expand=True)

        if os.path.exists(PAYMENT_QR_PATH):
            try:
                from PIL import Image, ImageTk
                pil_img = Image.open(PAYMENT_QR_PATH)
                # Resize keeping aspect ratio
                pil_img.thumbnail((320, 390), Image.Resampling.LANCZOS)
                self.payment_qr_photo = ImageTk.PhotoImage(pil_img)
                qr_lbl = tk.Label(qr_frame, image=self.payment_qr_photo, bg="#000000")
                qr_lbl.pack(pady=4)
            except Exception as img_err:
                tk.Label(qr_frame, text=f"PhonePe QR: MADIPALLI KARTHIK\n(Image error: {img_err})", fg="#ffffff", bg="#000000").pack(pady=20)
        else:
            tk.Label(qr_frame, text="PhonePe Payment Accepted Here\n\nMADIPALLI KARTHIK\nScan & Pay Using PhonePe App",
                     fg="#ffffff", bg="#000000", font=("Segoe UI", 12, "bold"), justify=tk.CENTER).pack(pady=40)

        # Bottom Unlock Code Entry & Actions
        action_frame = tk.Frame(dlg, bg="#ffffff", pady=12, padx=16, highlightbackground="#e5e7eb", highlightthickness=1)
        action_frame.pack(fill=tk.X, side=tk.BOTTOM)

        code_row = tk.Frame(action_frame, bg="#ffffff")
        code_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(code_row, text="Activation Code:", fg="#111827", bg="#ffffff", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        code_entry = tk.Entry(code_row, bg="#ffffff", fg="#000000", insertbackground="#000000", highlightbackground="#d1d5db", highlightcolor="#000000", highlightthickness=1, font=("Segoe UI", 10), width=16)
        code_entry.pack(side=tk.LEFT, padx=(8, 8))

        def try_unlock():
            code = code_entry.get().strip()
            if apply_admin_unlock(code):
                messagebox.showinfo("Success", "✓ License renewed successfully! Full tool functionality restored for 30 days.")
                dlg.destroy()
                self._check_cooling_period_status()
            else:
                messagebox.showerror("Invalid Code", "The activation code entered is invalid. Please complete payment via PhonePe or check your key.")

        unlock_btn = tk.Button(code_row, text="Unlock Now", bg="#000000", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, command=try_unlock)
        unlock_btn.pack(side=tk.LEFT)

        close_btn = tk.Button(action_frame, text="Continue in Cooling Mode (Wait 5 Days)", bg="#f3f4f6", fg="#555555", font=("Segoe UI", 8), relief="flat", command=dlg.destroy)
        close_btn.pack(side=tk.RIGHT)

    def request_start_automation(self):
        try:
            lic = get_license_state()
            if lic["is_blocked"]:
                self.show_payment_cooling_dialog(lic)
                self.log("⚠️ Outreach dispatch paused: Tool is in its 5-day cooling period after 30 days of active usage.")
                return

            if not self.excel_file_path or not os.path.exists(self.excel_file_path):
                messagebox.showwarning("File Missing", "Please select an Excel workbook first using 'Browse Excel (.xlsx)'.")
                return

            subj_pool = list(self.subject_lines_pool) if self.subject_lines_pool else []
            if not subj_pool:
                messagebox.showwarning(
                    "Subject File Required",
                    "Please upload a .txt file containing subject lines using the '📄 Upload .txt Subjects' button.\n\n"
                    "The tool randomly selects a subject from this file for each email."
                )
                return

            active_templates = []
            for txt_widget in self.template_text_widgets:
                content = txt_widget.get("1.0", tk.END).strip()
                if content:
                    active_templates.append(content)

            if not active_templates:
                messagebox.showwarning("No Templates", "Please provide at least 1 template in the 3 slots (paste HTML or load a file).")
                return

            # Read Excel to identify pending recipients
            df = pd.read_excel(self.excel_file_path)
            total_records = len(df)
            has_status_col = any(c.strip().lower() == "status" for c in df.columns)
            if has_status_col:
                status_col_name = next(c for c in df.columns if c.strip().lower() == "status")
                status_series = df[status_col_name].astype(str).str.strip().str.upper()
                pending_indices = [i for i, val in enumerate(status_series) if val != "SENT"]
            else:
                pending_indices = list(range(total_records))

            if not pending_indices:
                messagebox.showinfo("Completed", "All recipients in this spreadsheet are already sent!")
                return

            # Determine concurrency mode (1, 2, or 3 accounts)
            concur_mode = self.parallel_accounts_combo.get() if hasattr(self, "parallel_accounts_combo") else "3 Accounts"
            if "1" in concur_mode:
                active_accounts = [1]
            elif "2" in concur_mode:
                active_accounts = [1, 2]
            else:
                active_accounts = [1, 2, 3]

            prov = self.provider_combo.get() if hasattr(self, "provider_combo") else "Zoho Mail"
            url = self.webmail_url_entry.get().strip() if hasattr(self, "webmail_url_entry") else WEBMAIL_PROVIDERS.get(prov, {}).get("default_url", "https://mail.zoho.com/zm/")

            self.is_running = True
            self.stop_requested = False
            self._toggle_ui(running=True)

            self.seen_emails_in_session.clear()
            self.total_session_sent = 0
            self.active_workers_count = len(active_accounts)

            # Build thread-safe synchronized task queue
            task_queue = queue.Queue()
            for r_idx in pending_indices:
                task_queue.put(r_idx)

            params = {
                "provider_name": prov,
                "url": url,
                "subj_pool": subj_pool,
                "body_tmpls": active_templates,
                "base_mins": self.interval_var.get(),
                "use_jitter": self.jitter_var.get(),
                "excel_path": self.excel_file_path,
                "include_optout": self.optout_var.get(),
                "deduplicate_emails": self.dedup_var.get(),
                "chunk_rest_hours": max(1.0, float(self.chunk_rest_hours_var.get())),
                "max_session_cap": max(1, int(self.session_cap_var.get())),
                "total_pending": len(pending_indices)
            }

            self.log(f"=== Starting Automation: {len(pending_indices)} pending emails | {len(active_accounts)} Parallel Account(s) ({prov}) ===")
            self.ui_dispatch(lambda tp=len(pending_indices), nw=len(active_accounts): self.timer_display.config(
                text=f"🚀 Active: {tp} Pending across {nw} Accounts", fg="#15803d"
            ))

            for aid in active_accounts:
                self._ensure_worker_started(aid)
                self.account_queues[aid].put(("START_BATCH", {
                    "task_queue": task_queue,
                    "params": params,
                    "df": df
                }))

        except Exception as e:
            self.is_running = False
            self.stop_requested = False
            self._toggle_ui(running=False)
            self.log(f"⚠️ Error initiating automation: {e}")
            messagebox.showerror("Execution Error", f"Unable to start automation:\n\n{e}")

    def stop_automation(self):
        if self.is_running:
            self.stop_requested = True
            self.log("Emergency Stop triggered! Halting after current cycle...")
            self.stop_btn.config(state=tk.DISABLED)

    def _toggle_ui(self, running: bool):
        self.start_btn.config(state=tk.DISABLED if running else tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL if running else tk.DISABLED)
        self.browse_btn.config(state=tk.DISABLED if running else tk.NORMAL)

    # =========================================================================
    # DEDICATED MULTI-ACCOUNT WORKER LOOPS (ONE PER ACCOUNT, ZERO GREENLET ERRORS)
    # =========================================================================
    def _account_worker_loop(self, acc_id: int):
        profile_dir = get_account_profile_dir(acc_id)
        playwright_instance = None
        browser_context = None
        browser_page = None

        _ensure_chromium_installed()

        try:
            playwright_instance = sync_playwright().start()
            q = self.account_queues[acc_id]

            while True:
                try:
                    cmd, data = q.get(timeout=0.2)
                except queue.Empty:
                    continue

                if cmd == "QUIT":
                    break

                elif cmd == "OPEN_BROWSER":
                    try:
                        provider_name = data.get("provider_name", "Zoho Mail")
                        url = data.get("url", WEBMAIL_PROVIDERS.get(provider_name, {}).get("default_url"))

                        clean_profile_locks(profile_dir)
                        if not (browser_context and browser_page and not browser_page.is_closed()):
                            self.log(f"[Account {acc_id}] Launching persistent Chromium browser ({provider_name})...")
                            self.ui_dispatch(lambda a=acc_id: self.acc_indicators[a].config(text=f"● Acct {a}: Launching...", fg="#000000", bg="#fef08a"))

                            browser_context = playwright_instance.chromium.launch_persistent_context(
                                user_data_dir=profile_dir,
                                headless=False,
                                no_viewport=True,
                                timeout=60000,
                                permissions=["clipboard-read", "clipboard-write"],
                                ignore_default_args=["--enable-automation"],
                                args=[
                                    "--start-maximized",
                                    "--disable-blink-features=AutomationControlled",
                                    "--no-sandbox",
                                    "--disable-infobars",
                                    "--disable-dev-shm-usage",
                                    "--disable-gpu",
                                    "--no-first-run",
                                    "--no-default-browser-check"
                                ]
                            )
                            browser_context.add_init_script(STEALTH_INIT_SCRIPT)
                            browser_page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()

                            try:
                                browser_page.goto(url, timeout=60000, wait_until="domcontentloaded")
                            except Exception as nav_e:
                                self.log(f"[Account {acc_id}] Navigation warning: {nav_e}")

                            self.log(f"[Account {acc_id}] Browser is open! Please log into {provider_name} if prompted.")
                            self.ui_dispatch(lambda a=acc_id: self.acc_indicators[a].config(text=f"● Acct {a}: Logged In / Active", fg="#15803d", bg="#dcfce7"))
                        else:
                            self.log(f"[Account {acc_id}] Browser already open. Bringing to front...")
                            browser_page.bring_to_front()
                    except Exception as err:
                        err_msg = str(err)
                        self.log(f"[Account {acc_id}] Browser launch error: {err_msg}")
                        self.ui_dispatch(lambda a=acc_id: self.acc_indicators[a].config(text=f"● Acct {a}: Error", fg="#dc2626", bg="#fee2e2"))
                        self.ui_dispatch(lambda m=err_msg, a=acc_id: messagebox.showerror(f"Account {a} Browser Error", f"Unable to launch browser for Account {a}:\n\n{m[:250]}\n\nIf another window is using this profile, please close it and retry."))

                elif cmd == "START_BATCH":
                    task_queue = data["task_queue"]
                    params = data["params"]
                    df = data["df"]
                    provider_name = params.get("provider_name", "Zoho Mail")
                    url = params.get("url", WEBMAIL_PROVIDERS.get(provider_name, {}).get("default_url"))

                    try:
                        clean_profile_locks(profile_dir)
                        if not (browser_context and browser_page and not browser_page.is_closed()):
                            self.log(f"[Account {acc_id}] Opening browser session for outreach...")
                            browser_context = playwright_instance.chromium.launch_persistent_context(
                                user_data_dir=profile_dir,
                                headless=False,
                                no_viewport=True,
                                timeout=60000,
                                permissions=["clipboard-read", "clipboard-write"],
                                ignore_default_args=["--enable-automation"],
                                args=[
                                    "--start-maximized",
                                    "--disable-blink-features=AutomationControlled",
                                    "--no-sandbox",
                                    "--disable-infobars",
                                    "--disable-dev-shm-usage",
                                    "--disable-gpu",
                                    "--no-first-run",
                                    "--no-default-browser-check"
                                ]
                            )
                            browser_context.add_init_script(STEALTH_INIT_SCRIPT)
                            browser_page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()

                        browser_page.bring_to_front()

                        # Ensure logged in and in mailbox
                        self._ensure_inbox_ready(browser_page, provider_name, url, acc_id)
                        self.ui_dispatch(lambda a=acc_id: self.acc_indicators[a].config(text=f"● Acct {a}: Automating", fg="#2563eb", bg="#dbeafe"))

                        # Run parallel dispatch on shared queue
                        self._run_account_dispatch(browser_context, browser_page, task_queue, params, df, acc_id)

                    except Exception as batch_err:
                        self.log(f"[Account {acc_id}] Batch execution notice: {batch_err}")
                    finally:
                        self.ui_dispatch(lambda a=acc_id: self.acc_indicators[a].config(text=f"● Acct {a}: Idle", fg=self.fg_muted, bg="#f3f4f6"))
                        self._check_batch_completion()

        except Exception as fatal:
            self.log(f"[Account {acc_id}] Fatal worker error: {fatal}")
        finally:
            if browser_context:
                try:
                    browser_context.close()
                except Exception:
                    pass
            if playwright_instance:
                try:
                    playwright_instance.stop()
                except Exception:
                    pass

    def _ensure_inbox_ready(self, page, provider_name: str, target_url: str, acc_id: int):
        curr_url = page.url.lower() if page.url else ""
        if not curr_url or curr_url.startswith("about:"):
            self.log(f"[Account {acc_id}] Navigating to {provider_name}: {target_url} ...")
            try:
                page.goto(target_url, timeout=60000, wait_until="domcontentloaded")
            except Exception as e_nav:
                self.log(f"[Account {acc_id}] Navigation notice: {e_nav}")
            time.sleep(3)
            curr_url = page.url.lower()

        # Check login detection
        login_detected = False
        if provider_name == "Zoho Mail":
            if "accounts.zoho" in curr_url or "signin" in curr_url or "login" in curr_url:
                login_detected = True
        elif provider_name == "BigRock Titan Email":
            if "login" in curr_url or "signin" in curr_url or "auth" in curr_url:
                login_detected = True
            else:
                try:
                    login_detected = bool(page.query_selector("input[type='password']"))
                except Exception:
                    pass
        elif provider_name == "BigRock Webmail":
            try:
                login_detected = bool(page.query_selector("form#login-form, input[name='user'], input[name='pass'], #rcmloginuser"))
            except Exception:
                pass

        if login_detected:
            self.log(f"[Account {acc_id}] ⚠️ Login/Verification detected for {provider_name}. Waiting for you to complete sign-in in the browser...")
            self.ui_dispatch(lambda a=acc_id: self.acc_indicators[a].config(text=f"● Acct {a}: Waiting for Login...", fg="#b45309", bg="#fef3c7"))
            logged_in = False
            for _ in range(90):
                if self.stop_requested:
                    break
                time.sleep(2)
                try:
                    u = page.url.lower()
                    if provider_name == "Zoho Mail" and "mail.zoho" in u and "accounts.zoho" not in u:
                        logged_in = True
                        break
                    elif provider_name == "BigRock Titan Email" and "app.titan.email" in u and "login" not in u:
                        logged_in = True
                        break
                    elif provider_name == "BigRock Webmail" and not page.query_selector("form#login-form, input[name='user']"):
                        logged_in = True
                        break
                except Exception:
                    pass
            if not logged_in and not self.stop_requested:
                raise Exception(f"Account {acc_id} login timeout: Please complete your {provider_name} sign-in in the browser window and retry.")
            self.log(f"[Account {acc_id}] Login verified! Waiting for mailbox to load...")
            time.sleep(4)

        if provider_name == "Zoho Mail":
            for access_btn_sel in ["a:has-text('Access Email')", "button:has-text('Access Email')", "a:has-text('Access Mail')", "[title*='Access Email']", "a[href*='mail.zoho']"]:
                try:
                    btn = page.query_selector(access_btn_sel)
                    if btn and btn.is_visible():
                        btn.click()
                        time.sleep(4)
                        break
                except Exception:
                    pass
            try:
                mail_app_icon = page.query_selector(".zmNavMail, [data-app='mail'], [title='Mail']")
                if mail_app_icon and mail_app_icon.is_visible():
                    mail_app_icon.click()
                    time.sleep(1.5)
            except Exception:
                pass

    def _run_account_dispatch(self, context, page, task_queue, params, df, acc_id: int):
        excel_path = params["excel_path"]
        subj_pool = params.get("subj_pool", [])
        body_tmpls = params.get("body_tmpls", [])
        base_mins = params["base_mins"]
        use_jitter = params["use_jitter"]
        include_optout = params.get("include_optout", True)
        deduplicate_emails = params.get("deduplicate_emails", True)
        chunk_rest_hours = params.get("chunk_rest_hours", 1.5)
        max_session_cap = params.get("max_session_cap", 150)
        provider_name = params.get("provider_name", "Zoho Mail")

        col_map = {str(c).strip().lower(): c for c in df.columns}
        name_col = next((col_map[k] for k in [
            "author name", "author_name", "author", "name", "recipient name", 
            "recipient", "full name", "fullname", "contact name", "client name", "person"
        ] if k in col_map), "name")
        
        email_col = next((col_map[k] for k in [
            "email", "email address", "e-mail", "mail", "email_id", "email id",
            "recipient email", "e_mail", "mail address", "contact email"
        ] if k in col_map), "email")
        
        title_col = next((col_map[k] for k in [
            "title", "paper title", "paper_title", "article title", "article_title",
            "paper", "article", "manuscript title", "manuscript_title", "manuscript",
            "research title", "research_title", "publication title", "publication_title",
            "designation", "role", "job title", "job_title", "position",
            "occupation", "salutation", "profession"
        ] if k in col_map), "title")

        journal_col = next((col_map[k] for k in [
            "journal name", "journal_name", "journal", "publication", "source journal",
            "journal title", "journal_title", "magazine"
        ] if k in col_map), None)

        import re
        EMAIL_REGEX = re.compile(r"^[\w\.-]+@[\w\.-]+\.\w+$")

        acc_sent_count = 0

        while not self.stop_requested:
            # Check session cap
            with self._stats_lock:
                if self.total_session_sent >= max_session_cap:
                    self.log(f"[Account {acc_id}] Session cap of {max_session_cap} reached.")
                    break

            try:
                r_idx = task_queue.get_nowait()
            except queue.Empty:
                self.log(f"[Account {acc_id}] All pending recipients in task queue completed!")
                break

            row = df.loc[r_idx]
            r_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else "Valued Professional"
            r_email = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
            r_title = str(row[title_col]).strip() if pd.notna(row[title_col]) else "Professional"
            r_journal = str(row[journal_col]).strip() if (journal_col and pd.notna(row[journal_col])) else ""

            if not r_email or not EMAIL_REGEX.match(r_email):
                self.log(f"[Account {acc_id}] Row {r_idx + 1}: SKIPPED (Invalid email '{r_email}')")
                mark_excel_row_status(excel_path, r_idx, "INVALID_EMAIL")
                continue

            clean_email_lower = r_email.strip().lower()
            if deduplicate_emails:
                with self.seen_emails_lock:
                    if clean_email_lower in self.seen_emails_in_session:
                        self.log(f"[Account {acc_id}] Row {r_idx + 1}: SKIPPED (Duplicate '{r_email}' already reached)")
                        mark_excel_row_status(excel_path, r_idx, "SKIPPED_DUPLICATE")
                        continue
                    self.seen_emails_in_session.add(clean_email_lower)

            chosen_body_tmpl = random.choice(body_tmpls)
            chosen_subj_tmpl = random.choice(subj_pool) if subj_pool else DEFAULT_SUBJECT

            # Build HTML with Times New Roman and bold markers
            body_html, body_text = format_email_body_html(
                chosen_body_tmpl,
                name=r_name,
                title=r_title,
                journal_name=r_journal,
                include_optout=include_optout
            )

            subject = chosen_subj_tmpl.replace("{name}", r_name).replace("{title}", r_title).replace("{journal}", r_journal).strip()

            self.log(f"[Account {acc_id}] Sending to Row {r_idx + 1}: {r_name} <{r_email}> | Subj: {subject[:40]}...")

            send_success = False
            try:
                self._send_in_browser(page, r_email, subject, body_html, body_text, provider_name=provider_name)
                send_success = True
            except Exception as send_err:
                self.log(f"[Account {acc_id}] Dispatch failed for {r_email}: {send_err}")
                mark_excel_row_status(excel_path, r_idx, f"FAILED: {str(send_err)[:25]}")

            if send_success:
                mark_excel_row_status(excel_path, r_idx, "SENT")
                acc_sent_count += 1
                with self._stats_lock:
                    self.total_session_sent += 1
                    cur_total_sent = self.total_session_sent
                
                self.log(f"[Account {acc_id}] ✓ Successfully SENT to {r_name} <{r_email}> (Account total: {acc_sent_count} | Session total: {cur_total_sent})")

                # UI updates
                self.ui_dispatch(lambda cts=cur_total_sent, tr=self.total_rows_count, sr=self.sent_rows_count, na=self.active_workers_count: (
                    self.progress_bar.config(value=sr + cts),
                    self.stats_lbl.config(text=f"Total: {tr}  |  Sent: {sr + cts}  |  Pending: {max(0, tr - (sr + cts))}  |  Parallel Accounts: {na} Active")
                ))

                # Humanized pacing delay
                delay_sec = base_mins * 60
                if use_jitter:
                    jitter = random.randint(15, 35) * random.choice([-1, 1])
                    delay_sec = max(30, delay_sec + jitter)

                self.log(f"[Account {acc_id}] Humanized delay: waiting {delay_sec:.1f}s before next outreach...")
                for s in range(int(delay_sec)):
                    if self.stop_requested:
                        break
                    time.sleep(1)

                # 40-mail wave cooldown check per account
                if acc_sent_count > 0 and (acc_sent_count % 40 == 0):
                    self.log(f"[Account {acc_id}] Reached 40 dispatched emails. Entering {chunk_rest_hours:.1f}-hour cooldown wave...")
                    rest_sec = int(chunk_rest_hours * 3600)
                    for r_tick in range(rest_sec):
                        if self.stop_requested:
                            break
                        time.sleep(1)

    def _check_batch_completion(self):
        with self._stats_lock:
            self.active_workers_count = max(0, self.active_workers_count - 1)
            remaining = self.active_workers_count

        if remaining == 0:
            self.is_running = False
            self.ui_dispatch(lambda: self.timer_display.config(text="⏱ Queue: Idle / Completed", fg=self.fg_muted))
            self.ui_dispatch(lambda: self._toggle_ui(running=False))
            self.ui_dispatch(lambda s=self.total_session_sent: messagebox.showinfo(
                "Outreach Complete",
                f"All batches and accounts completed!\n\nTotal emails dispatched across parallel accounts: {s}"
            ))

    def _send_in_browser(self, page, to_email, subject, body_html, body_text=None, provider_name="Zoho Mail"):
        if body_text is None:
            body_text = clean_to_plain_text(body_html)
        prov_info = WEBMAIL_PROVIDERS.get(provider_name, WEBMAIL_PROVIDERS["Zoho Mail"])
        self.log(f"[{provider_name}] Opening compose editor for: {to_email}")

        # 1. Click Compose button
        clicked = False
        compose_selectors = prov_info.get("compose_selectors", [])
        for sel in compose_selectors:
            try:
                btn = page.query_selector(sel)
                if btn and btn.is_visible():
                    btn.click()
                    clicked = True
                    break
            except Exception:
                pass

        if not clicked:
            page.keyboard.press("Escape")
            time.sleep(0.5)
            shortcut = prov_info.get("compose_shortcut", "c")
            page.keyboard.press(shortcut)
            time.sleep(1)
            if not page.query_selector("input[placeholder*='Subject' i], #mailSubject, #compose-subject, input[data-testid*='subject']"):
                page.keyboard.press("c")

        time.sleep(2)

        # 2. Locate and fill 'To' recipient field
        filled_to = False
        to_selectors = prov_info.get("to_selectors", [])
        for sel in to_selectors:
            try:
                el = page.query_selector(sel)
                if el and el.is_visible():
                    el.click()
                    el.fill(to_email)
                    time.sleep(0.3)
                    page.keyboard.press("Enter")
                    filled_to = True
                    break
            except Exception:
                pass

        if not filled_to:
            try:
                inputs = page.query_selector_all("input")
                for inp in inputs:
                    ph = (inp.get_attribute("placeholder") or "").lower()
                    if any(k in ph for k in ["contact", "recipient", "to", "email"]) and inp.is_visible():
                        inp.click()
                        inp.fill(to_email)
                        time.sleep(0.3)
                        page.keyboard.press("Enter")
                        filled_to = True
                        break
            except Exception:
                pass

        time.sleep(0.8)

        # 3. Locate and fill Subject
        filled_subj = False
        subj_selectors = prov_info.get("subject_selectors", [])
        for sel in subj_selectors:
            try:
                el = page.query_selector(sel)
                if el and el.is_visible():
                    el.click()
                    el.fill(subject)
                    filled_subj = True
                    break
            except Exception:
                pass

        if not filled_subj:
            try:
                inputs = page.query_selector_all("input")
                for inp in inputs:
                    ph = (inp.get_attribute("placeholder") or "").lower()
                    if "subject" in ph and inp.is_visible():
                        inp.click()
                        inp.fill(subject)
                        filled_subj = True
                        break
            except Exception:
                pass

        time.sleep(0.8)

        # 4. Locate and fill formatted Message Body strictly in Times New Roman with bold placeholders
        filled_body = False

        # A. Try rich injection inside iframe first (Zoho editor or Roundcube TinyMCE)
        for frame in page.frames:
            try:
                body_elem = frame.query_selector("body[contenteditable='true'], body#tinymce, body")
                if body_elem and body_elem.is_visible():
                    body_elem.click()
                    time.sleep(0.2)
                    frame.keyboard.press("Control+A")
                    frame.keyboard.press("Backspace")
                    time.sleep(0.1)

                    inserted = frame.evaluate(
                        """(html) => {
                            try {
                                document.body.focus();
                                let ok = document.execCommand('insertHTML', false, html);
                                if (!ok || !document.body.innerHTML.trim() || document.body.innerText.trim().length === 0) {
                                    document.body.innerHTML = html;
                                }

                                document.execCommand('selectAll', false, null);
                                document.execCommand('fontName', false, 'Times New Roman');
                                document.execCommand('fontSize', false, '3');
                                document.execCommand('justifyLeft', false, null);

                                const sel = window.getSelection();
                                if (sel) sel.removeAllRanges();

                                document.body.style.setProperty('text-align', 'left', 'important');
                                document.body.style.setProperty('font-family', "'Times New Roman', Times, Georgia, serif", 'important');
                                document.body.style.setProperty('font-size', '16px', 'important');
                                document.body.style.setProperty('color', '#000000', 'important');
                                document.body.style.setProperty('line-height', '1.5', 'important');

                                document.body.querySelectorAll('*').forEach(node => {
                                    node.style.setProperty('font-family', "'Times New Roman', Times, Georgia, serif", 'important');
                                    node.setAttribute('align', 'left');
                                    node.style.setProperty('text-align', 'left', 'important');
                                });

                                document.body.dispatchEvent(new Event('input', { bubbles: true }));
                                document.body.dispatchEvent(new Event('change', { bubbles: true }));
                                return true;
                            } catch (e) {
                                document.body.innerHTML = html;
                                return true;
                            }
                        }""",
                        body_html
                    )
                    if inserted:
                        filled_body = True
                        break
            except Exception:
                pass

        # B. Try contenteditable divs (Zoho, Titan, or Roundcube div)
        if not filled_body:
            body_selectors = prov_info.get("body_selectors", [])
            for sel in body_selectors:
                try:
                    el = page.query_selector(sel)
                    if el and el.is_visible():
                        el.click()
                        time.sleep(0.2)
                        page.keyboard.press("Control+A")
                        page.keyboard.press("Backspace")
                        time.sleep(0.1)

                        inserted = page.evaluate(
                            """([selector, html]) => {
                                const e = document.querySelector(selector);
                                if (!e) return false;
                                try {
                                    e.focus();
                                    let ok = document.execCommand('insertHTML', false, html);
                                    if (!ok || !e.innerHTML.trim() || e.innerText.trim().length === 0) {
                                        e.innerHTML = html;
                                    }

                                    document.execCommand('selectAll', false, null);
                                    document.execCommand('fontName', false, 'Times New Roman');
                                    document.execCommand('fontSize', false, '3');
                                    document.execCommand('justifyLeft', false, null);

                                    const sel = window.getSelection();
                                    if (sel) sel.removeAllRanges();

                                    e.style.setProperty('text-align', 'left', 'important');
                                    e.style.setProperty('font-family', "'Times New Roman', Times, Georgia, serif", 'important');
                                    e.style.setProperty('font-size', '16px', 'important');
                                    e.style.setProperty('color', '#000000', 'important');
                                    e.style.setProperty('line-height', '1.5', 'important');

                                    e.querySelectorAll('*').forEach(node => {
                                        node.style.setProperty('font-family', "'Times New Roman', Times, Georgia, serif", 'important');
                                        node.setAttribute('align', 'left');
                                        node.style.setProperty('text-align', 'left', 'important');
                                    });

                                    e.dispatchEvent(new Event('input', { bubbles: true }));
                                    e.dispatchEvent(new Event('change', { bubbles: true }));
                                    return true;
                                } catch (err) {
                                    e.innerHTML = html;
                                    return true;
                                }
                            }""",
                            [sel, body_html]
                        )
                        if inserted:
                            filled_body = True
                            break
                except Exception:
                    pass

        # C. Try textarea element (e.g., Roundcube plain text composer)
        if not filled_body:
            for ta_sel in ["#composebody", "textarea[name='_message']", "textarea"]:
                try:
                    ta = page.query_selector(ta_sel)
                    if ta and ta.is_visible():
                        ta.click()
                        ta.fill(body_text)
                        filled_body = True
                        break
                except Exception:
                    pass

        # D. Clipboard fallback
        if not filled_body:
            try:
                page.keyboard.press("Tab")
                time.sleep(0.3)
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                page.evaluate(
                    """(html) => {
                        const blob = new Blob([html], { type: 'text/html' });
                        const textBlob = new Blob([html.replace(/<[^>]+>/g, '')], { type: 'text/plain' });
                        navigator.clipboard.write([
                            new ClipboardItem({ 'text/html': blob, 'text/plain': textBlob })
                        ]).catch(() => {});
                    }""",
                    body_html
                )
                time.sleep(0.3)
                page.keyboard.press("Control+v")
                time.sleep(0.5)
                filled_body = True
            except Exception:
                pass

        time.sleep(1.0)

        # 5. Click Send
        clicked_send = False
        send_selectors = prov_info.get("send_selectors", [])
        for sel in send_selectors:
            try:
                btn = page.query_selector(sel)
                if btn and btn.is_visible():
                    btn.click()
                    clicked_send = True
                    break
            except Exception:
                pass

        if not clicked_send:
            page.keyboard.press("Control+Enter")

        time.sleep(3.5)

    def _on_close(self):
        try:
            if hasattr(self, "excel_file_path") and self.excel_file_path:
                close_cached_excel_workbook(self.excel_file_path)
        except Exception:
            pass
        for aid in [1, 2, 3]:
            try:
                q = self.account_queues.get(aid)
                if q:
                    q.put(("QUIT", None))
            except Exception:
                pass
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernDarkZohoAutomationApp(root)
    root.mainloop()
