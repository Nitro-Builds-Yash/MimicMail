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

DEFAULT_SUBJECT = "{title}"
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(_BASE_DIR, "assets", "pulsus_logo.png")
if not os.path.exists(LOGO_PATH):
    LOGO_PATH = os.path.join(_BASE_DIR, "pulsus_logo.png")

# Default user-customizable templates for the 3 slots (Clean, professional text format matching exact spacing and alignment)
SAMPLE_HTML_TEMPLATES = [
    """Dear [Author's Name],

We are pleased to invite you to share your expertise in the field of cancer research by contributing to our upcoming publication. Your work has significantly impacted the community, and we believe your insights would greatly benefit our readers.

This is a valuable opportunity to showcase your research and connect with peers who are equally passionate about advancing cancer treatment. We look forward to your positive response and your essential contribution to this important endeavor.

Best regards,
MimicMail Editorial Board
MimicMail Publishing Directorate""",

    """Dear [Author's Name],

On behalf of our editorial committee, we cordially invite you to contribute your esteemed research to our upcoming edition on cutting-edge oncology and clinical therapeutics. Your pioneering findings continue to inspire colleagues worldwide.

By presenting your work with us, you will reach a global readership of specialists dedicated to accelerating clinical breakthroughs. We would be deeply honored to include your submission in this milestone release.

Best regards,
MimicMail Editorial Board
MimicMail Scientific Council""",

    """Dear [Author's Name],

We are writing to warmly invite your valuable contribution to our specialized medical issue exploring breakthrough approaches in cancer diagnostics and patient care. Your distinguished record in the discipline makes your perspective particularly vital.

This publication provides an international stage to highlight your latest discoveries and engage with top researchers across the scientific domain. We sincerely anticipate your participation and look forward to working together.

Best regards,
MimicMail Editorial Board
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

def build_formatted_email_content(raw_template: str, r_name: str, r_title: str, r_email: str):
    """
    Builds both:
    1. Rich HTML strictly rendered in Times New Roman (12pt), perfectly left-aligned,
       with auto-readjusted paragraph spacing, with '{name}', '{title}', 'Dear Dr' / salutations,
       and 'Best regards' / sign-offs bolded.
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
        "[Title]", "[title]", "[TITLE]", "[Designation]", "[designation]",
        "{title}", "{TITLE}", "{designation}", "{role}"
    ]
    email_placeholders = [
        "[Email]", "[email]", "[EMAIL]", "{email}", "{EMAIL}"
    ]

    plain_text = clean_text
    for ph in name_placeholders:
        plain_text = plain_text.replace(ph, r_name)
    for ph in title_placeholders:
        plain_text = plain_text.replace(ph, r_title)
    for ph in email_placeholders:
        plain_text = plain_text.replace(ph, r_email)

    # 2. Escape text safely for HTML insertion
    escaped_template = escape(clean_text)
    name_escaped = escape(r_name)
    title_escaped = escape(r_title)
    email_escaped = escape(r_email)

    bold_style = "font-family: 'Times New Roman', Times, serif !important; font-weight: bold !important; font-size: 12pt !important;"
    bold_name = f"<b class=\"bold-field\" style=\"{bold_style}\">{name_escaped}</b>"
    bold_title = f"<b class=\"bold-field\" style=\"{bold_style}\">{title_escaped}</b>"

    html_content = escaped_template
    for ph in name_placeholders:
        html_content = html_content.replace(escape(ph), bold_name)
    for ph in title_placeholders:
        html_content = html_content.replace(escape(ph), bold_title)
    for ph in email_placeholders:
        html_content = html_content.replace(escape(ph), email_escaped)

    # Automatically detect and bold salutations: "Dear Dr.", "Dear Dr", "Dear Professor", "Dear Prof.", "Dear"
    def bold_salutation(m):
        txt = m.group(1)
        return f'<b class="bold-field" style="{bold_style}">{txt}</b>'

    html_content = re.sub(
        r'(?i)\b(Dear\s+(?:Dr\.?|Doctor|Prof\.?|Professor|Mr\.?|Ms\.?|Mrs\.?|))(?=\s|<|\b)',
        bold_salutation,
        html_content
    )

    # Automatically detect and bold complimentary sign-offs:
    # "Best regards,", "Warm regards,", "Kind regards,", "With regards,", "Regards,", "Sincerely,", "Sincerely yours,"
    def bold_closing(m):
        txt = m.group(1)
        return f'<b class="bold-field" style="{bold_style}">{txt}</b>'

    html_content = re.sub(
        r'(?i)\b(Best\s+regards|Warm\s+regards|Kind\s+regards|With\s+regards|Regards|Sincerely(?:\s+yours)?)\b',
        bold_closing,
        html_content
    )

    # 3. Intelligent Paragraph & Spacing Auto-Readjustment:
    # Handles whatever raw text the user provides (double-spaced, single-spaced, irregular indents).
    raw_blocks = [b.strip() for b in re.split(r'\n\s*\n+', html_content) if b.strip()]
    salutation_block = []
    body_blocks = []
    closing_block = []
    mode = 'salutation'

    for block in raw_blocks:
        b_lines = [l.strip() for l in block.split('\n') if l.strip()]
        if not b_lines:
            continue

        # Split salutation if on first line
        if mode == 'salutation' and re.match(r'(?i)^(?:<b[^>]*>)?\s*(?:dear|hello|hi|greetings)\b', b_lines[0]):
            salutation_block = [b_lines[0]]
            b_lines = b_lines[1:]
            mode = 'body'
            if not b_lines:
                continue

        # Check if closing/sign-off begins in this block
        closing_idx = -1
        for idx, l in enumerate(b_lines):
            if re.search(r'(?i)\b(?:best\s+regards|warm\s+regards|kind\s+regards|with\s+regards|regards|sincerely)\b', l):
                closing_idx = idx
                break

        if closing_idx != -1:
            before_closing = b_lines[:closing_idx]
            after_closing = b_lines[closing_idx:]
            if before_closing:
                body_blocks.append(before_closing)
            closing_block.extend(after_closing)
            mode = 'closing'
            continue

        if mode == 'closing':
            closing_block.extend(b_lines)
        else:
            body_blocks.append(b_lines)

    # Reconstruct body paragraphs:
    # If the user separated thoughts by full blank lines, each block is its own paragraph.
    # Within each paragraph block, lines are combined with a space so word-wrapped text flows naturally without extra vertical line breaks.
    # If the user gave a single block where all sentences end in punctuation and are separate thoughts, break each line into its own paragraph.
    final_body_blocks = []
    for bb in body_blocks:
        if len(body_blocks) == 1 and len(bb) > 1 and all(len(l) > 20 and re.search(r'[\.\?!,;]$', l) for l in bb):
            for line in bb:
                final_body_blocks.append(line)
        else:
            # Join line-wrapped text within paragraph with space
            paragraph_text = " ".join(bb)
            final_body_blocks.append(paragraph_text)

    p_tags = []

    # 1. Salutation (always left-aligned with 0 top margin, 16px bottom margin)
    if salutation_block:
        s_inner = "<br>\n".join(salutation_block)
        p_tags.append(
            f'<p align="left" style="text-align: left !important; margin: 0 0 16px 0 !important; padding: 0 !important; font-family: \'Times New Roman\', Times, serif !important; font-size: 12pt !important; color: #000000 !important; line-height: 1.5 !important;">{s_inner}</p>'
        )

    # 2. Body Paragraphs (strictly justified with inter-word justification)
    for b_para in final_body_blocks:
        if not b_para.strip():
            continue
        p_tags.append(
            f'<p align="justify" style="text-align: justify !important; text-justify: inter-word !important; text-align-last: left !important; margin: 0 0 16px 0 !important; padding: 0 !important; font-family: \'Times New Roman\', Times, serif !important; font-size: 12pt !important; color: #000000 !important; line-height: 1.5 !important;">{b_para.strip()}</p>'
        )

    # 3. Closing & Sign-off Block (always left-aligned with a clean top space before Best regards)
    if closing_block:
        c_inner = "<br>\n".join(closing_block)
        p_tags.append(
            f'<p align="left" style="text-align: left !important; margin: 20px 0 0 0 !important; padding: 0 !important; font-family: \'Times New Roman\', Times, serif !important; font-size: 12pt !important; color: #000000 !important; line-height: 1.5 !important;">{c_inner}</p>'
        )

    rich_html = "\n".join(p_tags)

    return rich_html, plain_text

DEFAULT_PROFILE_NAME = "Instance 1 (Primary)"
AVAILABLE_INSTANCES = ["Instance 1", "Instance 2", "Instance 3"]

PROFILE_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "zoho_browser_profiles")
os.makedirs(PROFILE_BASE_DIR, exist_ok=True)

# Compatibility pointer for legacy default profile
LEGACY_PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "zoho_browser_profile")

def get_instance_profile_dir(instance_name: str) -> str:
    safe_slug = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in instance_name.strip()).lower()
    if not safe_slug:
        safe_slug = "instance_1"
    # If using instance 1 and old profile exists, preserve cookies from legacy folder
    target_dir = os.path.join(PROFILE_BASE_DIR, safe_slug)
    os.makedirs(target_dir, exist_ok=True)
    return target_dir

class ModernDarkZohoAutomationApp:
    def __init__(self, root: tk.Tk, initial_instance: str = "Instance 1"):
        self.root = root
        self.current_instance = initial_instance
        self.root.title(f"MimicMail — [{self.current_instance}] Zoho Outreach Automation")
        self.root.geometry("940x920")
        self.root.minsize(840, 800)

        # State flags
        self.is_running = False
        self.stop_requested = False
        self.excel_file_path = None
        self.total_rows_count = 0
        self.sent_rows_count = 0

        # Thread-safe work queue and background worker
        self.cmd_queue = queue.Queue()
        self.worker_thread = None

        self._apply_dark_theme()
        self._build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _ensure_worker_started(self):
        if not self.worker_thread or not self.worker_thread.is_alive():
            self.worker_thread = threading.Thread(target=self._dedicated_browser_thread, daemon=True)
            self.worker_thread.start()

    def _apply_dark_theme(self):
        # Clean, Prestigious Red, Cream, and White Palette
        self.bg_root = "#f8f5f0"        # Light Warm Cream canvas
        self.bg_card = "#ffffff"        # Pure Crisp White cards
        self.bg_card_alt = "#fff9f2"    # Soft Cream highlight
        self.border_card = "#e6ded1"    # Elegant Warm Grey/Cream border
        self.fg_primary = "#1f1d1a"     # Deep Charcoal/Soft Black for crystal legibility
        self.fg_muted = "#766d62"       # Muted Warm Grey
        self.accent_red = "#cf142b"     # Official Pulsus Crimson Red
        self.accent_red_hover = "#ad0d21"
        self.accent_red_light = "#fdf2f3"# Ultra-light Red Tint
        self.accent_red_border = "#e899a2"
        self.accent_cream = "#f4eee1"   # Warm Cream button surface

        self.root.configure(bg=self.bg_root)

        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.style.configure(".", background=self.bg_root, foreground=self.fg_primary)

        # Notebook & Tabs
        self.style.configure("Dark.TNotebook", background=self.bg_card, borderwidth=0)
        self.style.configure("Dark.TNotebook.Tab", background="#f4eee1", foreground="#5c5346", font=("Segoe UI", 9, "bold"), padding=[14, 5])
        self.style.map("Dark.TNotebook.Tab",
                       background=[("selected", "#cf142b"), ("active", "#e6ded1")],
                       foreground=[("selected", "#ffffff"), ("active", "#1f1d1a")])

        # LabelFrames
        self.style.configure("DarkCard.TLabelframe", background=self.bg_card, bordercolor=self.border_card, relief="solid", borderwidth=1)
        self.style.configure("DarkCard.TLabelframe.Label", background=self.bg_card, foreground=self.accent_red, font=("Segoe UI", 10, "bold"))

        # Frames
        self.style.configure("Root.TFrame", background=self.bg_root)
        self.style.configure("Card.TFrame", background=self.bg_card)

        # Labels
        self.style.configure("Dark.TLabel", background=self.bg_card, foreground=self.fg_primary, font=("Segoe UI", 9))
        self.style.configure("DarkMuted.TLabel", background=self.bg_card, foreground=self.fg_muted, font=("Segoe UI", 9))

        # Buttons - All unified royal blue palette as requested
        self.style.configure("StandardDark.TButton", font=("Segoe UI", 9, "bold"), padding=(12, 6),
                             background="#1a62c6", foreground="#ffffff", bordercolor="#124996", lightcolor="#2a75dc", darkcolor="#124996")
        self.style.map("StandardDark.TButton",
                       background=[("active", "#1550a2"), ("disabled", "#98b9e8")],
                       foreground=[("disabled", "#ffffff")])

        self.style.configure("ActionBlue.TButton", font=("Segoe UI", 9, "bold"), padding=(14, 6),
                             background="#1a62c6", foreground="#ffffff", bordercolor="#124996", lightcolor="#2a75dc", darkcolor="#124996")
        self.style.map("ActionBlue.TButton",
                       background=[("active", "#1550a2"), ("disabled", "#98b9e8")],
                       foreground=[("disabled", "#ffffff")])

        self.style.configure("ActionGreen.TButton", font=("Segoe UI", 13, "bold"), padding=(24, 12),
                             background="#1a62c6", foreground="#ffffff", bordercolor="#124996", lightcolor="#2a75dc", darkcolor="#124996")
        self.style.map("ActionGreen.TButton",
                       background=[("active", "#1550a2"), ("disabled", "#98b9e8")],
                       foreground=[("disabled", "#ffffff")])

        self.style.configure("DangerRed.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 12),
                             background="#2a3b5c", foreground="#ffffff", bordercolor="#124996", lightcolor="#3b4f76", darkcolor="#124996")
        self.style.map("DangerRed.TButton",
                       background=[("active", "#1c2840"), ("disabled", "#a0aec0")],
                       foreground=[("disabled", "#ffffff")])

        # Progressbar
        self.style.configure("Green.Horizontal.TProgressbar", troughcolor="#f4eee1", background="#1a62c6", bordercolor="#d5ccbe", lightcolor="#2a75dc", darkcolor="#1a62c6")

        # Sliders
        self.style.configure("Dark.Horizontal.TScale", background=self.bg_card, troughcolor="#f4eee1", bordercolor=self.border_card)

        # Checkbutton
        self.style.configure("Dark.TCheckbutton", background=self.bg_card, foreground=self.fg_primary, font=("Segoe UI", 9))
        self.style.map("Dark.TCheckbutton", background=[("active", self.bg_card)])

    def _build_ui(self):
        container = ttk.Frame(self.root, style="Root.TFrame", padding="16")
        container.pack(fill=tk.BOTH, expand=True)

        # =========================================================================
        # 1. TOP HEADER / BRANDING BANNER (Clean White & Crimson Accent)
        # =========================================================================
        header_card = tk.Frame(container, bg="#ffffff", highlightbackground=self.border_card, highlightthickness=1, padx=20, pady=12)
        header_card.pack(fill=tk.X, pady=(0, 12))

        header_row = tk.Frame(header_card, bg="#ffffff")
        header_row.pack(fill=tk.X)

        # Pulsus Logo
        self.logo_photo = None
        if os.path.exists(LOGO_PATH):
            try:
                pil_img = Image.open(LOGO_PATH)
                w_orig, h_orig = pil_img.size
                target_h = 44
                target_w = int(w_orig * (target_h / float(h_orig)))
                pil_resized = pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(pil_resized)
                logo_lbl = tk.Label(header_row, image=self.logo_photo, bg="#ffffff")
                logo_lbl.pack(side=tk.LEFT, padx=(0, 16))
            except Exception:
                pass

        title_box = tk.Frame(header_row, bg="#ffffff")
        title_box.pack(side=tk.LEFT, fill=tk.Y)

        title_lbl = tk.Label(title_box, text="MIMICMAIL  |  OUTREACH AUTOMATION", bg="#ffffff", fg=self.accent_red, font=("Segoe UI", 16, "bold"))
        title_lbl.pack(anchor=tk.W)

        sub_lbl = tk.Label(title_box, text="Pulsus Group Clinical Outreach & Playwright Delivery Engine", bg="#ffffff", fg=self.fg_muted, font=("Segoe UI", 9))
        sub_lbl.pack(anchor=tk.W)

        badge_box = tk.Frame(header_row, bg="#ffffff")
        badge_box.pack(side=tk.RIGHT)

        badge_content = tk.Frame(badge_box, bg="#fdf2f3", highlightbackground=self.accent_red_border, highlightthickness=1, padx=6, pady=3)
        badge_content.pack(side=tk.RIGHT)

        # 3D Live Glowing ECG Heartbeat Lifeline Monitor Canvas
        self.ecg_canvas = tk.Canvas(badge_content, width=86, height=26, bg="#1a0407", highlightthickness=1, highlightbackground="#cf142b")
        self.ecg_canvas.pack(side=tk.LEFT, padx=(2, 8))
        self.ecg_step = 0
        self._animate_ecg_heartbeat()

        badge_lbl = tk.Label(badge_content, text="OFFICIAL CLINICAL PLATFORM", bg="#fdf2f3", fg=self.accent_red,
                             font=("Segoe UI", 8, "bold"))
        badge_lbl.pack(side=tk.LEFT, padx=(0, 4))

        # =========================================================================
        # 2. TWO-COLUMN MAIN BODY (Left: Session & Templates | Right: Status & Actions)
        # =========================================================================
        main_columns = ttk.Frame(container, style="Root.TFrame")
        main_columns.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # --- LEFT COLUMN (Config & Templates, Width ~62%) ---
        left_col = ttk.Frame(main_columns, style="Root.TFrame")
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        # Card 1: Webmail Session & Spreadsheet Ingestion (Top action buttons parallel, info below)
        setup_card = ttk.LabelFrame(left_col, text="  1. Session & Recipient Registry  ", style="DarkCard.TLabelframe", padding="12")
        setup_card.pack(fill=tk.X, pady=(0, 10))

        # Row 0: Multi-Instance Profile Selector & New Instance Launcher
        inst_row = ttk.Frame(setup_card, style="Card.TFrame")
        inst_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(inst_row, text="👤 Profile / Instance:", style="Dark.TLabel").pack(side=tk.LEFT)
        self.instance_var = tk.StringVar(value=self.current_instance)
        self.instance_combo = ttk.Combobox(inst_row, textvariable=self.instance_var, values=AVAILABLE_INSTANCES, state="readonly", width=14)
        self.instance_combo.pack(side=tk.LEFT, padx=(6, 10))
        self.instance_combo.bind("<<ComboboxSelected>>", self._on_instance_changed)

        self.new_window_btn = ttk.Button(inst_row, text="➕ Launch New Instance", command=self.launch_new_app_instance, style="StandardDark.TButton")
        self.new_window_btn.pack(side=tk.RIGHT)

        ttk.Label(inst_row, text="(Multi-Instance Concurrency: Up to 3 Instances)", style="DarkMuted.TLabel").pack(side=tk.RIGHT, padx=6)

        # Top Button Row: Launch Zoho on left corner, Browse Excel on right corner
        btn_row = ttk.Frame(setup_card, style="Card.TFrame")
        btn_row.pack(fill=tk.X, pady=(0, 6))

        self.login_btn = ttk.Button(btn_row, text="🌐 Launch Zoho Webmail", command=self.request_open_browser, style="ActionBlue.TButton")
        self.login_btn.pack(side=tk.LEFT)

        self.browse_btn = ttk.Button(btn_row, text="📁 Browse Excel (.xlsx)", command=self.browse_excel, style="ActionBlue.TButton")
        self.browse_btn.pack(side=tk.RIGHT)

        # Bottom Info Row: Session indicator on left, loaded file status on right
        info_row = ttk.Frame(setup_card, style="Card.TFrame")
        info_row.pack(fill=tk.X)

        self.session_indicator = tk.Label(info_row, text=f"● Browser [{self.current_instance}]: Ready", bg="#ffffff", fg=self.fg_muted, font=("Segoe UI", 9, "bold"))
        self.session_indicator.pack(side=tk.LEFT)

        self.file_label = ttk.Label(info_row, text="📄 No registry loaded (Required: email, author name, title)", style="DarkMuted.TLabel")
        self.file_label.pack(side=tk.RIGHT)

        # Card 2: 5 HTML Template Slots
        tmpl_card = ttk.LabelFrame(left_col, text="  2. Custom HTML Templates (5 Slots - Random Selection)  ", style="DarkCard.TLabelframe", padding="12")
        tmpl_card.pack(fill=tk.BOTH, expand=True)

        subj_box = ttk.Frame(tmpl_card, style="Card.TFrame")
        subj_box.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(subj_box, text="Subject Line Template:", style="Dark.TLabel").pack(side=tk.LEFT)
        ttk.Label(subj_box, text="(Uses {title} per recipient)", style="DarkMuted.TLabel").pack(side=tk.LEFT, padx=6)

        self.subj_entry = tk.Entry(subj_box, bg="#ffffff", fg="#1f1d1a", insertbackground="#cf142b",
                                   highlightbackground=self.border_card, highlightcolor=self.accent_red, highlightthickness=1,
                                   relief="flat", font=("Segoe UI", 10))
        self.subj_entry.insert(0, DEFAULT_SUBJECT)
        self.subj_entry.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(8, 0), ipady=3)

        # Notebook tabs for 5 templates
        nb_header = ttk.Frame(tmpl_card, style="Card.TFrame")
        nb_header.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(nb_header, text="Paste or load email message templates below (Picks randomly per recipient):", style="DarkMuted.TLabel").pack(side=tk.LEFT)
        tk.Label(nb_header, text="Times New Roman | Bold {name} & {title}", bg="#fdf2f3", fg=self.accent_red, font=("Segoe UI", 8, "bold"), padx=6, pady=2).pack(side=tk.RIGHT)

        self.template_notebook = ttk.Notebook(tmpl_card, style="Dark.TNotebook")
        self.template_notebook.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        self.template_text_widgets = []
        for i in range(3):
            tab_frame = ttk.Frame(self.template_notebook, style="Card.TFrame")
            self.template_notebook.add(tab_frame, text=f" Template {i+1} ")

            tb = ttk.Frame(tab_frame, style="Card.TFrame")
            tb.pack(fill=tk.X, pady=(4, 4))
            ttk.Label(tb, text=f"Slot {i+1} message text:", style="DarkMuted.TLabel").pack(side=tk.LEFT)

            load_btn = ttk.Button(tb, text=f"📂 Load File", command=lambda idx=i: self._load_template_file(idx), style="StandardDark.TButton")
            load_btn.pack(side=tk.RIGHT)

            clear_btn = ttk.Button(tb, text="🗑 Clear", command=lambda idx=i: self._clear_template_slot(idx), style="StandardDark.TButton")
            clear_btn.pack(side=tk.RIGHT, padx=4)

            times_font = tkfont.Font(family="Times New Roman", size=11)
            txt = tk.Text(tab_frame, height=6, wrap=tk.WORD, bg="#ffffff", fg="#1f1d1a",
                          insertbackground="#cf142b", highlightbackground=self.border_card,
                          highlightcolor=self.accent_red, highlightthickness=1, relief="flat", font=times_font)
            txt.configure(font=times_font)
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

        ttk.Label(pacing_box, text="⏳ Rest Interval:", style="Dark.TLabel").pack(anchor=tk.W)
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

        # Multi-tab concurrency setting (Optimized for 4GB - 8GB RAM)
        tab_row = ttk.Frame(pacing_box, style="Card.TFrame")
        tab_row.pack(fill=tk.X, pady=(6, 2))
        ttk.Label(tab_row, text="⚡ Parallel Background Tabs:", style="Dark.TLabel").pack(side=tk.LEFT)
        self.tabs_var = tk.IntVar(value=1)
        self.tabs_spin = ttk.Spinbox(tab_row, from_=1, to=2, textvariable=self.tabs_var, width=5, state="readonly")
        self.tabs_spin.pack(side=tk.RIGHT)
        ttk.Label(pacing_box, text="(Hardware-safe concurrency: 1 or 2 tabs)", style="DarkMuted.TLabel").pack(anchor=tk.W)

        # Section: LIVE STATUS & GRAPH
        status_card = ttk.LabelFrame(right_col, text="  4. Delivery Status & Graph  ", style="DarkCard.TLabelframe", padding="14")
        status_card.pack(fill=tk.BOTH, expand=True)

        self.timer_display = tk.Label(status_card, text="⏱ Queue: Ready", bg="#ffffff", fg=self.accent_red, font=("Segoe UI", 10, "bold"))
        self.timer_display.pack(anchor=tk.W, pady=(0, 4))

        self.stats_lbl = ttk.Label(status_card, text="Total: 0  |  Sent: 0  |  Pending: 0", style="DarkMuted.TLabel")
        self.stats_lbl.pack(anchor=tk.W, pady=(0, 6))

        # Loading Progress Bar
        pbar_box = ttk.Frame(status_card, style="Card.TFrame")
        pbar_box.pack(fill=tk.X, pady=(0, 6))

        self.progress_bar = ttk.Progressbar(pbar_box, style="Green.Horizontal.TProgressbar", mode="determinate", maximum=100, value=0)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.progress_pct_lbl = tk.Label(pbar_box, text="0%", bg="#ffffff", fg=self.accent_red, font=("Segoe UI", 10, "bold"), width=5)
        self.progress_pct_lbl.pack(side=tk.RIGHT)

        # Status Graph Canvas
        ttk.Label(status_card, text="Live Ratio Breakdown:", style="Dark.TLabel").pack(anchor=tk.W, pady=(4, 2))
        self.graph_canvas = tk.Canvas(status_card, bg="#f8f5f0", height=38, highlightthickness=1, highlightbackground=self.border_card)
        self.graph_canvas.pack(fill=tk.X, pady=(0, 4))
        self.graph_canvas.bind("<Configure>", lambda e: self._draw_status_graph())

        self.graph_legend_lbl = tk.Label(status_card, text="🟥 Sent: 0  |  ⬜ Pending: 0  |  ⚠️ Errors: 0",
                                         bg="#ffffff", fg=self.fg_muted, font=("Segoe UI", 8), justify=tk.LEFT)
        self.graph_legend_lbl.pack(anchor=tk.W)

        # =========================================================================
        # 3. LIVE TERMINAL LOG AT THE BOTTOM (Clean Cream Background)
        # =========================================================================
        log_frame = ttk.LabelFrame(container, text="  Live Outreach Audit Terminal  ", style="DarkCard.TLabelframe", padding="8")
        log_frame.pack(fill=tk.X)

        self.log_text = tk.Text(log_frame, height=4, state=tk.DISABLED, wrap=tk.WORD, bg="#ffffff", fg="#2b261f",
                                insertbackground="#cf142b", relief="flat", font=("Consolas", 8),
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
            self.graph_canvas.create_rectangle(0, 0, w, h, fill="#f4eee1", outline="")
            self.graph_canvas.create_text(w // 2, h // 2, text="Load spreadsheet to view graph", fill="#8a8070", font=("Segoe UI", 8, "italic"))
            return

        sent_pct = (sent / total)
        pending_pct = (pending / total)
        err_pct = (errors / total)

        sent_w = int(w * sent_pct)
        pending_w = int(w * pending_pct)
        err_w = w - sent_w - pending_w

        # Draw segmented colored bar in Red, Cream, and Grey
        x_curr = 0
        if sent_w > 0:
            # Crimson Red for Sent
            self.graph_canvas.create_rectangle(x_curr, 0, x_curr + sent_w, h, fill="#cf142b", outline="")
            if sent_w > 40:
                self.graph_canvas.create_text(x_curr + sent_w // 2, h // 2, text=f"{sent}", fill="#ffffff", font=("Segoe UI", 8, "bold"))
            x_curr += sent_w

        if pending_w > 0:
            # Warm Cream for Pending
            self.graph_canvas.create_rectangle(x_curr, 0, x_curr + pending_w, h, fill="#f4eee1", outline="")
            if pending_w > 40:
                self.graph_canvas.create_text(x_curr + pending_w // 2, h // 2, text=f"{pending}", fill="#2b261f", font=("Segoe UI", 8, "bold"))
            x_curr += pending_w

        if err_w > 0:
            # Dark Slate for Errors/Skipped
            self.graph_canvas.create_rectangle(x_curr, 0, w, h, fill="#766d62", outline="")
            if err_w > 30:
                self.graph_canvas.create_text(x_curr + err_w // 2, h // 2, text=f"{errors}", fill="#ffffff", font=("Segoe UI", 8, "bold"))

        legend_text = f"🟥 Sent: {sent} ({sent_pct*100:.0f}%)   |   ⬜ Pending: {pending} ({pending_pct*100:.0f}%)   |   ⚠️ Errors: {errors}"
        self.graph_legend_lbl.config(text=legend_text)

    def _animate_ecg_heartbeat(self):
        """Draws a live, animated 3D-styled medical ECG heartbeat pulse wave with a glowing lead dot."""
        if not hasattr(self, "ecg_canvas") or not self.ecg_canvas.winfo_exists():
            return

        w = 86
        h = 26
        mid_y = h / 2.0

        self.ecg_canvas.delete("all")

        # 1. Subtle 3D background grid lines (medical monitor style)
        for gy in [6, 13, 20]:
            self.ecg_canvas.create_line(0, gy, w, gy, fill="#2b0d13", width=1)
        for gx in range(10, w, 15):
            self.ecg_canvas.create_line(gx, 0, gx, h, fill="#2b0d13", width=1)

        # 2. Wave pattern template (normalized y-offsets around center)
        # ECG rhythm: baseline -> small P wave -> Q dip -> sharp R spike -> S dip -> T wave -> baseline
        pattern = [0, 0, 0, -2, -1, 0, 0, 2, -10, 6, -1, 0, -3, -4, -2, 0, 0, 0, 0, 0]
        pat_len = len(pattern)

        points = []
        for x in range(0, w, 2):
            idx = int((x + self.ecg_step * 2) / 3) % pat_len
            y = mid_y + pattern[idx]
            points.append((x, y))

        # 3. Draw 3D shadow/ambient glow trace (deeper crimson)
        for i in range(len(points) - 1):
            x1, y1 = points[i]
            x2, y2 = points[i + 1]
            self.ecg_canvas.create_line(x1, y1 + 1.2, x2, y2 + 1.2, fill="#700816", width=2.5)

        # 4. Draw bright neon primary pulse line
        for i in range(len(points) - 1):
            x1, y1 = points[i]
            x2, y2 = points[i + 1]
            self.ecg_canvas.create_line(x1, y1, x2, y2, fill="#ff2a44", width=1.5)

        # 5. Glowing 3D lead pulse dot that travels across the monitor
        lead_x = (self.ecg_step * 3) % w
        lead_idx = int((lead_x + self.ecg_step * 2) / 3) % pat_len
        lead_y = mid_y + pattern[lead_idx]

        # Outer soft glow
        self.ecg_canvas.create_oval(lead_x - 4, lead_y - 4, lead_x + 4, lead_y + 4, fill="#a80b1e", outline="")
        # Bright 3D highlight core
        self.ecg_canvas.create_oval(lead_x - 2, lead_y - 2, lead_x + 2, lead_y + 2, fill="#ffffff", outline="#ff4d64")

        self.ecg_step += 1
        # Schedule next frame at ~30 FPS for smooth live cardiac monitoring
        self.root.after(35, self._animate_ecg_heartbeat)

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
        def _append():
            self.log_text.config(state=tk.NORMAL)
            t = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{t}] {text}\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        self.root.after(0, _append)

    def _on_instance_changed(self, event=None):
        new_inst = self.instance_var.get()
        if new_inst != self.current_instance:
            if self.is_running:
                messagebox.showwarning("Instance Busy", "Please stop the active automation before switching profiles in this window.")
                self.instance_var.set(self.current_instance)
                return
            self.current_instance = new_inst
            self.root.title(f"MimicMail — [{self.current_instance}] Zoho Outreach Automation")
            self.session_indicator.config(text=f"● Browser [{self.current_instance}]: Ready", fg=self.fg_muted)
            self.log(f"Switched active profile to: {self.current_instance} (Using isolated cookies & storage)")
            # If browser was already running, signal worker to reload context
            self.cmd_queue.put(("SWITCH_PROFILE", self.current_instance))

    def launch_new_app_instance(self):
        """Spawns the next available instance (Instance 1 -> Instance 2 -> Instance 3) in a separate concurrent process."""
        import subprocess
        # Cycle through AVAILABLE_INSTANCES (Instance 1 -> 2 -> 3 -> 1)
        try:
            curr_idx = AVAILABLE_INSTANCES.index(self.current_instance)
            next_inst = AVAILABLE_INSTANCES[(curr_idx + 1) % len(AVAILABLE_INSTANCES)]
        except ValueError:
            next_inst = AVAILABLE_INSTANCES[0]

        script_path = os.path.abspath(__file__)
        python_exe = sys.executable
        self.log(f"Spawning concurrent process for: {next_inst}...")
        try:
            subprocess.Popen([python_exe, script_path, "--instance", next_inst])
            self.log(f"✓ Concurrent MimicMail window launched for {next_inst}!")
        except Exception as e:
            messagebox.showerror("Launch Error", f"Unable to spawn instance:\n{str(e)}")

    def request_open_browser(self):
        self._ensure_worker_started()
        self.cmd_queue.put(("OPEN_BROWSER", None))

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

            # Validation 2: Check required columns (email, author name, title)
            col_map = {str(c).strip().lower(): c for c in df.columns}
            
            author_col = next((col_map[k] for k in ["author name", "author_name", "author", "name"] if k in col_map), None)
            email_col = next((col_map[k] for k in ["email", "email address", "e-mail"] if k in col_map), None)
            title_col = next((col_map[k] for k in ["title", "designation", "role"] if k in col_map), None)

            missing = []
            if not email_col:
                missing.append("email")
            if not author_col:
                missing.append("author name")
            if not title_col:
                missing.append("title")

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
                    f"• title (Recipient's clinical role or subject)\n\n"
                    f"Note: No 'Status' column is required; dispatch tracking is handled automatically."
                )
                self.log(f"⚠️ Warning: '{os.path.basename(path)}' is missing required columns: {missing_str}")
                return

            # Validation 3: Check for empty essential values
            name_col = author_col
            email_col = col_map["email"]
            empty_emails = df[email_col].isna().sum() + (df[email_col].astype(str).str.strip() == "").sum()
            
            if empty_emails > 0:
                messagebox.showwarning(
                    "⚠️ Excel Warning — Blank Emails Detected",
                    f"Note: Found {empty_emails} row(s) with missing or blank email addresses.\n"
                    f"These will be flagged or skipped during outreach."
                )
                self.log(f"⚠️ Note: {empty_emails} rows have missing email values.")

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

    def request_start_automation(self):
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            messagebox.showwarning("File Missing", "Please select an Excel workbook first.")
            return

        subj_tmpl = self.subj_entry.get().strip()
        if not subj_tmpl:
            messagebox.showwarning("Missing Subject", "Please enter a subject line template (e.g. {title}).")
            return

        # Gather all non-empty user templates across the 3 slots
        active_templates = []
        for i, txt_widget in enumerate(self.template_text_widgets):
            content = txt_widget.get("1.0", tk.END).strip()
            if content:
                active_templates.append(content)

        if not active_templates:
            messagebox.showwarning("No Templates", "Please provide at least 1 template in the 3 slots (paste HTML or load a file).")
            return

        self.log(f"Verified {len(active_templates)} active user template(s). Random picker enabled across them.")

        self.is_running = True
        self.stop_requested = False
        self._toggle_ui(running=True)

        params = {
            "subj_tmpl": subj_tmpl,
            "body_tmpls": active_templates,
            "base_mins": self.interval_var.get(),
            "use_jitter": self.jitter_var.get(),
            "excel_path": self.excel_file_path,
            "parallel_tabs": max(1, self.tabs_var.get())
        }
        self._ensure_worker_started()
        self.cmd_queue.put(("START_AUTOMATION", params))

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
    # SINGLE DEDICATED PLAYWRIGHT WORKER THREAD
    # All Playwright calls stay inside this thread forever, preventing Greenlet errors!
    # =========================================================================
    def _dedicated_browser_thread(self):
        playwright_instance = None
        browser_context = None
        browser_page = None

        try:
            playwright_instance = sync_playwright().start()

            while True:
                try:
                    cmd, data = self.cmd_queue.get(timeout=0.2)
                except queue.Empty:
                    continue

                if cmd == "QUIT":
                    break

                elif cmd == "SWITCH_PROFILE":
                    try:
                        if browser_context:
                            self.log("Closing browser context for profile switch...")
                            browser_context.close()
                            browser_context = None
                            browser_page = None
                            self.root.after(0, lambda: self.session_indicator.config(text=f"● Browser [{self.current_instance}]: Ready", fg=self.fg_muted))
                    except Exception as err:
                        self.log(f"Profile switch warning: {err}")

                elif cmd == "OPEN_BROWSER":
                    try:
                        active_profile_dir = get_instance_profile_dir(self.current_instance)
                        # Clean up any stale singleton locks if browser process died uncleanly
                        for lock_name in ["SingletonLock", "SingletonSocket", "SingletonCookie", "lockfile"]:
                            lock_path = os.path.join(active_profile_dir, lock_name)
                            if os.path.exists(lock_path):
                                try:
                                    os.remove(lock_path)
                                except Exception:
                                    pass

                        if not (browser_context and browser_page and not browser_page.is_closed()):
                            self.log(f"Opening persistent Chromium browser for {self.current_instance}...")
                            self.root.after(0, lambda: self.session_indicator.config(text="● Launching...", fg="#e67e22"))

                            browser_context = playwright_instance.chromium.launch_persistent_context(
                                user_data_dir=active_profile_dir,
                                headless=False,
                                no_viewport=True,
                                timeout=45000,
                                permissions=["clipboard-read", "clipboard-write"],
                                args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
                            )
                            browser_page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()
                            try:
                                browser_page.goto("https://mail.zoho.com", timeout=45000)
                            except Exception as nav_e:
                                self.log(f"Navigation warning (will proceed): {nav_e}")

                            self.log(f"Browser for {self.current_instance} is open! Log in and leave this window open.")
                            self.root.after(0, lambda: self.session_indicator.config(text=f"● Active & Connected [{self.current_instance}]", fg=self.accent_red))
                        else:
                            self.log("Browser already open. Bringing tab to front...")
                            browser_page.bring_to_front()
                    except Exception as err:
                        err_msg = str(err)
                        self.log(f"Open browser error: {err_msg}")
                        if "ProcessSingleton" in err_msg or "Target page, context or browser has been closed" in err_msg or "locked" in err_msg.lower():
                            self.log("💡 Tip: A browser window for this instance is already open. Close existing Chromium windows and retry.")
                        self.root.after(0, lambda: self.session_indicator.config(text="● Open Error (Check Logs)", fg=self.accent_red))
                        self.root.after(0, lambda m=err_msg: messagebox.showerror("Browser Launch Error", f"Unable to launch browser:\n\n{m[:200]}\n\nIf another browser window is already using this profile, please close it and try again."))

                elif cmd == "START_AUTOMATION":
                    try:
                        active_profile_dir = get_instance_profile_dir(self.current_instance)
                        # 1. Ensure browser is open in this same thread
                        if not (browser_context and browser_page and not browser_page.is_closed()):
                            self.log(f"Opening browser session for {self.current_instance}...")
                            browser_context = playwright_instance.chromium.launch_persistent_context(
                                user_data_dir=active_profile_dir,
                                headless=False,
                                permissions=["clipboard-read", "clipboard-write"],
                                args=["--start-maximized"]
                            )
                            browser_page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()

                        browser_page.bring_to_front()

                        # Ensure we are on the actual Zoho Mail client (/zm/), not accounts or calendar or landing page
                        curr_url = browser_page.url.lower()
                        self.log(f"Current page URL: {curr_url}")

                        # Check if "Access Email" button or "Mail" tile exists on screen (Zoho home portal)
                        for access_btn_sel in [
                            "a:has-text('Access Email')",
                            "button:has-text('Access Email')",
                            "a:has-text('Access Mail')",
                            "[title*='Access Email']",
                            "a[href*='mail.zoho']"
                        ]:
                            try:
                                btn = browser_page.query_selector(access_btn_sel)
                                if btn and btn.is_visible():
                                    self.log("Found 'Access Email' button on page. Clicking to navigate into Mail...")
                                    btn.click()
                                    time.sleep(4)
                                    break
                            except Exception:
                                pass

                        # If not inside the mailbox app, directly navigate to https://mail.zoho.com/zm/
                        if "mail.zoho" not in browser_page.url.lower() or "/zm" not in browser_page.url.lower():
                            self.log("Navigating directly to Zoho Mail Inbox: https://mail.zoho.com/zm/ ...")
                            browser_page.goto("https://mail.zoho.com/zm/")
                            time.sleep(5)

                        # Make sure Mail icon is selected in the left navigation sidebar (not Calendar/Events)
                        try:
                            mail_app_icon = browser_page.query_selector(".zmNavMail, [data-app='mail'], [title='Mail']")
                            if mail_app_icon and mail_app_icon.is_visible():
                                mail_app_icon.click()
                                time.sleep(1.5)
                        except Exception:
                            pass

                        self.root.after(0, lambda: self.session_indicator.config(text="● In Mailbox & Automating", fg=self.accent_red))

                        # 2. Run batch dispatch across single or multiple background tabs
                        self._execute_batch(browser_context, browser_page, data)

                    except Exception as batch_err:
                        self.log(f"Automation execution error: {batch_err}")
                        self.root.after(0, lambda e=str(batch_err): messagebox.showerror("Automation Error", e))
                    finally:
                        self.is_running = False
                        self.root.after(0, lambda: self.timer_display.config(text="⏱ Queue: Idle", fg=self.fg_muted))
                        self.root.after(0, lambda: self._toggle_ui(running=False))

        except Exception as fatal:
            self.log(f"Fatal worker error: {fatal}")
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

    def _execute_batch(self, context, primary_page, params):
        excel_path = params["excel_path"]
        subj_tmpl = params["subj_tmpl"]
        body_tmpls = params.get("body_tmpls", [params.get("body_tmpl", "")])
        base_mins = params["base_mins"]
        use_jitter = params["use_jitter"]
        parallel_tabs = params.get("parallel_tabs", 1)

        df = pd.read_excel(excel_path)
        col_map = {str(c).strip().lower(): c for c in df.columns}
        name_col = next((col_map[k] for k in ["author name", "author_name", "author", "name"] if k in col_map), "name")
        email_col = next((col_map[k] for k in ["email", "email address", "e-mail"] if k in col_map), "email")
        title_col = next((col_map[k] for k in ["title", "designation", "role"] if k in col_map), "title")

        total_records = len(df)
        has_status_col = any(c.strip().lower() == "status" for c in df.columns)
        if has_status_col:
            status_col_name = next(c for c in df.columns if c.strip().lower() == "status")
            status_series = df[status_col_name].astype(str).str.strip().str.upper()
            pending_indices = [i for i, val in enumerate(status_series) if val != "SENT"]
            prev_sent_count = (status_series == "SENT").sum()
        else:
            pending_indices = list(range(total_records))
            prev_sent_count = 0

        if not pending_indices:
            self.log("All rows in this spreadsheet are already marked as SENT.")
            self.root.after(0, lambda: messagebox.showinfo("Completed", "All recipients in this spreadsheet are already sent!"))
            return

        BATCH_CAP = 40
        target_indices = pending_indices[:BATCH_CAP]
        total_in_batch = len(target_indices)

        self.log(f"Dispatching batch of {total_in_batch} recipients across {parallel_tabs} parallel tab(s) with {len(body_tmpls)} template pool...")

        sent_count = 0
        errors_in_session = 0

        # Create/manage pool of tabs inside this authenticated session
        tab_pool = [primary_page]
        if parallel_tabs > 1:
            self.log(f"Spinning up {parallel_tabs - 1} extra background tab(s) for parallel processing...")
            for t_idx in range(parallel_tabs - 1):
                try:
                    new_tab = context.new_page()
                    new_tab.goto("https://mail.zoho.com/zm/")
                    time.sleep(2)
                    tab_pool.append(new_tab)
                except Exception as t_err:
                    self.log(f"Tab creation notice: {t_err}")

        # Chunk the recipients into parallel slices according to active tab count
        chunk_size = max(1, len(tab_pool))
        for chunk_start in range(0, total_in_batch, chunk_size):
            if self.stop_requested:
                self.log("Automation interrupted by user.")
                break

            current_chunk = target_indices[chunk_start:chunk_start + chunk_size]

            for tab_idx, r_idx in enumerate(current_chunk):
                if self.stop_requested:
                    break

                active_page = tab_pool[tab_idx % len(tab_pool)]
                row = df.loc[r_idx]
                r_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else "Valued Professional"
                r_email = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
                r_title = str(row[title_col]).strip() if pd.notna(row[title_col]) else "Professional"

                import re
                EMAIL_REGEX = re.compile(r"^[\w\.-]+@[\w\.-]+\.\w+$")
                if not r_email or not EMAIL_REGEX.match(r_email):
                    self.log(f"Row {r_idx + 1}: SKIPPED (Invalid email address '{r_email}')")
                    errors_in_session += 1
                    continue

                chosen_body_tmpl = random.choice(body_tmpls)

                personalized_subject = subj_tmpl
                for ph, val in [("{title}", r_title), ("{name}", r_name), ("{TITLE}", r_title), ("{NAME}", r_name)]:
                    personalized_subject = personalized_subject.replace(ph, val)

                # Generate formatted Times New Roman HTML with bolded {name} and {title}, plus clean plain text fallback
                rich_body_html, plain_body_text = build_formatted_email_content(
                    raw_template=chosen_body_tmpl,
                    r_name=r_name,
                    r_title=r_title,
                    r_email=r_email
                )

                tab_label = f"Tab {tab_idx + 1}" if len(tab_pool) > 1 else "Primary"
                self.log(f"[{tab_label}] Dispatching ({chunk_start + tab_idx + 1}/{total_in_batch}) to {r_name} <{r_email}>...")

                try:
                    self._send_in_browser(active_page, r_email, personalized_subject, rich_body_html, plain_body_text)
                    sent_count += 1
                    self.log(f"SUCCESS [{tab_label}]: Email dispatched to {r_email}.")
                except Exception as err:
                    errors_in_session += 1
                    self.log(f"SKIPPED ROW {r_idx + 1} ({r_email}) due to error: {err}")
                    try:
                        active_page.keyboard.press("Escape")
                        time.sleep(1)
                    except Exception:
                        pass

                # Update live stats and status graph purely in UI memory
                curr_sent = prev_sent_count + sent_count
                curr_errors = errors_in_session
                curr_pending = max(0, total_records - curr_sent - curr_errors)
                curr_pct = int((curr_sent / total_records * 100)) if total_records > 0 else 0

                self.root.after(0, lambda s=curr_sent, t=total_records, p=curr_pending, e=curr_errors, pct=curr_pct: (
                    self.stats_lbl.config(text=f"Total: {t} | Sent: {s} | Pending: {p}"),
                    self.progress_bar.config(value=pct),
                    self.progress_pct_lbl.config(text=f"{pct}%"),
                    self._draw_status_graph(sent=s, pending=p, errors=e)
                ))

            # Anti-bot human delay between dispatch waves
            if chunk_start + chunk_size < total_in_batch and not self.stop_requested:
                base_sec = base_mins * 60.0
                jitter_sec = random.randint(-25, 35) if use_jitter else 0
                sleep_sec = max(20.0, base_sec + jitter_sec)

                self.log(f"Pacing: Resting for {sleep_sec:.0f} seconds before next dispatch wave...")

                remaining = int(sleep_sec)
                while remaining > 0 and not self.stop_requested:
                    m, s = divmod(remaining, 60)
                    self.root.after(0, lambda mm=m, ss=s: self.timer_display.config(
                        text=f"⏱ Next dispatch: {mm:02d}:{ss:02d}", fg=self.accent_red
                    ))
                    time.sleep(1)
                    remaining -= 1

        # Clean up any auxiliary tabs created
        if len(tab_pool) > 1:
            for extra_tab in tab_pool[1:]:
                try:
                    extra_tab.close()
                except Exception:
                    pass

        self.log(f"Batch completed! Total sent: {sent_count}")
        self.root.after(0, lambda s=sent_count: messagebox.showinfo("Done", f"Automation finished!\nTotal sent: {s}"))

    def _send_in_browser(self, page, to_email, subject, body_html, body_text=None):
        if body_text is None:
            body_text = clean_to_plain_text(body_html)
        self.log(f"Opening compose tab for: {to_email}")

        # 1. Click Compose / New Mail button (Strictly Mail, avoid Calendar / Reminders / Tasks!)
        clicked = False
        compose_selectors = [
            "button[title='New Mail']",
            "button[title*='New Mail']",
            "button.zmBtn_primary[title*='New Mail']",
            "button:has-text('New Mail')",
            "[data-test-id='compose-mail']",
            "button[aria-label='New Mail']",
            "button:has-text('Compose Mail')",
            "button:has-text('Compose')"
        ]
        for sel in compose_selectors:
            try:
                btn = page.query_selector(sel)
                if btn and btn.is_visible():
                    btn.click()
                    clicked = True
                    self.log(f"Clicked compose button via: {sel}")
                    break
            except Exception:
                pass

        if not clicked:
            # Dismiss any popups or reminder cards on screen first by pressing Escape
            page.keyboard.press("Escape")
            time.sleep(0.5)
            self.log("Using shortcut 'm' (or 'c') to open Mail composer...")
            page.keyboard.press("m")
            time.sleep(1)
            # Check if compose window appeared; if not try 'c'
            if not page.query_selector("input[placeholder*='Subject' i], #mailSubject"):
                page.keyboard.press("c")

        time.sleep(2)

        # 2. Locate and fill 'To' recipient field
        self.log(f"Entering recipient email: {to_email}")
        filled_to = False
        to_selectors = [
            "input[placeholder*='Add contact']",
            "input[placeholder*='to' i]",
            "div[data-name='to'] input",
            "#mailTo input",
            ".select2-search__field",
            "input[aria-label*='To' i]",
            "div[contenteditable='true'][aria-label*='To' i]"
        ]
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
            # Try finding any visible input in the top header of composer
            try:
                inputs = page.query_selector_all("input")
                for inp in inputs:
                    ph = inp.get_attribute("placeholder") or ""
                    if any(k in ph.lower() for k in ["contact", "recipient", "to", "email"]) and inp.is_visible():
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
        self.log(f"Entering subject: {subject}")
        filled_subj = False
        subj_selectors = [
            "input[placeholder*='Subject' i]",
            "input[aria-label*='Subject' i]",
            "input[name='subject']",
            "#mailSubject"
        ]
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
            # Fallback by tab or attribute scan
            try:
                inputs = page.query_selector_all("input")
                for inp in inputs:
                    ph = inp.get_attribute("placeholder") or ""
                    if "subject" in ph.lower() and inp.is_visible():
                        inp.click()
                        inp.fill(subject)
                        filled_subj = True
                        break
            except Exception:
                pass

        time.sleep(0.8)

        # 4. Locate and fill Message Body strictly formatted in Times New Roman with bolded {name} and {title}
        self.log(f"Inserting formatted email body in Times New Roman with bold name & title...")
        filled_body = False

        # Try rich clipboard paste / insertHTML inside iframe first
        for frame in page.frames:
            try:
                body_elem = frame.query_selector("body[contenteditable='true'], body")
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
                                document.body.style.removeProperty('text-align');
                                document.body.style.setProperty('font-family', "'Times New Roman', Times, serif", 'important');
                                document.body.style.setProperty('font-size', '12pt', 'important');
                                document.body.style.setProperty('color', '#000000', 'important');
                                document.body.style.setProperty('line-height', '1.6', 'important');
                                document.body.style.setProperty('margin', '0', 'important');
                                document.body.style.setProperty('padding', '0', 'important');

                                const sel = window.getSelection();
                                const range = document.createRange();
                                range.selectNodeContents(document.body);
                                sel.removeAllRanges();
                                sel.addRange(range);

                                // Try standard insertHTML
                                let ok = document.execCommand('insertHTML', false, html);
                                if (!ok || !document.body.innerHTML.trim() || document.body.innerText.trim().length === 0) {
                                    document.body.innerHTML = html;
                                }

                                // Remove any unwanted leading empty elements/br that create blank lines before Dear Dr
                                while (document.body.firstChild) {
                                    const fc = document.body.firstChild;
                                    const tag = (fc.nodeName || '').toLowerCase();
                                    const txt = (fc.textContent || '').replace(/[\\u200B-\\u200D\\uFEFF\\s]/g, '');
                                    if (tag === 'br' || (['p', 'div'].includes(tag) && txt === '')) {
                                        document.body.removeChild(fc);
                                    } else {
                                        break;
                                    }
                                }

                                // Apply styling across p, div, and block containers (skip wrapper elements with children)
                                const blocks = document.body.querySelectorAll('p, div');
                                blocks.forEach(el => {
                                    if (el === document.body) return;
                                    // Skip container divs that already hold p or div children to avoid double margin cascading
                                    if (el.querySelector('p, div')) return;

                                    const text = (el.innerText || el.textContent || '').trim();
                                    if (!text) return;

                                    el.style.setProperty('padding', '0', 'important');
                                    el.style.setProperty('line-height', '1.5', 'important');
                                    el.style.setProperty('font-family', "'Times New Roman', Times, serif", 'important');
                                    el.style.setProperty('font-size', '12pt', 'important');

                                    const firstLine = text.split('\\n')[0].trim();
                                    const isSalutation = /^(?:dear|hello|hi|greetings)/i.test(firstLine);
                                    const isClosing = /^(?:best\\s+regards|warm\\s+regards|kind\\s+regards|with\\s+regards|regards|sincerely)/i.test(firstLine) ||
                                                      /(?:best\\s+regards|warm\\s+regards|kind\\s+regards|with\\s+regards|regards|sincerely)/i.test(text);

                                    if (isClosing) {
                                        el.setAttribute('align', 'left');
                                        el.style.setProperty('text-align', 'left', 'important');
                                        el.style.setProperty('margin-top', '22px', 'important');
                                        el.style.setProperty('margin-bottom', '0', 'important');
                                    } else if (isSalutation) {
                                        el.setAttribute('align', 'left');
                                        el.style.setProperty('text-align', 'left', 'important');
                                        el.style.setProperty('margin-top', '0', 'important');
                                        el.style.setProperty('margin-bottom', '16px', 'important');
                                    } else {
                                        el.setAttribute('align', 'justify');
                                        el.style.setProperty('text-align', 'justify', 'important');
                                        el.style.setProperty('text-justify', 'inter-word', 'important');
                                        el.style.setProperty('text-align-last', 'left', 'important');
                                        el.style.setProperty('margin-top', '0', 'important');
                                        el.style.setProperty('margin-bottom', '16px', 'important');
                                    }
                                });

                                document.body.querySelectorAll('b, strong, .bold-field').forEach(el => {
                                    el.style.setProperty('font-weight', 'bold', 'important');
                                    el.style.setProperty('font-family', "'Times New Roman', Times, serif", 'important');
                                    el.style.setProperty('font-size', '12pt', 'important');
                                    el.style.setProperty('display', 'inline', 'important');
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
                        self.log("Formatted Times New Roman email successfully rendered into editor iframe!")
                        break
            except Exception:
                pass

        if not filled_body:
            body_selectors = [
                "div[contenteditable='true'][aria-label*='Message body' i]",
                "div[contenteditable='true'][aria-label*='body' i]",
                "div[contenteditable='true'][aria-label*='Mail' i]",
                "div[contenteditable='true']",
                "#mailContent",
                ".zmMailContent",
                "[data-test-id='mail-compose-body']",
                ".zmComposeBody"
            ]
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
                                    e.style.removeProperty('text-align');
                                    e.style.setProperty('font-family', "'Times New Roman', Times, serif", 'important');
                                    e.style.setProperty('font-size', '12pt', 'important');
                                    e.style.setProperty('color', '#000000', 'important');
                                    e.style.setProperty('line-height', '1.5', 'important');
                                    e.style.setProperty('margin', '0', 'important');
                                    e.style.setProperty('padding', '0', 'important');

                                    const sel = window.getSelection();
                                    const range = document.createRange();
                                    range.selectNodeContents(e);
                                    sel.removeAllRanges();
                                    sel.addRange(range);

                                    let ok = document.execCommand('insertHTML', false, html);
                                    if (!ok || !e.innerHTML.trim() || e.innerText.trim().length === 0) {
                                        e.innerHTML = html;
                                    }

                                    // Remove any unwanted leading empty elements/br before Dear Dr
                                    while (e.firstChild) {
                                        const fc = e.firstChild;
                                        const tag = (fc.nodeName || '').toLowerCase();
                                        const txt = (fc.textContent || '').replace(/[\\u200B-\\u200D\\uFEFF\\s]/g, '');
                                        if (tag === 'br' || (['p', 'div'].includes(tag) && txt === '')) {
                                            e.removeChild(fc);
                                        } else {
                                            break;
                                        }
                                    }

                                    const blocks = e.querySelectorAll('p, div');
                                    blocks.forEach(p => {
                                        if (p === e) return;
                                        if (p.querySelector('p, div')) return;
                                        const text = (p.innerText || p.textContent || '').trim();
                                        if (!text) return;

                                        p.style.setProperty('padding', '0', 'important');
                                        p.style.setProperty('line-height', '1.5', 'important');
                                        p.style.setProperty('font-family', "'Times New Roman', Times, serif", 'important');
                                        p.style.setProperty('font-size', '12pt', 'important');

                                        const firstLine = text.split('\\n')[0].trim();
                                        const isSalutation = /^(?:dear|hello|hi|greetings)/i.test(firstLine);
                                        const isClosing = /^(?:best\\s+regards|warm\\s+regards|kind\\s+regards|with\\s+regards|regards|sincerely)/i.test(firstLine) ||
                                                          /(?:best\\s+regards|warm\\s+regards|kind\\s+regards|with\\s+regards|regards|sincerely)/i.test(text);

                                        if (isClosing) {
                                            p.setAttribute('align', 'left');
                                            p.style.setProperty('text-align', 'left', 'important');
                                            p.style.setProperty('margin-top', '22px', 'important');
                                            p.style.setProperty('margin-bottom', '0', 'important');
                                        } else if (isSalutation) {
                                            p.setAttribute('align', 'left');
                                            p.style.setProperty('text-align', 'left', 'important');
                                            p.style.setProperty('margin-top', '0', 'important');
                                            p.style.setProperty('margin-bottom', '16px', 'important');
                                        } else {
                                            p.setAttribute('align', 'justify');
                                            p.style.setProperty('text-align', 'justify', 'important');
                                            p.style.setProperty('text-justify', 'inter-word', 'important');
                                            p.style.setProperty('text-align-last', 'left', 'important');
                                            p.style.setProperty('margin-top', '0', 'important');
                                            p.style.setProperty('margin-bottom', '16px', 'important');
                                        }
                                    });

                                    e.querySelectorAll('b, strong, .bold-field').forEach(child => {
                                        child.style.setProperty('font-weight', 'bold', 'important');
                                        child.style.setProperty('font-family', "'Times New Roman', Times, serif", 'important');
                                        child.style.setProperty('font-size', '12pt', 'important');
                                        child.style.setProperty('display', 'inline', 'important');
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
                            self.log("Formatted Times New Roman email successfully rendered into compose editor!")
                            break
                except Exception:
                    pass

        if not filled_body:
            # Fallback: Focus editor and use clipboard paste with HTML format
            self.log("Using clipboard paste fallback for formatted email...")
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

                # Post-paste styling enforcement across active elements
                page.evaluate(
                    """() => {
                        const containers = document.querySelectorAll("div[contenteditable='true'], body[contenteditable='true']");
                        containers.forEach(e => {
                            while (e.firstChild) {
                                const fc = e.firstChild;
                                const tag = (fc.nodeName || '').toLowerCase();
                                const txt = (fc.textContent || '').replace(/[\\u200B-\\u200D\\uFEFF\\s]/g, '');
                                if (tag === 'br' || (['p', 'div'].includes(tag) && txt === '')) {
                                    e.removeChild(fc);
                                } else {
                                    break;
                                }
                            }
                            e.querySelectorAll('p, div').forEach(p => {
                                if (p === e) return;
                                if (p.querySelector('p, div')) return;
                                const text = (p.innerText || p.textContent || '').trim();
                                if (!text) return;
                                const firstLine = text.split('\\n')[0].trim();
                                const isSalutation = /^(?:dear|hello|hi|greetings)/i.test(firstLine);
                                const isClosing = /^(?:best\\s+regards|warm\\s+regards|kind\\s+regards|with\\s+regards|regards|sincerely)/i.test(firstLine) ||
                                                  /(?:best\\s+regards|warm\\s+regards|kind\\s+regards|with\\s+regards|regards|sincerely)/i.test(text);
                                if (isClosing) {
                                    p.setAttribute('align', 'left');
                                    p.style.setProperty('text-align', 'left', 'important');
                                    p.style.setProperty('margin-top', '22px', 'important');
                                    p.style.setProperty('margin-bottom', '0', 'important');
                                } else if (isSalutation) {
                                    p.setAttribute('align', 'left');
                                    p.style.setProperty('text-align', 'left', 'important');
                                    p.style.setProperty('margin-top', '0', 'important');
                                    p.style.setProperty('margin-bottom', '16px', 'important');
                                } else {
                                    p.setAttribute('align', 'justify');
                                    p.style.setProperty('text-align', 'justify', 'important');
                                    p.style.setProperty('text-justify', 'inter-word', 'important');
                                    p.style.setProperty('text-align-last', 'left', 'important');
                                    p.style.setProperty('margin-top', '0', 'important');
                                    p.style.setProperty('margin-bottom', '16px', 'important');
                                }
                            });
                        });
                    }"""
                )
                filled_body = True
            except Exception:
                pass

        time.sleep(1.5)

        # 5. Click Send
        self.log("Clicking Send button...")
        clicked_send = False
        send_selectors = [
            "button[title*='Send' i]",
            "button:has-text('Send')",
            "span:text-is('Send')",
            "[data-test-id='send-button']",
            ".zmBtn.zmBtn_primary:has-text('Send')"
        ]
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
            self.log("Triggering shortcut 'Control+Enter' to dispatch...")
            page.keyboard.press("Control+Enter")

        time.sleep(3.5)

    def _on_close(self):
        self.cmd_queue.put(("QUIT", None))
        self.root.destroy()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="MimicMail — Multi-Instance Zoho Outreach Engine")
    parser.add_argument("--instance", type=str, default="Instance 1", help="Profile instance identifier (e.g. 'Instance 1', 'Instance 2')")
    args, unknown = parser.parse_known_args()

    root = tk.Tk()
    app = ModernDarkZohoAutomationApp(root, initial_instance=args.instance)
    root.mainloop()
