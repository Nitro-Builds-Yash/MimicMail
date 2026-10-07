import os
import sys
import time
import random
import queue
import threading
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from playwright.sync_api import sync_playwright
from PIL import Image, ImageTk

DEFAULT_SUBJECT = "{title}"
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pulsus_logo.png")

# Default user-customizable HTML templates for the 5 slots
SAMPLE_HTML_TEMPLATES = [
    """<p>Dear {name},</p>
<p>I hope this email finds you well.</p>
<p>I am reaching out regarding your esteemed work as {title}. We are connecting clinical leaders and researchers to collaborate on high-impact medical publications and global initiatives.</p>
<p>We would welcome the opportunity to discuss potential synergies with you.</p>
<p>Best regards,<br>MimicMail Editorial Board<br>editorial@mimicmail.com</p>""",

    """<p>Dear {name},</p>
<p>Greetings.</p>
<p>Given your notable leadership as {title}, our executive committee cordially invites you to participate as an honored speaker in our upcoming Global Medical Summit.</p>
<p>Your expertise would provide invaluable perspective to our attending clinicians and researchers.</p>
<p>Sincerely,<br>Scientific Program Committee<br>conferences@mimicmail.com</p>""",

    """<p>Dear {name},</p>
<p>I trust you are having a productive week.</p>
<p>In light of your distinguished track record as {title}, we are pleased to invite you to join our Specialized Advisory &amp; Peer Review Council.</p>
<p>We would appreciate the opportunity to share our upcoming clinical review guidelines with you.</p>
<p>Warm regards,<br>MimicMail Directorate<br>advisory@mimicmail.com</p>""",

    """<p>Dear {name},</p>
<p>I hope you are doing well.</p>
<p>We are reaching out to prominent professionals regarding your initiatives as {title}. We are coordinating multi-center clinical trials and translational research collaborations in your therapeutic domain.</p>
<p>Would you be open to reviewing a brief project abstract?</p>
<p>Best regards,<br>Clinical Research Division<br>trials@mimicmail.com</p>""",

    """<p>Dear {name},</p>
<p>I hope this message finds you having a wonderful day.</p>
<p>I wanted to connect directly with you regarding your ongoing clinical focus as {title}. Our organization partners with leading medical teams to accelerate knowledge exchange.</p>
<p>Thank you for your valuable time and consideration.</p>
<p>Best regards,<br>Outreach Operations Team<br>partners@mimicmail.com</p>"""
]

PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "zoho_browser_profile")

class ModernDarkZohoAutomationApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MimicMail — Zoho Outreach Automation")
        self.root.geometry("920x920")
        self.root.minsize(820, 800)

        # State flags
        self.is_running = False
        self.stop_requested = False
        self.excel_file_path = None
        self.total_rows_count = 0
        self.sent_rows_count = 0

        # Thread-safe work queue and background worker
        self.cmd_queue = queue.Queue()
        self.worker_thread = None

        os.makedirs(PROFILE_DIR, exist_ok=True)
        self._apply_dark_theme()
        self._build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _ensure_worker_started(self):
        if not self.worker_thread or not self.worker_thread.is_alive():
            self.worker_thread = threading.Thread(target=self._dedicated_browser_thread, daemon=True)
            self.worker_thread.start()

    def _apply_dark_theme(self):
        self.bg_root = "#0d1117"        # Dark canvas
        self.bg_card = "#161b22"        # Card surface
        self.bg_card_alt = "#1c2128"    # Highlighted banner
        self.border_card = "#30363d"    # Clean borders
        self.fg_primary = "#f0f6fc"     # Main white text
        self.fg_muted = "#8b949e"       # Subdued text
        self.accent_green = "#238636"   # Start accent
        self.accent_green_hover = "#2ea043"
        self.accent_blue = "#1f6feb"    # Login button
        self.accent_blue_hover = "#388bfd"
        self.accent_orange = "#f0883e"  # Active timer
        self.accent_red = "#da3633"     # Stop / Danger
        self.accent_red_hover = "#f85149"

        self.root.configure(bg=self.bg_root)

        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.style.configure(".", background=self.bg_root, foreground=self.fg_primary)

        # Notebook & Tabs
        self.style.configure("Dark.TNotebook", background=self.bg_card, borderwidth=0)
        self.style.configure("Dark.TNotebook.Tab", background="#21262d", foreground="#8b949e", font=("Segoe UI", 9, "bold"), padding=[10, 4])
        self.style.map("Dark.TNotebook.Tab",
                       background=[("selected", "#30363d"), ("active", "#282e33")],
                       foreground=[("selected", "#58a6ff"), ("active", "#c9d1d9")])

        # LabelFrames
        self.style.configure("DarkCard.TLabelframe", background=self.bg_card, bordercolor=self.border_card, relief="solid", borderwidth=1)
        self.style.configure("DarkCard.TLabelframe.Label", background=self.bg_card, foreground="#58a6ff", font=("Segoe UI", 10, "bold"))

        # Frames
        self.style.configure("Root.TFrame", background=self.bg_root)
        self.style.configure("Card.TFrame", background=self.bg_card)

        # Labels
        self.style.configure("Dark.TLabel", background=self.bg_card, foreground=self.fg_primary, font=("Segoe UI", 9))
        self.style.configure("DarkMuted.TLabel", background=self.bg_card, foreground=self.fg_muted, font=("Segoe UI", 9))

        # Buttons
        self.style.configure("StandardDark.TButton", font=("Segoe UI", 9, "bold"), padding=(12, 6),
                             background="#21262d", foreground="#c9d1d9", bordercolor=self.border_card, lightcolor="#21262d", darkcolor="#21262d")
        self.style.map("StandardDark.TButton",
                       background=[("active", "#30363d"), ("disabled", "#161b22")],
                       foreground=[("disabled", "#484f58")])

        self.style.configure("ActionBlue.TButton", font=("Segoe UI", 9, "bold"), padding=(14, 6),
                             background=self.accent_blue, foreground="#ffffff", bordercolor=self.border_card, lightcolor=self.accent_blue, darkcolor=self.accent_blue)
        self.style.map("ActionBlue.TButton",
                       background=[("active", self.accent_blue_hover), ("disabled", "#1b2d4b")],
                       foreground=[("disabled", "#52677e")])

        self.style.configure("ActionGreen.TButton", font=("Segoe UI", 12, "bold"), padding=(20, 10),
                             background="#238636", foreground="#ffffff", bordercolor="#3fb950", lightcolor="#2ea043", darkcolor="#238636")
        self.style.map("ActionGreen.TButton",
                       background=[("active", "#2ea043"), ("disabled", "#143a1f")],
                       foreground=[("disabled", "#4b7354")])

        self.style.configure("DangerRed.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 10),
                             background=self.accent_red, foreground="#ffffff", bordercolor=self.border_card, lightcolor=self.accent_red, darkcolor=self.accent_red)
        self.style.map("DangerRed.TButton",
                       background=[("active", self.accent_red_hover), ("disabled", "#3a1717")],
                       foreground=[("disabled", "#664141")])

        # Progressbar
        self.style.configure("Green.Horizontal.TProgressbar", troughcolor="#161b22", background="#2ea043", bordercolor="#30363d", lightcolor="#3fb950", darkcolor="#2ea043")

        # Sliders
        self.style.configure("Dark.Horizontal.TScale", background=self.bg_card, troughcolor="#21262d", bordercolor=self.border_card)

        # Checkbutton
        self.style.configure("Dark.TCheckbutton", background=self.bg_card, foreground=self.fg_primary, font=("Segoe UI", 9))
        self.style.map("Dark.TCheckbutton", background=[("active", self.bg_card)])

    def _build_ui(self):
        container = ttk.Frame(self.root, style="Root.TFrame", padding="14")
        container.pack(fill=tk.BOTH, expand=True)

        # --- Top Header / Hero Banner with PULSUS Logo ---
        hero_frame = tk.Frame(container, bg="#10151f", highlightbackground="#303e54", highlightthickness=1, padx=16, pady=10)
        hero_frame.pack(fill=tk.X, pady=(0, 10))

        header_top_row = tk.Frame(hero_frame, bg="#10151f")
        header_top_row.pack(fill=tk.X)

        # Load & display Pulsus logo
        self.logo_photo = None
        if os.path.exists(LOGO_PATH):
            try:
                pil_img = Image.open(LOGO_PATH)
                # Proportional resize to fit header neatly (approx height 38-42px)
                w_orig, h_orig = pil_img.size
                target_h = 40
                target_w = int(w_orig * (target_h / float(h_orig)))
                pil_resized = pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(pil_resized)
                logo_lbl = tk.Label(header_top_row, image=self.logo_photo, bg="#10151f")
                logo_lbl.pack(side=tk.LEFT, padx=(0, 14))
            except Exception as e:
                pass

        title_container = tk.Frame(header_top_row, bg="#10151f")
        title_container.pack(side=tk.LEFT, fill=tk.Y)

        title_lbl = tk.Label(title_container, text="MIMICMAIL", bg="#10151f", fg="#5bc0be", font=("Segoe UI", 15, "bold"))
        title_lbl.pack(anchor=tk.W)

        sub_lbl = tk.Label(title_container, text="Medical Leadership & Outreach Automation Platform | Powered by Pulsus Group",
                           bg="#10151f", fg="#8da4c4", font=("Segoe UI", 8))
        sub_lbl.pack(anchor=tk.W)

        badge_box = tk.Frame(header_top_row, bg="#10151f")
        badge_box.pack(side=tk.RIGHT)

        badge_lbl = tk.Label(badge_box, text="🔬 CLINICAL INTELLIGENCE EDITION", bg="#0b253a", fg="#64dfdf", font=("Segoe UI", 8, "bold"), padx=8, pady=4, relief="ridge", bd=1)
        badge_lbl.pack(anchor=tk.E)

        # --- Card 1: Browser Session ---
        auth_card = ttk.LabelFrame(container, text="  🌐 1. Clinical Web Portal Session  ", style="DarkCard.TLabelframe", padding="10")
        auth_card.pack(fill=tk.X, pady=(0, 8))

        auth_top_lbl = ttk.Label(
            auth_card,
            text="🔐 Authenticate securely with your Zoho medical email session.\n"
                 "Leave the browser open after sign-in; the clinical queue dispatches directly inside your active tab.",
            style="DarkMuted.TLabel"
        )
        auth_top_lbl.pack(anchor=tk.W, pady=(0, 6))

        auth_action_row = ttk.Frame(auth_card, style="Card.TFrame")
        auth_action_row.pack(fill=tk.X)

        self.login_btn = ttk.Button(auth_action_row, text="🌐 Launch Webmail & Sign In", command=self.request_open_browser, style="ActionBlue.TButton")
        self.login_btn.pack(side=tk.LEFT)

        self.session_indicator = tk.Label(auth_action_row, text="⚕️ Status: Ready for connection", bg=self.bg_card, fg=self.fg_muted, font=("Segoe UI", 9, "bold"), padx=12)
        self.session_indicator.pack(side=tk.LEFT)

        # --- Card 2: Excel Spreadsheet Ingestion ---
        file_card = ttk.LabelFrame(container, text="  📋 2. Recipient Medical Registry (.xlsx)  ", style="DarkCard.TLabelframe", padding="10")
        file_card.pack(fill=tk.X, pady=(0, 8))

        self.file_label = ttk.Label(file_card, text="📄 No registry loaded (Required columns: name, email, title)", style="DarkMuted.TLabel")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.browse_btn = ttk.Button(file_card, text="📁 Browse Registry (.xlsx)", command=self.browse_excel, style="StandardDark.TButton")
        self.browse_btn.pack(side=tk.RIGHT)

        # --- Card 3: Dynamic HTML Templates (5 Slots) ---
        template_card = ttk.LabelFrame(container, text="  ✉️ 3. Drop / Paste Custom HTML Templates (5 Slots - Random Selection)  ", style="DarkCard.TLabelframe", padding="10")
        template_card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        subj_header_frame = ttk.Frame(template_card, style="Card.TFrame")
        subj_header_frame.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(subj_header_frame, text="📌 Subject Line Template:", style="Dark.TLabel").pack(side=tk.LEFT)
        ttk.Label(subj_header_frame, text="(Defaults to {title} from each row)", style="DarkMuted.TLabel").pack(side=tk.LEFT, padx=6)

        self.subj_entry = tk.Entry(template_card, bg="#0d1117", fg="#f0f6fc", insertbackground="#f0f6fc",
                                   highlightbackground=self.border_card, highlightcolor="#58a6ff", highlightthickness=1,
                                   relief="flat", font=("Segoe UI", 9))
        self.subj_entry.insert(0, DEFAULT_SUBJECT)
        self.subj_entry.pack(fill=tk.X, pady=(0, 6), ipady=3)

        # Tabbed Notebook for 5 HTML Templates
        notebook_header = ttk.Frame(template_card, style="Card.TFrame")
        notebook_header.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(notebook_header, text="📝 Drop/Paste up to 5 HTML Templates below. For each recipient, a template is picked at random:", style="Dark.TLabel").pack(side=tk.LEFT)
        tk.Label(notebook_header, text="Fonts: Times New Roman | Bold: {name} & {title}", bg="#0b253a", fg="#64dfdf", font=("Segoe UI", 8, "bold"), padx=6, pady=2).pack(side=tk.RIGHT)

        self.template_notebook = ttk.Notebook(template_card, style="Dark.TNotebook")
        self.template_notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 2))

        self.template_text_widgets = []
        for i in range(5):
            tab_frame = ttk.Frame(self.template_notebook, style="Card.TFrame")
            self.template_notebook.add(tab_frame, text=f" 📄 Template {i+1} ")

            toolbar = ttk.Frame(tab_frame, style="Card.TFrame")
            toolbar.pack(fill=tk.X, pady=(2, 2))
            ttk.Label(toolbar, text=f"Slot {i+1} (HTML format):", style="DarkMuted.TLabel").pack(side=tk.LEFT)

            load_btn = ttk.Button(toolbar, text=f"📂 Load HTML File into Slot {i+1}",
                                  command=lambda idx=i: self._load_template_file(idx),
                                  style="StandardDark.TButton")
            load_btn.pack(side=tk.RIGHT)

            clear_btn = ttk.Button(toolbar, text="🗑 Clear",
                                   command=lambda idx=i: self._clear_template_slot(idx),
                                   style="StandardDark.TButton")
            clear_btn.pack(side=tk.RIGHT, padx=4)

            txt = tk.Text(tab_frame, height=4, wrap=tk.WORD, bg="#0d1117", fg="#f0f6fc",
                          insertbackground="#f0f6fc", highlightbackground=self.border_card,
                          highlightcolor="#58a6ff", highlightthickness=1, relief="flat", font=("Times New Roman", 11))
            initial_content = SAMPLE_HTML_TEMPLATES[i] if i < len(SAMPLE_HTML_TEMPLATES) else ""
            txt.insert(tk.END, initial_content)
            txt.pack(fill=tk.BOTH, expand=True, pady=(0, 2))
            self.template_text_widgets.append(txt)

        # --- Card 4: Anti-Bot Pacing & Delivery Safeguards ---
        pacing_card = ttk.LabelFrame(container, text="  ⏱ 4. Anti-Bot Pacing & Delivery Safeguards  ", style="DarkCard.TLabelframe", padding="10")
        pacing_card.pack(fill=tk.X, pady=(0, 8))

        slider_row = ttk.Frame(pacing_card, style="Card.TFrame")
        slider_row.pack(fill=tk.X, pady=(0, 2))

        ttk.Label(slider_row, text="⏳ Rest Interval:", style="Dark.TLabel").pack(side=tk.LEFT, padx=(0, 8))
        self.interval_var = tk.DoubleVar(value=2.0)
        self.slider = ttk.Scale(slider_row, from_=1.0, to=5.0, variable=self.interval_var, command=self._slider_changed, style="Dark.Horizontal.TScale")
        self.slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)

        self.slider_val_lbl = ttk.Label(slider_row, text="2.0 mins", style="Dark.TLabel", width=8)
        self.slider_val_lbl.pack(side=tk.LEFT)

        self.jitter_var = tk.BooleanVar(value=True)
        self.jitter_check = ttk.Checkbutton(pacing_card, text="🩺 Human Behavioral Jitter (±15 to 35s randomized cadence)",
                                            variable=self.jitter_var, style="Dark.TCheckbutton")
        self.jitter_check.pack(anchor=tk.W, pady=(2, 4))

        timer_row = ttk.Frame(pacing_card, style="Card.TFrame")
        timer_row.pack(fill=tk.X)

        self.timer_display = tk.Label(timer_row, text="⏱ Status: Queue Ready", bg=self.bg_card, fg="#58a6ff", font=("Segoe UI", 9, "bold"))
        self.timer_display.pack(side=tk.LEFT)

        self.stats_lbl = ttk.Label(timer_row, text="📊 Progress: 0 / 0 | Sent: 0", style="DarkMuted.TLabel")
        self.stats_lbl.pack(side=tk.RIGHT)

        # --- Card 5: Live Status Graph & Loading Bar ---
        status_card = ttk.LabelFrame(container, text="  📈 5. Live Outreach Status & Interactive Progress Graph  ", style="DarkCard.TLabelframe", padding="10")
        status_card.pack(fill=tk.X, pady=(0, 8))

        # Progress bar + status text
        pbar_row = ttk.Frame(status_card, style="Card.TFrame")
        pbar_row.pack(fill=tk.X, pady=(0, 4))

        self.progress_bar = ttk.Progressbar(pbar_row, style="Green.Horizontal.TProgressbar", mode="determinate", maximum=100, value=0)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.progress_pct_lbl = tk.Label(pbar_row, text="0%", bg=self.bg_card, fg="#2ea043", font=("Segoe UI", 10, "bold"), width=5)
        self.progress_pct_lbl.pack(side=tk.RIGHT)

        # Canvas for Status Bar Graph
        graph_box = ttk.Frame(status_card, style="Card.TFrame")
        graph_box.pack(fill=tk.X, pady=(4, 2))

        self.graph_canvas = tk.Canvas(graph_box, bg="#0d1117", height=42, highlightthickness=1, highlightbackground=self.border_card)
        self.graph_canvas.pack(fill=tk.X, expand=True)
        self.graph_canvas.bind("<Configure>", lambda e: self._draw_status_graph())

        self.graph_legend_lbl = tk.Label(status_card, text="🟩 Sent: 0 (0%)  |  🟦 Pending: 0 (0%)  |  🟥 Skipped / Error: 0 (0%)",
                                         bg=self.bg_card, fg=self.fg_muted, font=("Segoe UI", 8))
        self.graph_legend_lbl.pack(anchor=tk.W, pady=(2, 0))

        # --- Section 6: Start / Stop Controls & Live Terminal ---
        ctrl_frame = ttk.Frame(container, style="Root.TFrame")
        ctrl_frame.pack(fill=tk.X, pady=(0, 6))

        self.start_btn = ttk.Button(ctrl_frame, text="🚀 START AUTOMATED OUTREACH", command=self.request_start_automation, style="ActionGreen.TButton")
        self.start_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 6), ipady=3)

        self.stop_btn = ttk.Button(ctrl_frame, text="🛑 EMERGENCY STOP", command=self.stop_automation, state=tk.DISABLED, style="DangerRed.TButton")
        self.stop_btn.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(6, 0), ipady=3)

        # Terminal Log
        log_frame = ttk.LabelFrame(container, text="  🩺 Medical Dispatch Live Audit Terminal  ", style="DarkCard.TLabelframe", padding="8")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=4, state=tk.DISABLED, wrap=tk.WORD, bg="#0d1117", fg="#7ee787",
                                insertbackground="#f0f6fc", relief="flat", font=("Consolas", 8),
                                highlightbackground=self.border_card, highlightthickness=1)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _draw_status_graph(self, sent: int = 0, pending: int = 0, errors: int = 0):
        self.graph_canvas.delete("all")
        w = self.graph_canvas.winfo_width()
        h = self.graph_canvas.winfo_height()
        if w <= 1:
            w = 860
        if h <= 1:
            h = 42

        total = sent + pending + errors
        if total == 0:
            # Draw placeholder background
            self.graph_canvas.create_rectangle(0, 0, w, h, fill="#161b22", outline="")
            self.graph_canvas.create_text(w // 2, h // 2, text="Load spreadsheet to initialize live status graph", fill="#484f58", font=("Segoe UI", 9, "italic"))
            return

        sent_pct = (sent / total)
        pending_pct = (pending / total)
        err_pct = (errors / total)

        sent_w = int(w * sent_pct)
        pending_w = int(w * pending_pct)
        err_w = w - sent_w - pending_w

        # Draw segmented colored bar
        x_curr = 0
        if sent_w > 0:
            self.graph_canvas.create_rectangle(x_curr, 0, x_curr + sent_w, h, fill="#238636", outline="")
            if sent_w > 50:
                self.graph_canvas.create_text(x_curr + sent_w // 2, h // 2, text=f"Sent: {sent} ({int(sent_pct*100)}%)", fill="#ffffff", font=("Segoe UI", 8, "bold"))
            x_curr += sent_w

        if pending_w > 0:
            self.graph_canvas.create_rectangle(x_curr, 0, x_curr + pending_w, h, fill="#1f6feb", outline="")
            if pending_w > 50:
                self.graph_canvas.create_text(x_curr + pending_w // 2, h // 2, text=f"Pending: {pending} ({int(pending_pct*100)}%)", fill="#ffffff", font=("Segoe UI", 8, "bold"))
            x_curr += pending_w

        if err_w > 0:
            self.graph_canvas.create_rectangle(x_curr, 0, w, h, fill="#da3633", outline="")
            if err_w > 50:
                self.graph_canvas.create_text(x_curr + err_w // 2, h // 2, text=f"Error/Skip: {errors}", fill="#ffffff", font=("Segoe UI", 8, "bold"))

        legend_text = f"🟩 Sent: {sent} ({sent_pct*100:.1f}%)   |   🟦 Pending: {pending} ({pending_pct*100:.1f}%)   |   🟥 Skipped / Error: {errors} ({err_pct*100:.1f}%)   |   Total: {total}"
        self.graph_legend_lbl.config(text=legend_text)

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
            req = {"name", "email", "title"}
            col_map = {c.strip().lower(): c for c in df.columns}
            missing = req - set(col_map.keys())

            if missing:
                messagebox.showerror("Invalid Headers", f"The Excel sheet must contain 'name', 'email', and 'title'.\n\nMissing: {', '.join(missing)}")
                return

            self.excel_file_path = path
            self.file_label.config(text=f"Loaded: {os.path.basename(path)}", foreground=self.fg_primary)

            status_col = next((c for c in df.columns if c.strip().lower() == "status"), None)
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
            self.log(f"Workbook loaded: {total} total rows ({sent} already SENT, {pending} ready).")

        except Exception as e:
            messagebox.showerror("File Error", f"Unable to read Excel file:\n{str(e)}")

    def request_start_automation(self):
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            messagebox.showwarning("File Missing", "Please select an Excel workbook first.")
            return

        subj_tmpl = self.subj_entry.get().strip()
        if not subj_tmpl:
            messagebox.showwarning("Missing Subject", "Please enter a subject line template (e.g. {title}).")
            return

        # Gather all non-empty user templates across the 5 slots
        active_templates = []
        for i, txt_widget in enumerate(self.template_text_widgets):
            content = txt_widget.get("1.0", tk.END).strip()
            if content:
                active_templates.append(content)

        if not active_templates:
            messagebox.showwarning("No Templates", "Please provide at least 1 template in the 5 slots (paste HTML or load a file).")
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
            "excel_path": self.excel_file_path
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

                elif cmd == "OPEN_BROWSER":
                    try:
                        if not (browser_context and browser_page and not browser_page.is_closed()):
                            self.log("Opening persistent Chromium browser...")
                            self.root.after(0, lambda: self.session_indicator.config(text="● Launching...", fg=self.accent_orange))

                            browser_context = playwright_instance.chromium.launch_persistent_context(
                                user_data_dir=PROFILE_DIR,
                                headless=False,
                                args=["--start-maximized"]
                            )
                            browser_page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()
                            browser_page.goto("https://mail.zoho.com")

                            self.log("Browser is open! Log in and leave this window open.")
                            self.root.after(0, lambda: self.session_indicator.config(text="● Browser: Active & Connected", fg="#3fb950"))
                        else:
                            self.log("Browser already open. Bringing tab to front...")
                            browser_page.bring_to_front()
                    except Exception as err:
                        self.log(f"Open browser error: {err}")
                        self.root.after(0, lambda: self.session_indicator.config(text="● Open Error", fg=self.accent_red))

                elif cmd == "START_AUTOMATION":
                    try:
                        # 1. Ensure browser is open in this same thread
                        if not (browser_context and browser_page and not browser_page.is_closed()):
                            self.log("Opening browser session...")
                            browser_context = playwright_instance.chromium.launch_persistent_context(
                                user_data_dir=PROFILE_DIR,
                                headless=False,
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

                        self.root.after(0, lambda: self.session_indicator.config(text="● Browser: In Mailbox & Automating", fg="#3fb950"))

                        # 2. Run batch dispatch
                        self._execute_batch(browser_page, data)

                    except Exception as batch_err:
                        self.log(f"Automation execution error: {batch_err}")
                        self.root.after(0, lambda e=str(batch_err): messagebox.showerror("Automation Error", e))
                    finally:
                        self.is_running = False
                        self.root.after(0, lambda: self.timer_display.config(text="⏱ Timer: Idle", fg="#58a6ff"))
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

    def _execute_batch(self, page, params):
        excel_path = params["excel_path"]
        subj_tmpl = params["subj_tmpl"]
        body_tmpls = params.get("body_tmpls", [params.get("body_tmpl", "")])
        base_mins = params["base_mins"]
        use_jitter = params["use_jitter"]

        df = pd.read_excel(excel_path)
        col_map = {c.strip().lower(): c for c in df.columns}

        name_col = col_map["name"]
        email_col = col_map["email"]
        title_col = col_map["title"]

        status_col = next((c for c in df.columns if c.strip().lower() == "status"), "Status")
        if status_col not in df.columns:
            df[status_col] = ""
        df[status_col] = df[status_col].astype(object)

        timestamp_col = next((c for c in df.columns if c.strip().lower() == "sent_at"), "Sent_At")
        if timestamp_col not in df.columns:
            df[timestamp_col] = ""
        df[timestamp_col] = df[timestamp_col].astype(object)

        pending_indices = [
            i for i, row in df.iterrows()
            if str(row[status_col]).strip().upper() != "SENT"
        ]

        if not pending_indices:
            self.log("All rows in this spreadsheet are already marked as SENT.")
            self.root.after(0, lambda: messagebox.showinfo("Completed", "All recipients in this spreadsheet are already sent!"))
            return

        BATCH_CAP = 40
        target_indices = pending_indices[:BATCH_CAP]
        total_in_batch = len(target_indices)

        self.log(f"Dispatching batch of {total_in_batch} recipients across {len(body_tmpls)} random template pool...")

        sent_count = 0
        for q_idx, r_idx in enumerate(target_indices):
            if self.stop_requested:
                self.log("Automation interrupted by user.")
                break

            row = df.loc[r_idx]
            r_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else "Valued Professional"
            r_email = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
            r_title = str(row[title_col]).strip() if pd.notna(row[title_col]) else "Professional"

            import re
            EMAIL_REGEX = re.compile(r"^[\w\.-]+@[\w\.-]+\.\w+$")
            if not r_email or not EMAIL_REGEX.match(r_email):
                self.log(f"Row {r_idx + 1}: SKIPPED (Invalid email address '{r_email}')")
                df.loc[r_idx, status_col] = "INVALID_EMAIL"
                df.to_excel(excel_path, index=False)
                continue

            # Randomly select 1 template from the user-provided templates
            chosen_body_tmpl = random.choice(body_tmpls)

            # Robust placeholder substitution (supports {name}, {title}, {email} in any case/whitespace)
            personalized_subject = subj_tmpl
            for ph, val in [("{title}", r_title), ("{name}", r_name), ("{TITLE}", r_title), ("{NAME}", r_name)]:
                personalized_subject = personalized_subject.replace(ph, val)

            # Bold name and title in the rendered email body as requested
            r_name_bold = f"<b>{r_name}</b>"
            r_title_bold = f"<b>{r_title}</b>"

            personalized_body = chosen_body_tmpl
            for ph, val in [("{name}", r_name_bold), ("{title}", r_title_bold), ("{email}", r_email),
                            ("{NAME}", r_name_bold), ("{TITLE}", r_title_bold), ("{EMAIL}", r_email)]:
                personalized_body = personalized_body.replace(ph, val)

            # Strip any leading spaces or lines so 'Dear {name}' always starts at character 0
            personalized_body = personalized_body.strip()

            self.log(f"Prepared subject: '{personalized_subject}'")
            self.log(f"Dispatching ({q_idx + 1}/{total_in_batch}) to {r_name} <{r_email}> (Bold Name & Title)...")

            try:
                self._send_in_browser(page, r_email, personalized_subject, personalized_body)
                sent_count += 1
                df.at[r_idx, status_col] = str("SENT")
                df.at[r_idx, timestamp_col] = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                df.to_excel(excel_path, index=False)
                self.log(f"SUCCESS: Email dispatched to {r_email} & saved to Excel.")
            except Exception as err:
                # Catch any unexpected row errors, log, record to Excel, close compose if stuck, and continue with next recipient!
                self.log(f"SKIPPED ROW {r_idx + 1} ({r_email}) due to error: {err}")
                df.at[r_idx, status_col] = str(f"ERROR: {str(err)[:50]}")
                df.to_excel(excel_path, index=False)
                try:
                    # Press Escape or discard draft to keep browser clean for next recipient
                    page.keyboard.press("Escape")
                    time.sleep(1)
                except Exception:
                    pass

            # Recalculate stats for live status graph and progress bar
            curr_status_series = df[status_col].astype(str).str.strip().str.upper()
            curr_sent = (curr_status_series == "SENT").sum()
            curr_errors = curr_status_series.str.startswith("ERROR").sum() + (curr_status_series == "INVALID_EMAIL").sum()
            curr_pending = len(df) - curr_sent - curr_errors
            curr_pct = int((curr_sent / len(df) * 100)) if len(df) > 0 else 0

            self.root.after(0, lambda s=curr_sent, t=len(df), p=curr_pending, e=curr_errors, pct=curr_pct: (
                self.stats_lbl.config(text=f"Total: {t} | Sent: {s} | Pending: {p}"),
                self.progress_bar.config(value=pct),
                self.progress_pct_lbl.config(text=f"{pct}%"),
                self._draw_status_graph(sent=s, pending=p, errors=e)
            ))

            # Anti-bot human delay
            if q_idx < total_in_batch - 1 and not self.stop_requested:
                base_sec = base_mins * 60.0
                jitter_sec = random.randint(-25, 35) if use_jitter else 0
                sleep_sec = max(30.0, base_sec + jitter_sec)

                self.log(f"Pacing: Resting for {sleep_sec:.0f} seconds before next recipient...")

                remaining = int(sleep_sec)
                while remaining > 0 and not self.stop_requested:
                    m, s = divmod(remaining, 60)
                    self.root.after(0, lambda mm=m, ss=s: self.timer_display.config(
                        text=f"⏱ Next email in: {mm:02d}:{ss:02d}", fg=self.accent_orange
                    ))
                    time.sleep(1)
                    remaining -= 1

        self.log(f"Batch completed! Total sent: {sent_count}")
        self.root.after(0, lambda s=sent_count: messagebox.showinfo("Done", f"Automation finished!\nTotal sent: {s}"))

    def _send_in_browser(self, page, to_email, subject, body):
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

        # 4. Locate and fill Message Body in Times New Roman font
        self.log(f"Inserting body text ({len(body)} characters) in Times New Roman...")
        filled_body = False

        # Convert body (supports raw HTML drops or text) to styled HTML paragraphs in Times New Roman
        if "<p" in body.lower() or "<div" in body.lower() or "<br" in body.lower():
            # Already HTML formatted, wrap in Times New Roman font container
            styled_html = f"<div style=\"font-family: 'Times New Roman', Times, serif; font-size: 12pt; color: #000000; line-height: 1.5;\">{body}</div>"
        else:
            # Plain text converted to styled paragraphs
            html_paragraphs = "".join([f"<p style=\"margin: 0 0 10px 0; font-family: 'Times New Roman', Times, serif; font-size: 12pt;\">{p.strip()}</p>" if p.strip() else "<p><br></p>" for p in body.split("\n\n")])
            styled_html = f"<div style=\"font-family: 'Times New Roman', Times, serif; font-size: 12pt; color: #000000; line-height: 1.5;\">{html_paragraphs}</div>"

        # First check inside frames/iframes (TinyMCE editor in Zoho)
        for frame in page.frames:
            try:
                body_elem = frame.query_selector("body[contenteditable='true'], body")
                if body_elem and body_elem.is_visible():
                    body_elem.click()
                    time.sleep(0.2)
                    # Clear out any pre-existing whitespace / blank lines by pressing Backspace and Home
                    frame.keyboard.press("Control+Home")
                    frame.keyboard.press("Backspace")
                    time.sleep(0.1)

                    # Directly insert formatted Times New Roman HTML so no leading spaces occur
                    frame.evaluate(f"(html) => {{ document.body.innerHTML = html; }}", styled_html)
                    filled_body = True
                    self.log("Body successfully rendered into editor iframe in Times New Roman!")
                    break
            except Exception:
                pass

        if not filled_body:
            body_selectors = [
                "div[contenteditable='true'][aria-label*='Message body' i]",
                "div[contenteditable='true'][aria-label*='body' i]",
                "div[contenteditable='true']",
                "#mailContent",
                ".zmMailContent",
                "[data-test-id='mail-compose-body']"
            ]
            for sel in body_selectors:
                try:
                    el = page.query_selector(sel)
                    if el and el.is_visible():
                        el.click()
                        time.sleep(0.2)
                        page.keyboard.press("Control+Home")
                        page.keyboard.press("Backspace")
                        time.sleep(0.1)
                        page.evaluate(f"(selector, html) => {{ const e = document.querySelector(selector); if (e) e.innerHTML = html; }}", sel, styled_html)
                        filled_body = True
                        self.log("Body successfully rendered into contenteditable div in Times New Roman!")
                        break
                except Exception:
                    pass

        if not filled_body:
            # Fallback typing
            self.log("Focusing editor via Tab key navigation...")
            page.keyboard.press("Tab")
            time.sleep(0.3)
            page.keyboard.press("Control+Home")
            page.keyboard.press("Backspace")
            page.keyboard.type(body.strip(), delay=5)

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
    root = tk.Tk()
    app = ModernDarkZohoAutomationApp(root)
    root.mainloop()
