"""Visible, browser-assisted mail merge with a human-controlled send step."""

from __future__ import annotations

import tkinter as tk
import os
import struct
import sys
import time
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

# Some embedded Windows Python launches omit this standard architecture
# variable, which tkinterdnd2 uses to select its bundled native extension.
if sys.platform == "win32" and not os.environ.get("PROCESSOR_ARCHITECTURE"):
    os.environ["PROCESSOR_ARCHITECTURE"] = "AMD64" if struct.calcsize("P") == 8 else "x86"

from zoho_mail_automation.data.workbook import Recipient, RecipientWorkbook


class MailMergeApp:
    """Prepare one spreadsheet message at a time for manual sending in Zoho Mail."""

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Zoho Mail Merge — Review and Send")
        root.geometry("900x850")
        root.minsize(760, 720)
        self._style()
        self.path = tk.StringVar()
        self.login_id = tk.StringVar()
        self.interval_minutes = tk.IntVar(value=3)
        self.subject = tk.StringVar()
        self.default_body = ("Hi {name},\n\nI came across your profile and wanted to reach out "
                             "regarding your role as {title}.\n\nWe would love to connect and share some "
                             "insights with you.\n\nBest regards,\n[Your Name]")
        self.status = tk.StringVar(value="Drop a workbook here or choose a file.")
        self.progress_text = tk.StringVar(value="0 of 0 messages confirmed sent")
        self.columns = {
            "name": "name", "email": "email", "title": "title",
            "status": "Status", "sent_at": "Sent_At",
        }
        self.workbook: RecipientWorkbook | None = None
        self.recipients: list[Recipient] = []
        self.skipped: list[tuple[int, str]] = []
        self.index = 0
        self.current: Recipient | None = None
        self.ready_at = 0.0
        self.pace_job: str | None = None
        self._build()
        root.protocol("WM_DELETE_WINDOW", self._close)

    def _style(self) -> None:
        self.root.configure(bg="#edf2f8")
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background="#edf2f8")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Hero.TFrame", background="#152d4c")
        style.configure("HeroTitle.TLabel", background="#152d4c", foreground="#ffffff",
                        font=("Segoe UI", 17, "bold"))
        style.configure("HeroSub.TLabel", background="#152d4c", foreground="#d8e5f3",
                        font=("Segoe UI", 10))
        style.configure("Card.TLabelframe", background="#ffffff", bordercolor="#d7e0eb")
        style.configure("Card.TLabelframe.Label", background="#ffffff", foreground="#203653",
                        font=("Segoe UI", 11, "bold"))
        style.configure("TLabel", foreground="#27384d", background="#edf2f8", font=("Segoe UI", 10))
        style.configure("Card.TLabel", background="#ffffff", foreground="#27384d", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=(11, 7))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 8),
                        foreground="#ffffff", background="#167c80")
        style.map("Accent.TButton", background=[("active", "#11686c"), ("disabled", "#9bb9b9")])
        style.configure("Horizontal.TProgressbar", troughcolor="#dde6ef", background="#16878b",
                        bordercolor="#dde6ef", lightcolor="#16878b", darkcolor="#16878b", thickness=14)

    def _build(self) -> None:
        frame = ttk.Frame(self.root, padding=20, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(5, weight=1)

        hero = ttk.Frame(frame, padding=(20, 16), style="Hero.TFrame")
        hero.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        ttk.Label(hero, text="Zoho Mail Merge", style="HeroTitle.TLabel").pack(anchor="w")
        ttk.Label(hero, text="Prepare, review, and send one message at a time",
                  style="HeroSub.TLabel").pack(anchor="w", pady=(4, 0))
        ttk.Label(frame, text=("Sign in and send in Zoho Mail. This app does not collect your password "
                              "or send messages automatically."), wraplength=820).grid(
            row=1, column=0, sticky="w", pady=(0, 12))

        drop = ttk.Label(frame, text="Drop an Excel (.xlsx) or CSV file here\n—or click to browse",
                         anchor="center", relief="groove", padding=18, background="#ffffff",
                         foreground="#167c80", font=("Segoe UI", 11, "bold"))
        drop.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        drop.bind("<Button-1>", lambda _event: self._browse())
        if DND_FILES is not None:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self._drop_file)
        ttk.Label(frame, textvariable=self.path).grid(row=3, column=0, sticky="w", pady=(0, 10))

        progress_card = ttk.Frame(frame, padding=(2, 2), style="App.TFrame")
        progress_card.grid(row=4, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(progress_card, textvariable=self.progress_text).pack(anchor="w", pady=(0, 5))
        self.progress = ttk.Progressbar(progress_card, mode="determinate", maximum=1,
                                         style="Horizontal.TProgressbar")
        self.progress.pack(fill="x")

        review = ttk.LabelFrame(frame, text="Message review", padding=14, style="Card.TLabelframe")
        review.grid(row=5, column=0, sticky="nsew")
        review.columnconfigure(1, weight=1)
        review.rowconfigure(5, weight=1)
        ttk.Label(review, text="Zoho login ID", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(review, textvariable=self.login_id, width=38).grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(review, text="Password entry stays on Zoho's sign-in page.", style="Card.TLabel").grid(
            row=0, column=2, sticky="w", padx=8, pady=5)
        ttk.Label(review, text="Delay between messages").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        delay = ttk.Combobox(review, textvariable=self.interval_minutes, values=(1, 2, 3, 4, 5),
                             state="readonly", width=8)
        delay.grid(row=1, column=1, sticky="w", pady=5)
        ttk.Label(review, text="minutes").grid(row=1, column=2, sticky="w", padx=8, pady=5)

        ttk.Label(review, text="Subject (from workbook title)", style="Card.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(review, textvariable=self.subject, state="readonly").grid(
            row=2, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Label(review, text="Body template", style="Card.TLabel").grid(row=3, column=0, sticky="nw", padx=(0, 10), pady=5)
        self.body_template = tk.Text(review, height=7, wrap="word")
        self.body_template.grid(row=3, column=1, columnspan=2, sticky="ew", pady=5)
        self.body_template.insert("1.0", self.default_body)
        ttk.Button(review, text="Refresh preview", command=self._render_preview).grid(
            row=4, column=1, sticky="w", pady=(4, 8))
        self.counter = ttk.Label(review, text="No message loaded", font=("Segoe UI", 10, "bold"))
        self.counter.grid(row=4, column=2, sticky="e", pady=(4, 8))
        self.preview = tk.Text(review, height=12, wrap="word", state="disabled")
        self.preview.grid(row=5, column=0, columnspan=3, sticky="nsew")
        self.preview.configure(bg="#f7f9fc", fg="#243449", insertbackground="#243449",
                               relief="flat", padx=10, pady=10, font=("Consolas", 10))
        buttons = ttk.Frame(review)
        buttons.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 2))
        self.open_button = ttk.Button(buttons, text="Open Zoho Mail", command=self._open_zoho)
        self.open_button.pack(side="left")
        self.auto_btn = ttk.Button(buttons, text="⚡ Run Full Automation", command=self._start_automation, style="Accent.TButton")
        self.auto_btn.pack(side="left", padx=6)
        self.stop_btn = ttk.Button(buttons, text="Stop Automation", command=self._stop_automation, state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        self.copy_subject_button = ttk.Button(buttons, text="Copy subject", command=self._copy_subject)
        self.copy_subject_button.pack(side="left", padx=6)
        self.copy_button = ttk.Button(buttons, text="Copy message", command=self._copy_message)
        self.copy_button.pack(side="left")
        ttk.Button(buttons, text="Mark sent & next", command=self._mark_sent).pack(side="right")
        ttk.Button(buttons, text="Load workbook", command=self._browse).pack(side="right", padx=8)

        ttk.Label(frame, textvariable=self.status).grid(row=6, column=0, sticky="w", pady=(10, 0))

    def _browse(self) -> None:
        selected = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx"), ("CSV file", "*.csv")])
        if selected:
            self._load(Path(selected))

    def _drop_file(self, event) -> None:
        try:
            paths = self.root.tk.splitlist(event.data)
            if paths:
                self._load(Path(paths[0]))
        except Exception as exc:
            messagebox.showerror("File drop failed", str(exc))

    def _load(self, path: Path) -> None:
        try:
            if self.workbook:
                self.workbook.close()
            self.workbook = RecipientWorkbook(path, self.columns)
            self.recipients, self.skipped = self.workbook.valid_recipients()
            self.index = 0
            self.path.set(str(path))
            self.progress.configure(maximum=max(1, len(self.recipients)), value=0)
            self.progress_text.set(f"0 of {len(self.recipients)} messages confirmed sent")
            self.status.set(f"Loaded {len(self.recipients)} valid unsent rows; skipped {len(self.skipped)} rows.")
            self._show_current()
            self._open_zoho()
        except Exception as exc:
            self.workbook = None
            messagebox.showerror("Could not load workbook", str(exc))

    def _show_current(self) -> None:
        if self.index >= len(self.recipients):
            self.current = None
            self.subject.set("")
            self.counter.configure(text="No pending messages")
            self.status.set("All valid unsent rows have been marked sent.")
            self._set_preview("Queue complete.")
            return
        self.current = self.recipients[self.index]
        item = self.current
        self.subject.set(item.title)
        self.counter.configure(text=f"Message {self.index + 1} of {len(self.recipients)}  •  workbook row {item.row_number}")
        self._render_preview()

    def _render_preview(self) -> None:
        if not self.current:
            return
        item = self.current
        try:
            values = {"name": item.name, "email": item.email, "title": item.title}
            body = self.body_template.get("1.0", "end-1c").format_map(values)
        except (KeyError, ValueError, IndexError) as exc:
            messagebox.showerror("Template error", f"Check template placeholders and braces: {exc}")
            return
        subject = self.subject.get().strip()
        self._set_preview(f"To: {item.email}\nSubject: {subject}\n\n{body}")

    def _set_preview(self, text: str) -> None:
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)
        self.preview.configure(state="disabled")

    def _open_zoho(self) -> None:
        if time.monotonic() < self.ready_at:
            return
        if self.login_id.get().strip():
            self.root.clipboard_clear()
            self.root.clipboard_append(self.login_id.get().strip())
            self.status.set("Zoho Mail opened; login ID copied. Log in on the Zoho page.")
        else:
            self.status.set("Zoho Mail opened. Log in on the Zoho page.")
        webbrowser.open("https://mail.zoho.com/")

    def _copy_message(self) -> None:
        if not self.current or time.monotonic() < self.ready_at:
            return
        self._render_preview()
        self.root.clipboard_clear()
        self.root.clipboard_append(self.preview.get("1.0", "end-1c"))
        self.status.set("Recipient, subject, and body copied. Review and send in Zoho Mail.")

    def _copy_subject(self) -> None:
        if not self.current or time.monotonic() < self.ready_at:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.subject.get())
        self.status.set("Subject copied. Paste it into the subject field in Zoho Mail.")

    def _mark_sent(self) -> None:
        if not self.current or not self.workbook:
            return
        if not messagebox.askyesno("Confirm status", "Did you manually send this message in Zoho Mail?"):
            return
        try:
            self.workbook.mark_sent(self.current.row_number)
            self.index += 1
            self.progress.configure(value=self.index)
            self.progress_text.set(f"{self.index} of {len(self.recipients)} messages confirmed sent")
            self._show_current()
            if self.index < len(self.recipients):
                self.ready_at = time.monotonic() + int(self.interval_minutes.get()) * 60
                self.open_button.state(["disabled"])
                self.copy_subject_button.state(["disabled"])
                self.copy_button.state(["disabled"])
                if self.pace_job:
                    self.root.after_cancel(self.pace_job)
                self._pace_tick()
        except Exception as exc:
            messagebox.showerror("Could not save status", str(exc))

    def _pace_tick(self) -> None:
        remaining = max(0, int(self.ready_at - time.monotonic() + 0.999))
        if remaining == 0:
            self.open_button.state(["!disabled"])
            self.copy_subject_button.state(["!disabled"])
            self.copy_button.state(["!disabled"])
            self.status.set("Pacing interval complete. Review the next message when ready.")
            self.pace_job = None
            return
        self.status.set(f"Manual pacing interval: {remaining // 60:02d}:{remaining % 60:02d} remaining.")
        self.pace_job = self.root.after(1000, self._pace_tick)

    def _start_automation(self) -> None:
        if not self.workbook or not self.recipients:
            messagebox.showwarning("No workbook", "Please load an Excel or CSV file first.")
            return
        if self.index >= len(self.recipients):
            messagebox.showinfo("Done", "All recipients have already been sent.")
            return

        self.is_auto_running = True
        self.auto_btn.state(["disabled"])
        self.stop_btn.state(["!disabled"])
        self.status.set("Starting automated browser session...")

        import threading
        threading.Thread(target=self._auto_worker, daemon=True).start()

    def _stop_automation(self) -> None:
        self.is_auto_running = False
        self.status.set("Automation stop requested...")
        self.stop_btn.state(["disabled"])

    def _auto_worker(self) -> None:
        from playwright.sync_api import sync_playwright
        import random

        profile_dir = Path(__file__).resolve().parent.parent.parent / "zoho_browser_profile"
        profile_dir.mkdir(exist_ok=True)

        try:
            with sync_playwright() as p:
                context = p.chromium.launch_persistent_context(
                    user_data_dir=str(profile_dir),
                    headless=False,
                    args=["--start-maximized"]
                )
                page = context.pages[0] if context.pages else context.new_page()
                page.goto("https://mail.zoho.com")
                time.sleep(5)

                while self.is_auto_running and self.index < len(self.recipients):
                    current = self.recipients[self.index]
                    subj = current.title
                    template = self.body_template.get("1.0", "end-1c")
                    body = template.replace("{name}", current.name).replace("{email}", current.email).replace("{title}", current.title)

                    self.root.after(0, lambda n=current.name, e=current.email: self.status.set(f"Composing for {n} ({e})..."))

                    # Click compose/new mail
                    for sel in ["button[title*='New Mail']", "button:has-text('New Mail')", "[aria-label='New Mail']", ".zmBtn.zmBtn_primary"]:
                        try:
                            page.wait_for_selector(sel, timeout=2000)
                            page.click(sel)
                            break
                        except Exception:
                            pass
                    time.sleep(1.5)

                    # Fill to
                    for sel in ["input[placeholder*='Add contact']", "input[placeholder*='to']", "#mailTo input", ".select2-search__field"]:
                        try:
                            page.wait_for_selector(sel, timeout=2000)
                            page.fill(sel, current.email)
                            page.keyboard.press("Enter")
                            break
                        except Exception:
                            pass
                    time.sleep(0.5)

                    # Fill subject
                    for sel in ["input[placeholder*='Subject']", "input[aria-label*='Subject']", "input[name='subject']"]:
                        try:
                            page.wait_for_selector(sel, timeout=2000)
                            page.fill(sel, subj)
                            break
                        except Exception:
                            pass
                    time.sleep(0.5)

                    # Fill body
                    for sel in ["div[contenteditable='true']", "div[aria-label*='Message body']", "#mailContent"]:
                        try:
                            page.wait_for_selector(sel, timeout=2000)
                            page.click(sel)
                            page.keyboard.type(body, delay=10)
                            break
                        except Exception:
                            pass
                    time.sleep(1)

                    # Click send
                    for sel in ["button[title*='Send (']", "button:has-text('Send')", "span:text-is('Send')"]:
                        try:
                            page.wait_for_selector(sel, timeout=2500)
                            page.click(sel)
                            break
                        except Exception:
                            pass
                    time.sleep(2)

                    # Mark sent in workbook & UI
                    def _update_success(row_num=current.row_number):
                        self.workbook.mark_sent(row_num)
                        self.index += 1
                        self.progress.configure(value=self.index)
                        self.progress_text.set(f"{self.index} of {len(self.recipients)} messages confirmed sent")
                        self._show_current()

                    self.root.after(0, _update_success)

                    if self.index < len(self.recipients) and self.is_auto_running:
                        base = int(self.interval_minutes.get()) * 60
                        jitter = random.randint(-20, 25)
                        delay = max(30, base + jitter)
                        for sec in range(delay, 0, -1):
                            if not self.is_auto_running:
                                break
                            self.root.after(0, lambda s=sec: self.status.set(f"Auto pacing: {s // 60:02d}:{s % 60:02d} before next recipient..."))
                            time.sleep(1)

                context.close()
        except Exception as e:
            self.root.after(0, lambda err=str(e): messagebox.showerror("Automation Error", err))
        finally:
            self.is_auto_running = False
            self.root.after(0, lambda: self.auto_btn.state(["!disabled"]))
            self.root.after(0, lambda: self.stop_btn.state(["disabled"]))
            self.root.after(0, lambda: self.status.set("Automation completed / stopped."))

    def _close(self) -> None:
        if self.workbook:
            self.workbook.close()
        self.root.destroy()


def main() -> None:
    global DND_FILES
    if TkinterDnD is not None:
        try:
            root = TkinterDnD.Tk()
        except RuntimeError:
            # Keep the desktop app usable if native drop support is unavailable.
            DND_FILES = None
            root = tk.Tk()
    else:
        root = tk.Tk()
    MailMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
