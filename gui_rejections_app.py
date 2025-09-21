# gui_rejections_app.py

import re
import json
import threading
import queue
import webbrowser
from pathlib import Path
from typing import List

import customtkinter as ctk
from tkinter import filedialog, messagebox, Listbox, END, SINGLE

from rejections_core import (
    run_sender,
    DEFAULT_SPREADSHEET_ID,
    DEFAULT_TAB_PREFERRED,
    DEFAULT_READ_RANGE,
)

# Settings file lives alongside this GUI script
APP_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = APP_DIR / "rejections_gui_settings.json"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.title("Sheets → Gmail Rejections | CustomTkinter GUI")
        self.geometry("1020x720")
        self.minsize(940, 640)

        # state
        self.attachments: List[str] = []
        self.worker_thread: threading.Thread | None = None
        self.logq: queue.Queue = queue.Queue()
        self.stop_event = threading.Event()

        # layout
        self._build_layout()
        self._load_settings()
        self._poll_log_queue()

    # ------------- UI construction -------------
    def _build_layout(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)

        self.tab_config = self.tabs.add("Config")
        self.tab_templates = self.tabs.add("Templates")
        self.tab_attachments = self.tabs.add("Attachments")
        self.tab_run = self.tabs.add("Run")

        # ---- Config tab ----
        g = self.tab_config
        for i in range(3):
            g.columnconfigure(i, weight=1)

        self.sender_var = ctk.StringVar()
        self.subject_var = ctk.StringVar(value="Regarding your application for {{ role }} at {{ company }}")
        self.cc_var = ctk.StringVar()
        self.bcc_var = ctk.StringVar()
        self.reply_to_var = ctk.StringVar(value="recruiting@yourcompany.com")
        self.throttle_var = ctk.StringVar(value="2.0")
        self.domain_throttle_var = ctk.StringVar(value="0.0")
        self.preview_n_var = ctk.StringVar(value="0")
        self.dry_run_var = ctk.BooleanVar(value=True)

        # Default creds/token paths next to the GUI script
        self.credentials_var = ctk.StringVar(value=str(APP_DIR / "credentials.json"))
        self.token_var = ctk.StringVar(value=str(APP_DIR / "token.json"))

        self.spreadsheet_id_var = ctk.StringVar(value=DEFAULT_SPREADSHEET_ID)
        self.tab_var = ctk.StringVar(value=DEFAULT_TAB_PREFERRED)
        self.range_var = ctk.StringVar(value=DEFAULT_READ_RANGE)
        self.sender_name_var = ctk.StringVar(value="Recruiting Team")
        self.sender_title_var = ctk.StringVar(value="Talent Acquisition")

        row = 0
        ctk.CTkLabel(g, text="Sender (email address)").grid(row=row, column=0, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.sender_var).grid(row=row + 1, column=0, columnspan=2, sticky="ew", padx=8)
        ctk.CTkLabel(g, text="Subject (Jinja supported)").grid(row=row, column=2, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.subject_var).grid(row=row + 1, column=2, sticky="ew", padx=8)

        row += 2
        ctk.CTkLabel(g, text="Reply-To").grid(row=row, column=0, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.reply_to_var).grid(row=row + 1, column=0, sticky="ew", padx=8)
        ctk.CTkLabel(g, text="Cc (optional)").grid(row=row, column=1, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.cc_var).grid(row=row + 1, column=1, sticky="ew", padx=8)
        ctk.CTkLabel(g, text="Bcc (optional)").grid(row=row, column=2, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.bcc_var).grid(row=row + 1, column=2, sticky="ew", padx=8)

        row += 2
        ctk.CTkLabel(g, text="Throttle (seconds)").grid(row=row, column=0, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.throttle_var).grid(row=row + 1, column=0, sticky="ew", padx=8)
        ctk.CTkLabel(g, text="Domain throttle (seconds)").grid(row=row, column=1, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.domain_throttle_var).grid(row=row + 1, column=1, sticky="ew", padx=8)
        ctk.CTkLabel(g, text="Preview N (0=off)").grid(row=row, column=2, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.preview_n_var).grid(row=row + 1, column=2, sticky="ew", padx=8)

        row += 2
        ctk.CTkLabel(g, text="Spreadsheet ID").grid(row=row, column=0, sticky="w", padx=8, pady=(8, 0))
        ssid_frame = ctk.CTkFrame(g)
        ssid_frame.grid(row=row + 1, column=0, sticky="ew", padx=8)
        ssid_frame.columnconfigure(0, weight=1)
        ctk.CTkEntry(ssid_frame, textvariable=self.spreadsheet_id_var).grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=6)
        ctk.CTkButton(ssid_frame, text="Open Spreadsheet", command=self._open_spreadsheet).grid(row=0, column=1)

        ctk.CTkLabel(g, text="Preferred Tab").grid(row=row, column=1, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.tab_var).grid(row=row + 1, column=1, sticky="ew", padx=8)
        ctk.CTkLabel(g, text="Read Range").grid(row=row, column=2, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, textvariable=self.range_var).grid(row=row + 1, column=2, sticky="ew", padx=8)

        row += 2
        ctk.CTkLabel(g, text="Credentials JSON").grid(row=row, column=0, sticky="w", padx=8, pady=(8, 0))
        cred_frame = ctk.CTkFrame(g)
        cred_frame.grid(row=row + 1, column=0, sticky="ew", padx=8)
        cred_frame.columnconfigure(0, weight=1)
        ctk.CTkEntry(cred_frame, textvariable=self.credentials_var).grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=6)
        ctk.CTkButton(cred_frame, text="Browse", command=self._pick_credentials).grid(row=0, column=1)

        ctk.CTkLabel(g, text="Token JSON (created after consent)").grid(row=row, column=1, sticky="w", padx=8, pady=(8, 0))
        token_frame = ctk.CTkFrame(g)
        token_frame.grid(row=row + 1, column=1, sticky="ew", padx=8)
        token_frame.columnconfigure(0, weight=1)
        ctk.CTkEntry(token_frame, textvariable=self.token_var).grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=6)
        ctk.CTkButton(token_frame, text="Browse", command=self._pick_token).grid(row=0, column=1)

        ctk.CTkCheckBox(g, text="Dry run (no emails sent)", variable=self.dry_run_var).grid(row=row + 1, column=2, sticky="w", padx=8)

        row += 2
        ctk.CTkLabel(g, text="Sender Signature (used if env vars missing)").grid(row=row, column=0, columnspan=3, sticky="w", padx=8, pady=(8, 0))
        ctk.CTkEntry(g, placeholder_text="Sender name", textvariable=self.sender_name_var).grid(row=row + 1, column=0, sticky="ew", padx=8)
        ctk.CTkEntry(g, placeholder_text="Sender title", textvariable=self.sender_title_var).grid(row=row + 1, column=1, sticky="ew", padx=8)
        button_row = ctk.CTkFrame(g)
        button_row.grid(row=row + 1, column=2, sticky="e", padx=8)
        ctk.CTkButton(button_row, text="Save Settings", command=self._save_settings).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(button_row, text="Test Send To Me", command=self._start_test_send).grid(row=0, column=1)

        # ---- Templates tab ----
        t = self.tab_templates
        for i in range(2):
            t.columnconfigure(i, weight=1)
        self.text_template_var = ctk.StringVar(value="template.txt")
        self.html_template_var = ctk.StringVar(value="template.html")

        ctk.CTkLabel(t, text="Text template (required)").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 0))
        tf = ctk.CTkFrame(t)
        tf.grid(row=1, column=0, sticky="ew", padx=8)
        tf.columnconfigure(0, weight=1)
        ctk.CTkEntry(tf, textvariable=self.text_template_var).grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=6)
        ctk.CTkButton(tf, text="Browse", command=self._pick_text_template).grid(row=0, column=1)

        ctk.CTkLabel(t, text="HTML template (optional)").grid(row=0, column=1, sticky="w", padx=8, pady=(8, 0))
        hf = ctk.CTkFrame(t)
        hf.grid(row=1, column=1, sticky="ew", padx=8)
        hf.columnconfigure(0, weight=1)
        ctk.CTkEntry(hf, textvariable=self.html_template_var).grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=6)
        ctk.CTkButton(hf, text="Browse", command=self._pick_html_template).grid(row=0, column=1)

        # ---- Attachments tab ----
        a = self.tab_attachments
        a.columnconfigure(0, weight=1)

        # Use a standard Tk Listbox so we can select & remove items
        self.attach_listbox = Listbox(a, height=16, selectmode=SINGLE)
        self.attach_listbox.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        btns = ctk.CTkFrame(a)
        btns.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        ctk.CTkButton(btns, text="Add attachment(s)", command=self._add_attachments).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(btns, text="Remove selected", command=self._remove_selected).grid(row=0, column=1, padx=(0, 8))
        ctk.CTkButton(btns, text="Remove missing", command=self._remove_missing).grid(row=0, column=2)

        # ---- Run tab ----
        r = self.tab_run
        r.columnconfigure(0, weight=1)
        r.rowconfigure(1, weight=1)
        self.run_buttons = ctk.CTkFrame(r)
        self.run_buttons.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        ctk.CTkButton(self.run_buttons, text="Dry Run", command=self._start_dry_run).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(self.run_buttons, text="Send", fg_color="#0a7", hover_color="#096", command=self._start_send).grid(row=0, column=1)
        ctk.CTkButton(self.run_buttons, text="Cancel", fg_color="#b33", hover_color="#a22", command=self._cancel).grid(row=0, column=2, padx=8)

        self.prog = ctk.CTkProgressBar(r)
        self.prog.grid(row=0, column=0, sticky="ew", padx=8, pady=(48, 0))
        self.prog.set(0)

        self.log = ctk.CTkTextbox(r)
        self.log.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        self.log.configure(state="disabled")

    # ------------- Settings persistence -------------
    def _settings_dict(self):
        return {
            "sender": self.sender_var.get(),
            "subject": self.subject_var.get(),
            "cc": self.cc_var.get(),
            "bcc": self.bcc_var.get(),
            "reply_to": self.reply_to_var.get(),
            "throttle": self.throttle_var.get(),
            "domain_throttle": self.domain_throttle_var.get(),
            "preview_n": self.preview_n_var.get(),
            "dry_run": self.dry_run_var.get(),
            "credentials": self.credentials_var.get(),
            "token": self.token_var.get(),
            "spreadsheet_id": self.spreadsheet_id_var.get(),
            "tab": self.tab_var.get(),
            "read_range": self.range_var.get(),
            "text_template": self.text_template_var.get(),
            "html_template": self.html_template_var.get(),
            "attachments": self.attachments,
            "sender_name": self.sender_name_var.get(),
            "sender_title": self.sender_title_var.get(),
        }

    def _save_settings(self):
        SETTINGS_FILE.write_text(json.dumps(self._settings_dict(), indent=2), encoding="utf-8")
        messagebox.showinfo("Saved", f"Settings saved to {SETTINGS_FILE}")

    def _load_settings(self):
        if SETTINGS_FILE.exists():
            try:
                data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
                self.sender_var.set(data.get("sender", self.sender_var.get()))
                self.subject_var.set(data.get("subject", self.subject_var.get()))
                self.cc_var.set(data.get("cc", ""))
                self.bcc_var.set(data.get("bcc", ""))
                self.reply_to_var.set(data.get("reply_to", self.reply_to_var.get()))
                self.throttle_var.set(str(data.get("throttle", self.throttle_var.get())))
                self.domain_throttle_var.set(str(data.get("domain_throttle", self.domain_throttle_var.get())))
                self.preview_n_var.set(str(data.get("preview_n", self.preview_n_var.get())))
                self.dry_run_var.set(bool(data.get("dry_run", True)))
                self.credentials_var.set(data.get("credentials", self.credentials_var.get()))
                self.token_var.set(data.get("token", self.token_var.get()))
                self.spreadsheet_id_var.set(data.get("spreadsheet_id", self.spreadsheet_id_var.get()))
                self.tab_var.set(data.get("tab", self.tab_var.get()))
                self.range_var.set(data.get("read_range", self.range_var.get()))
                self.text_template_var.set(data.get("text_template", self.text_template_var.get()))
                self.html_template_var.set(data.get("html_template", self.html_template_var.get()))
                self.attachments = list(data.get("attachments", []))
                self.sender_name_var.set(data.get("sender_name", self.sender_name_var.get()))
                self.sender_title_var.set(data.get("sender_title", self.sender_title_var.get()))
                self._refresh_attach_view()
            except Exception as e:
                messagebox.showwarning("Settings", f"Could not load settings: {e}")

    # ------------- File pickers / actions -------------
    def _pick_credentials(self):
        p = filedialog.askopenfilename(title="Pick credentials.json", filetypes=[["JSON Files", "*.json"], ["All Files", "*"]])
        if p:
            self.credentials_var.set(p)

    def _pick_token(self):
        p = filedialog.askopenfilename(title="Pick token.json", filetypes=[["JSON Files", "*.json"], ["All Files", "*"]])
        if p:
            self.token_var.set(p)

    def _pick_text_template(self):
        p = filedialog.askopenfilename(title="Pick text template", filetypes=[["Text/HTML/Markdown", "*.txt *.html *.md"], ["All Files", "*"]])
        if p:
            self.text_template_var.set(p)

    def _pick_html_template(self):
        p = filedialog.askopenfilename(title="Pick HTML template", filetypes=[["HTML", "*.html"], ["All Files", "*"]])
        if p:
            self.html_template_var.set(p)

    def _add_attachments(self):
        paths = filedialog.askopenfilenames(title="Pick attachment(s)")
        if paths:
            for p in paths:
                if p not in self.attachments:
                    self.attachments.append(p)
            self._refresh_attach_view()

    def _remove_selected(self):
        sel = self.attach_listbox.curselection()
        if not sel:
            messagebox.showinfo("Attachments", "Select an item to remove.")
            return
        idx = sel[0]
        if 0 <= idx < len(self.attachments):
            removed = self.attachments.pop(idx)
            self._refresh_attach_view()
            messagebox.showinfo("Attachments", f"Removed:\n{removed}")

    def _remove_missing(self):
        before = len(self.attachments)
        self.attachments = [p for p in self.attachments if Path(p).exists()]
        self._refresh_attach_view()
        messagebox.showinfo("Attachments", f"Removed {before - len(self.attachments)} missing item(s)")

    def _refresh_attach_view(self):
        self.attach_listbox.delete(0, END)
        for p in self.attachments:
            mark = "✓" if Path(p).exists() else "✗"
            self.attach_listbox.insert(END, f"{mark} {p}")

    def _open_spreadsheet(self):
        ssid = self.spreadsheet_id_var.get().strip()
        if ssid:
            webbrowser.open(f"https://docs.google.com/spreadsheets/d/{ssid}")

    # ------------- Run / Worker control -------------
    def _start_dry_run(self):
        self._start_worker(dry=True, test_to_self=False)

    def _start_send(self):
        self._start_worker(dry=False, test_to_self=False)

    def _start_test_send(self):
        self._start_worker(dry=False, test_to_self=True)

    def _cancel(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.stop_event.set()

    def _start_worker(self, dry: bool, test_to_self: bool):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Busy", "A run is already in progress.")
            return
        cfg = self._settings_dict()

        # validate sender address
        if not cfg["sender"] or not re.match(r"[^@]+@[^@]+\.[^@]+", cfg["sender"]):
            messagebox.showerror("Required", "Sender must be a valid email address.")
            return
        if not cfg["text_template"]:
            messagebox.showerror("Required", "Text template file is required.")
            return
        from pathlib import Path as _Path
        if not _Path(cfg["text_template"]).exists():
            messagebox.showerror("Required", "Text template file must exist.")
            return

        # inject runtime flags
        cfg["dry_run"] = dry
        cfg["test_to_self"] = test_to_self

        self._save_settings()
        self._clear_log()
        self.prog.set(0)
        self.stop_event.clear()
        self.worker_thread = threading.Thread(target=run_sender, args=(cfg, self.logq, self.stop_event), daemon=True)
        self.worker_thread.start()

    # ------------- Logging / progress -------------
    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _write_log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _poll_log_queue(self):
        try:
            while True:
                msg = self.logq.get_nowait()
                if isinstance(msg, str) and msg.startswith("__PROG__"):
                    try:
                        frac = msg.split("__PROG__")[1]
                        num_s, denom_s = frac.split("/")
                        num, denom = int(num_s), max(1, int(denom_s))
                        self.prog.set(max(0.01, num / denom))
                    except Exception:
                        pass
                else:
                    self._write_log(msg)
        except queue.Empty:
            pass

        if self.worker_thread and not self.worker_thread.is_alive():
            try:
                self.prog.set(1.0)
            except Exception:
                pass
        self.after(150, self._poll_log_queue)


if __name__ == "__main__":
    app = App()
    app.mainloop()
