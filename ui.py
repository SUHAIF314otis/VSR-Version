from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

from batch_processor import extract_eml_contents
from reconciliation import create_reconciliation_file, validate_reference_files
from soa_extraction import (
    call_llm_for_extraction,
    extract_msg_contents,
    extract_standalone_file_contents,
    get_gemini_api_key,
)


SUPPORTED_EXTENSIONS = (
    ".msg", ".eml", ".pdf", ".xlsx", ".xls", ".xlsm", ".csv",
    ".docx", ".png", ".jpg", ".jpeg", ".bmp", ".webp", ".txt",
)


def runtime_directory():
    """Return a stable application directory in source and PyInstaller builds."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolve_resource_file(filename):
    """Find optional reports beside the script/executable or launch folder."""
    directories = [runtime_directory(), Path.cwd()]
    if hasattr(sys, "_MEIPASS"):
        directories.append(Path(sys._MEIPASS))
    for directory in directories:
        candidate = directory / filename
        if candidate.is_file():
            return str(candidate)
    return ""


def default_output_file():
    documents = Path.home() / "Documents"
    output_directory = documents if documents.is_dir() else runtime_directory()
    return str(output_directory / "Final_Reconciliation_Output.xlsx")


def format_duration(seconds):
    total_seconds = max(0, int(seconds))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours:d}:{minutes:02d}:{seconds:02d}"
    return f"{minutes:02d}:{seconds:02d}"


def calculate_window_geometry(screen_width, screen_height):
    """Choose a usable window size at 100–150% Windows display scaling."""
    screen_width = max(1, int(screen_width))
    screen_height = max(1, int(screen_height))
    window_width = min(screen_width, min(1240, max(680, int(screen_width * 0.96))))
    window_height = min(screen_height, min(860, max(500, int(screen_height * 0.90))))
    left = max(0, (screen_width - window_width) // 2)
    top = max(0, (screen_height - window_height) // 2)
    return window_width, window_height, left, top


def timestamped_output_path(filepath, when=None):
    """Add one fresh YYYYMMDD_HHMMSS suffix to an output workbook name."""
    path = Path(filepath).expanduser().resolve()
    timestamp = (when or datetime.now()).strftime("%Y%m%d_%H%M%S")
    base_name = re.sub(r"_\d{8}_\d{6}$", "", path.stem)
    return path.with_name(f"{base_name}_{timestamp}.xlsx")


def token_report_path(output_path):
    path = Path(output_path)
    return path.with_name(f"{path.stem}_tokens.txt")


def build_token_report(api_usage, totals, started_at, elapsed_seconds, output_path):
    lines = [
        "VSR-U AI TOKEN USAGE",
        "=" * 24,
        f"Run started: {started_at.strftime('%d/%m/%Y %H:%M:%S')}",
        f"Processing time: {format_duration(elapsed_seconds)}",
        f"Final workbook: {output_path}",
        "",
        "API REQUESTS",
        "------------",
    ]
    for number, item in enumerate(api_usage, start=1):
        result = "Completed" if item["status"] == "success" else item["status"].title()
        lines.extend([
            f"{number}. {item['filename']}",
            f"   Result: {result}",
            f"   API called: {item['api_started_at'].strftime('%d/%m/%Y %H:%M:%S')}",
            f"   API time: {format_duration(item['api_seconds'])}",
            f"   Prompt tokens: {item['tokens']['prompt']:,}",
            f"   Output tokens: {item['tokens']['output']:,}",
            f"   Total tokens: {item['tokens']['total']:,}",
            "",
        ])
    lines.extend([
        "TOTAL",
        "-----",
        f"API requests: {len(api_usage)}",
        f"Prompt tokens: {totals['prompt']:,}",
        f"Output tokens: {totals['output']:,}",
        f"Total tokens: {totals['total']:,}",
    ])
    return "\n".join(lines) + "\n"


def open_file_externally(filepath):
    """Open a completed workbook without invoking a command shell."""
    path = os.path.abspath(filepath)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class WorkflowError(RuntimeError):
    pass


class ReconApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OTIS UK — Statement Reconciliation")
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width, window_height, left, top = calculate_window_geometry(
            screen_width, screen_height
        )
        self.geometry(f"{window_width}x{window_height}+{left}+{top}")
        self.minsize(min(680, window_width), min(500, window_height))

        self.colors = {
            "navy": "#17324D",
            "blue": "#0067A0",
            "blue_hover": "#005482",
            "green": "#19734B",
            "green_bg": "#EAF7F0",
            "amber": "#9A6700",
            "amber_bg": "#FFF7DA",
            "red": "#B42318",
            "red_bg": "#FDECEC",
            "text": "#1F2937",
            "muted": "#667085",
            "bg": "#F5F7FA",
            "panel": "#FFFFFF",
            "border": "#D9E1E8",
        }

        self.input_files = []
        self.file_rows = {}
        self.compact_file_rows = {}
        self.input_controls = []
        self.result_df = None
        self.log_lines = []
        self.log_window = None
        self.log_widget = None
        self.ui_queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.inputs_collapsed = False
        self.started_at = None
        self.run_started_at = None
        self.timer_after_id = None

        self.cora_file = tk.StringVar(value=resolve_resource_file("Cora-Report.xlsx"))
        self.jde_file = tk.StringVar(value=resolve_resource_file("JDE Report.xlsx"))
        self.logic_file = tk.StringVar(value=resolve_resource_file("Core-Logic.xlsx"))
        self.output_file = tk.StringVar(value=default_output_file())
        self.cora_display = tk.StringVar()
        self.jde_display = tk.StringVar()
        self.logic_display = tk.StringVar()
        self.output_display = tk.StringVar()
        self.auto_open = tk.BooleanVar(value=True)
        self.progress_value = tk.DoubleVar(value=0)
        self.status_text = tk.StringVar(value="Ready — complete the three steps below.")
        self.file_count = tk.StringVar(value="No statements selected")
        self.result_summary = tk.StringVar(value="No reconciliation has been run yet.")
        self.input_summary = tk.StringVar(value="Inputs are expanded")
        self.processing_time_text = tk.StringVar(value="Processing time: 00:00")
        self.token_usage_text = tk.StringVar(value="AI tokens: —")
        self.save_location_text = tk.StringVar(value="")
        self.token_report_file = ""

        self._configure_styles()
        self._build_ui()
        self._refresh_path_labels()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(50, self._drain_ui_queue)

    def _configure_styles(self):
        self.configure(bg=self.colors["bg"])
        self.option_add("*Font", "{Segoe UI} 9")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=self.colors["bg"])
        style.configure("Panel.TFrame", background=self.colors["panel"])
        style.configure(
            "TLabelframe",
            background=self.colors["panel"],
            bordercolor=self.colors["border"],
            relief="solid",
        )
        style.configure(
            "TLabelframe.Label",
            background=self.colors["panel"],
            foreground=self.colors["navy"],
            font=("Segoe UI", 10, "bold"),
        )
        style.configure("TLabel", background=self.colors["bg"], foreground=self.colors["text"])
        style.configure("Panel.TLabel", background=self.colors["panel"], foreground=self.colors["text"])
        style.configure("Muted.Panel.TLabel", background=self.colors["panel"], foreground=self.colors["muted"])
        style.configure(
            "Primary.TButton",
            background=self.colors["blue"],
            foreground="white",
            borderwidth=0,
            padding=(18, 10),
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[("active", self.colors["blue_hover"]), ("disabled", "#9BBAD3")],
        )
        style.configure("Secondary.TButton", padding=(9, 6))
        style.configure("Danger.TButton", foreground=self.colors["red"], padding=(9, 6))
        body_font = tkfont.Font(root=self, family="Segoe UI", size=9)
        row_height = max(24, body_font.metrics("linespace") + 8)
        style.configure("Treeview", rowheight=row_height, background="white", fieldbackground="white", borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), padding=(7, 8))
        style.map("Treeview", background=[("selected", "#DCEEFF")], foreground=[("selected", self.colors["navy"])])

    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = tk.Frame(self, bg=self.colors["navy"])
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(
            header,
            text="Statement Reconciliation",
            bg=self.colors["navy"],
            fg="white",
            font=("Segoe UI", 17, "bold"),
        ).pack(anchor="w", padx=22, pady=(12, 0))
        tk.Label(
            header,
            text="One simple workflow from supplier statements to the final JDE and Cora result",
            bg=self.colors["navy"],
            fg="#DCE8F1",
        ).pack(anchor="w", padx=22, pady=(2, 10))

        viewport = ttk.Frame(self)
        viewport.grid(row=1, column=0, sticky="nsew")
        viewport.grid_rowconfigure(0, weight=1)
        viewport.grid_columnconfigure(0, weight=1)

        self.content_canvas = tk.Canvas(
            viewport,
            background=self.colors["bg"],
            borderwidth=0,
            highlightthickness=0,
        )
        content_scrollbar = ttk.Scrollbar(
            viewport, orient="vertical", command=self.content_canvas.yview
        )
        self.content_canvas.configure(yscrollcommand=content_scrollbar.set)
        self.content_canvas.grid(row=0, column=0, sticky="nsew")
        content_scrollbar.grid(row=0, column=1, sticky="ns")

        content = ttk.Frame(self.content_canvas, padding=(12, 8, 12, 10))
        self.content_frame = content
        self.content_window = self.content_canvas.create_window(
            (0, 0), window=content, anchor="nw"
        )
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(4, weight=1)

        content.bind("<Configure>", self._sync_content_view)
        self.content_canvas.bind("<Configure>", self._sync_content_view)
        self.bind_all("<MouseWheel>", self._on_content_mousewheel, add="+")

        self._build_input_header(content)
        self._build_statement_step(content)
        self._build_reference_step(content)
        self._build_run_step(content)
        self._build_results(content)

    def _sync_content_view(self, _event=None):
        if not hasattr(self, "content_canvas"):
            return
        canvas_width = max(1, self.content_canvas.winfo_width())
        canvas_height = max(1, self.content_canvas.winfo_height())
        requested_height = max(1, self.content_frame.winfo_reqheight())
        self.content_canvas.itemconfigure(
            self.content_window,
            width=canvas_width,
            height=max(canvas_height, requested_height),
        )
        self.content_canvas.configure(scrollregion=self.content_canvas.bbox("all"))
        if hasattr(self, "save_location_label"):
            self.save_location_label.configure(wraplength=max(280, canvas_width - 70))
        if hasattr(self, "status_label"):
            self.status_label.configure(wraplength=max(280, canvas_width - 70))

    def _on_content_mousewheel(self, event):
        try:
            if event.widget.winfo_toplevel() != self:
                return None
        except tk.TclError:
            return None
        if isinstance(event.widget, ttk.Treeview):
            return None
        bounds = self.content_canvas.bbox("all")
        if not bounds or bounds[3] <= self.content_canvas.winfo_height():
            return None
        direction = -1 if event.delta > 0 else 1
        self.content_canvas.yview_scroll(direction * 3, "units")
        return "break"

    def _build_input_header(self, parent):
        header = ttk.Frame(parent)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(
            header,
            textvariable=self.input_summary,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left")
        self.input_toggle_button = ttk.Button(
            header,
            text="Hide inputs",
            command=self.toggle_inputs,
            style="Secondary.TButton",
        )
        self.input_toggle_button.pack(side="right")

        self.collapsed_file_panel = ttk.LabelFrame(
            parent, text="Statement processing status", padding=6
        )
        self.collapsed_file_panel.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        collapsed_area = ttk.Frame(self.collapsed_file_panel, style="Panel.TFrame")
        collapsed_area.pack(fill="both", expand=True)
        self.compact_file_tree = ttk.Treeview(
            collapsed_area,
            columns=("name", "status"),
            show="headings",
            height=3,
        )
        self.compact_file_tree.heading("name", text="Statement file")
        self.compact_file_tree.heading("status", text="Status")
        self.compact_file_tree.column("name", width=680, anchor="w")
        self.compact_file_tree.column("status", width=130, anchor="center", stretch=False)
        collapsed_scroll = ttk.Scrollbar(
            collapsed_area, orient="vertical", command=self.compact_file_tree.yview
        )
        self.compact_file_tree.configure(yscrollcommand=collapsed_scroll.set)
        self.compact_file_tree.pack(side="left", fill="both", expand=True)
        collapsed_scroll.pack(side="right", fill="y")
        self.collapsed_file_panel.grid_remove()

    def _build_statement_step(self, parent):
        panel = ttk.LabelFrame(parent, text="1  Add supplier statements", padding=9)
        panel.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self.statement_panel = panel
        panel.grid_columnconfigure(1, weight=1)

        actions = ttk.Frame(panel, style="Panel.TFrame")
        actions.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        add_files = ttk.Button(actions, text="Choose files", command=self.select_files, style="Primary.TButton")
        add_folder = ttk.Button(actions, text="Choose folder", command=self.select_folder, style="Secondary.TButton")
        remove_files = ttk.Button(actions, text="Remove selected", command=self.remove_selected, style="Secondary.TButton")
        clear_files = ttk.Button(actions, text="Clear list", command=self.clear_files, style="Danger.TButton")
        add_files.pack(fill="x")
        add_folder.pack(fill="x", pady=(6, 0))
        remove_files.pack(fill="x", pady=(14, 0))
        clear_files.pack(fill="x", pady=(6, 0))
        self.input_controls.extend([add_files, add_folder, remove_files, clear_files])

        file_area = ttk.Frame(panel, style="Panel.TFrame")
        file_area.grid(row=0, column=1, sticky="nsew")
        ttk.Label(file_area, textvariable=self.file_count, style="Muted.Panel.TLabel").pack(anchor="w", pady=(0, 5))
        columns = ("name", "type", "status")
        self.file_tree = ttk.Treeview(file_area, columns=columns, show="headings", height=4)
        self.file_tree.heading("name", text="File name")
        self.file_tree.heading("type", text="Type")
        self.file_tree.heading("status", text="Status")
        self.file_tree.column("name", width=450, anchor="w")
        self.file_tree.column("type", width=70, anchor="center", stretch=False)
        self.file_tree.column("status", width=110, anchor="center", stretch=False)
        scrollbar = ttk.Scrollbar(file_area, orient="vertical", command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        self.file_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _build_reference_step(self, parent):
        panel = ttk.LabelFrame(parent, text="2  Choose the reports used for matching", padding=7)
        panel.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self.reference_panel = panel
        panel.grid_columnconfigure(2, weight=1)

        self._source_row(panel, 0, "Cora report", "Required", self.cora_display, self.browse_cora)
        self._source_row(panel, 1, "JDE supplier ledger", "Required", self.jde_display, self.browse_jde)
        self._source_row(panel, 2, "Logic workbook", "Optional", self.logic_display, self.browse_logic)

    def _source_row(self, parent, row, title, note, display_variable, command):
        ttk.Label(
            parent, text=title, style="Panel.TLabel", font=("Segoe UI", 9, "bold"), width=20
        ).grid(row=row, column=0, sticky="w", padx=(2, 8), pady=3)
        ttk.Label(parent, text=note, style="Muted.Panel.TLabel", width=10).grid(
            row=row, column=1, sticky="w", padx=(0, 8), pady=3
        )
        ttk.Label(
            parent,
            textvariable=display_variable,
            style="Panel.TLabel",
            foreground=self.colors["blue"],
            anchor="w",
            width=24,
        ).grid(row=row, column=2, sticky="ew", padx=(0, 8), pady=3)
        button = ttk.Button(parent, text="Choose", command=command, style="Secondary.TButton")
        button.grid(row=row, column=3, sticky="e", padx=(0, 2), pady=3)
        self.input_controls.append(button)

    def _build_run_step(self, parent):
        panel = ttk.LabelFrame(parent, text="3  Create the final reconciliation workbook", padding=9)
        panel.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        panel.grid_columnconfigure(0, weight=1)

        output_area = ttk.Frame(panel, style="Panel.TFrame")
        output_area.grid(row=0, column=0, columnspan=2, sticky="ew")
        output_area.grid_columnconfigure(1, weight=1)
        ttk.Label(output_area, text="Save as", style="Panel.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(output_area, textvariable=self.output_display, style="Muted.Panel.TLabel", anchor="w", width=24).grid(row=0, column=1, sticky="ew", padx=8)
        output_button = ttk.Button(output_area, text="Change location", command=self.browse_output, style="Secondary.TButton")
        output_button.grid(row=0, column=2, padx=(0, 8))
        auto_open = ttk.Checkbutton(output_area, text="Open when finished", variable=self.auto_open)
        auto_open.grid(row=0, column=3, sticky="e")
        self.input_controls.extend([output_button, auto_open])

        action_area = ttk.Frame(panel, style="Panel.TFrame")
        action_area.grid(row=3, column=1, sticky="e", padx=(14, 0), pady=(5, 0))
        self.run_button = ttk.Button(action_area, text="Run reconciliation", command=self.start_workflow, style="Primary.TButton")
        self.run_button.pack(side="left")
        self.cancel_button = ttk.Button(action_area, text="Cancel", command=self.cancel_workflow, style="Danger.TButton", state="disabled")
        self.cancel_button.pack(side="left", padx=(7, 0))

        self.progress = ttk.Progressbar(panel, variable=self.progress_value, maximum=100)
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(9, 2))
        self.status_label = ttk.Label(panel, textvariable=self.status_text, style="Muted.Panel.TLabel")
        self.status_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 0))

        run_info = ttk.Frame(panel, style="Panel.TFrame")
        run_info.grid(row=3, column=0, sticky="ew", pady=(5, 0))
        ttk.Label(run_info, textvariable=self.processing_time_text, style="Muted.Panel.TLabel").pack(side="left")
        ttk.Label(run_info, text="  |  ", style="Muted.Panel.TLabel").pack(side="left")
        ttk.Label(run_info, textvariable=self.token_usage_text, style="Muted.Panel.TLabel").pack(side="left")
        self.save_location_label = ttk.Label(
            panel,
            textvariable=self.save_location_text,
            style="Muted.Panel.TLabel",
            wraplength=820,
        )
        self.save_location_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=(3, 0))

    def _build_results(self, parent):
        panel = ttk.LabelFrame(parent, text="Results", padding=8)
        panel.grid(row=4, column=0, sticky="nsew")
        panel.grid_rowconfigure(1, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        toolbar = ttk.Frame(panel, style="Panel.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        toolbar.grid_columnconfigure(0, weight=1)
        ttk.Label(toolbar, textvariable=self.result_summary, style="Panel.TLabel", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        self.open_button = ttk.Button(toolbar, text="Open final workbook", command=self.open_result, style="Secondary.TButton", state="disabled")
        self.open_button.grid(row=0, column=2, sticky="e")
        ttk.Button(toolbar, text="View activity", command=self.show_activity, style="Secondary.TButton").grid(row=0, column=1, sticky="e", padx=(6, 6))

        table_area = ttk.Frame(panel, style="Panel.TFrame")
        table_area.grid(row=1, column=0, sticky="nsew")
        table_area.grid_rowconfigure(0, weight=1)
        table_area.grid_columnconfigure(0, weight=1)
        columns = ("invoice", "amount", "status", "pending", "source")
        self.result_tree = ttk.Treeview(table_area, columns=columns, show="headings")
        headings = {
            "invoice": "Invoice number",
            "amount": "Amount",
            "status": "Invoice Status",
            "pending": "Pending with",
            "source": "Statement source",
        }
        widths = {"invoice": 135, "amount": 100, "status": 410, "pending": 145, "source": 210}
        for column in columns:
            self.result_tree.heading(column, text=headings[column])
            self.result_tree.column(column, width=widths[column], anchor="e" if column == "amount" else "w", stretch=column in {"status", "source"})
        self.result_tree.tag_configure("normal", background="#F2FAF5")
        self.result_tree.tag_configure("progress", background="#FFF9E8")
        self.result_tree.tag_configure("review", background="#FFF0F0")
        scroll_y = ttk.Scrollbar(table_area, orient="vertical", command=self.result_tree.yview)
        scroll_x = ttk.Scrollbar(table_area, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

    def select_files(self):
        paths = filedialog.askopenfilenames(
            title="Choose supplier statements",
            filetypes=[
                ("Supported statements", "*.pdf *.msg *.eml *.xlsx *.xls *.xlsm *.csv *.docx *.png *.jpg *.jpeg *.bmp *.webp *.txt"),
                ("All files", "*.*"),
            ],
        )
        self._add_files(paths)

    def select_folder(self):
        folder = filedialog.askdirectory(title="Choose a folder containing supplier statements")
        if not folder:
            return
        paths = []
        for root, _directories, filenames in os.walk(folder):
            for filename in filenames:
                if Path(filename).suffix.lower() in SUPPORTED_EXTENSIONS:
                    paths.append(os.path.join(root, filename))
        self._add_files(paths)

    def _add_files(self, paths):
        for path in paths:
            normalized = os.path.abspath(path)
            if normalized not in self.input_files:
                self.input_files.append(normalized)
        self._refresh_file_tree()

    def remove_selected(self):
        selected = {self.file_rows[item] for item in self.file_tree.selection() if item in self.file_rows}
        self.input_files = [path for path in self.input_files if path not in selected]
        self._refresh_file_tree()

    def clear_files(self):
        self.input_files.clear()
        self._refresh_file_tree()

    def _refresh_file_tree(self):
        self.file_tree.delete(*self.file_tree.get_children())
        self.compact_file_tree.delete(*self.compact_file_tree.get_children())
        self.file_rows.clear()
        self.compact_file_rows.clear()
        for path in self.input_files:
            item = self.file_tree.insert("", "end", values=(Path(path).name, Path(path).suffix.upper(), "Ready"))
            self.file_rows[item] = path
            compact_item = self.compact_file_tree.insert(
                "", "end", values=(Path(path).name, "Ready")
            )
            self.compact_file_rows[compact_item] = path
        count = len(self.input_files)
        self.file_count.set(f"{count} statement{'s' if count != 1 else ''} selected" if count else "No statements selected")
        self._refresh_input_summary()

    def _set_file_status(self, path, status):
        for item, item_path in self.file_rows.items():
            if item_path == path and self.file_tree.exists(item):
                values = list(self.file_tree.item(item, "values"))
                values[-1] = status
                self.file_tree.item(item, values=values)
                break
        for item, item_path in self.compact_file_rows.items():
            if item_path == path and self.compact_file_tree.exists(item):
                values = list(self.compact_file_tree.item(item, "values"))
                values[-1] = status
                self.compact_file_tree.item(item, values=values)
                break

    def browse_cora(self):
        self._browse_input(self.cora_file, "Choose Cora report")

    def browse_jde(self):
        self._browse_input(self.jde_file, "Choose JDE supplier ledger")

    def browse_logic(self):
        self._browse_input(self.logic_file, "Choose optional logic workbook")

    def _browse_input(self, variable, title):
        path = filedialog.askopenfilename(title=title, filetypes=[("Excel workbooks", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            variable.set(path)
            self._refresh_path_labels()

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save final reconciliation workbook",
            defaultextension=".xlsx",
            initialfile=Path(self.output_file.get()).name,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if path:
            self.output_file.set(path)
            self._refresh_path_labels()

    def _refresh_path_labels(self):
        self.cora_display.set(Path(self.cora_file.get()).name if self.cora_file.get() else "Not selected")
        self.jde_display.set(Path(self.jde_file.get()).name if self.jde_file.get() else "Not selected")
        self.logic_display.set(Path(self.logic_file.get()).name if self.logic_file.get() else "Not selected")
        self.output_display.set(Path(self.output_file.get()).name if self.output_file.get() else "Not selected")
        self.save_location_text.set(
            f"Saving to: {self.output_file.get()}" if self.output_file.get() else "Saving to: Not selected"
        )
        self._refresh_input_summary()

    def _refresh_input_summary(self):
        file_text = f"{len(self.input_files)} statement{'s' if len(self.input_files) != 1 else ''}"
        cora_text = "Cora selected" if self.cora_file.get() else "Cora needed"
        jde_text = "JDE selected" if self.jde_file.get() else "JDE needed"
        self.input_summary.set(f"Inputs: {file_text}  •  {cora_text}  •  {jde_text}")

    def toggle_inputs(self):
        self._set_inputs_collapsed(not self.inputs_collapsed)

    def _set_inputs_collapsed(self, collapsed):
        self.inputs_collapsed = collapsed
        if collapsed:
            self.statement_panel.grid_remove()
            self.reference_panel.grid_remove()
            self.collapsed_file_panel.grid()
            self.input_toggle_button.configure(text="Change files")
        else:
            self.collapsed_file_panel.grid_remove()
            self.statement_panel.grid()
            self.reference_panel.grid()
            self.input_toggle_button.configure(text="Hide inputs")

    def _validate_setup(self):
        if not self.input_files:
            raise WorkflowError("Choose at least one supplier statement in Step 1.")
        if not get_gemini_api_key():
            raise WorkflowError("The team Gemini configuration is unavailable.")

        cora = self.cora_file.get().strip()
        jde = self.jde_file.get().strip()
        logic = self.logic_file.get().strip()
        output = self.output_file.get().strip()
        if not cora or not os.path.isfile(cora):
            raise WorkflowError("Choose a valid Cora report in Step 2.")
        if not jde or not os.path.isfile(jde):
            raise WorkflowError("Choose a valid JDE supplier ledger in Step 2.")
        if logic and not os.path.isfile(logic):
            raise WorkflowError("The optional logic workbook cannot be found.")
        if not output or Path(output).suffix.lower() != ".xlsx":
            raise WorkflowError("Choose a valid Excel output location in Step 3.")

        output_path = timestamped_output_path(output)
        if not output_path.parent.is_dir():
            raise WorkflowError("The selected output folder does not exist.")
        source_paths = [*self.input_files, cora, jde, *([logic] if logic else [])]
        if any(Path(path).expanduser().resolve() == output_path for path in source_paths):
            raise WorkflowError("Choose a different output name. Input files cannot be overwritten.")
        self.output_file.set(str(output_path))
        self.token_report_file = str(token_report_path(output_path))
        self.output_display.set(output_path.name)
        self.save_location_text.set(
            f"Workbook: {output_path}\nToken report: {self.token_report_file}"
        )
        return list(self.input_files), cora, jde, logic or None, str(output_path)

    def start_workflow(self):
        try:
            arguments = self._validate_setup()
        except WorkflowError as exc:
            messagebox.showwarning("Check your selections", str(exc))
            return

        self.cancel_event.clear()
        self._set_inputs_collapsed(True)
        self._set_controls_running(True)
        self.cancel_button.configure(state="normal")
        self.open_button.configure(state="disabled")
        self.progress_value.set(0)
        self.processing_time_text.set("Processing time: 00:00")
        self.token_usage_text.set("AI tokens: 0")
        self.started_at = time.monotonic()
        self.run_started_at = datetime.now()
        self._schedule_timer_update()
        self.status_text.set("Checking the Cora and JDE reports...")
        self.status_label.configure(foreground=self.colors["muted"])
        self.result_summary.set("Working…")
        self.result_df = None
        self.result_tree.delete(*self.result_tree.get_children())
        self.log_lines.clear()
        self._refresh_file_tree()
        self._log("Starting end-to-end reconciliation")

        threading.Thread(target=self._run_workflow, args=(arguments,), daemon=True).start()

    def cancel_workflow(self):
        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self.status_text.set("Cancelling after the current file finishes...")
        self._log("Cancellation requested.")

    def _run_workflow(self, arguments):
        files, cora, jde, logic, output = arguments
        workflow_started = time.monotonic()
        run_started_at = self.run_started_at or datetime.now()
        tokens = {"prompt": 0, "output": 0, "total": 0}
        api_usage = []
        try:
            validate_reference_files(cora, jde)
            if self.cancel_event.is_set():
                raise WorkflowError("Run cancelled. No new workbook was created.")
            self._post_ui(self._set_progress, 8, "Reports are valid. Reading supplier statements...")
            self._log("Cora and JDE validation passed.")

            records = []
            failures = []
            missed = []
            with ThreadPoolExecutor(max_workers=min(4, len(files))) as executor:
                futures = {executor.submit(self._process_file, path): path for path in files}
                for completed, future in enumerate(as_completed(futures), start=1):
                    if self.cancel_event.is_set():
                        for pending in futures:
                            pending.cancel()
                        raise WorkflowError("Run cancelled. No new workbook was created.")
                    path = futures[future]
                    result = future.result()
                    if result["api_called"]:
                        api_usage.append({
                            "filename": result["filename"],
                            "status": result["status"],
                            "api_started_at": result["api_started_at"],
                            "api_seconds": result["api_seconds"],
                            "tokens": result["tokens"].copy(),
                        })
                    for key in tokens:
                        tokens[key] += result["tokens"][key]
                    self._post_ui(self._set_token_usage, tokens["total"])
                    if result["status"] == "success":
                        records.extend(result["records"])
                    elif result["status"] == "missed":
                        missed.append(result)
                    else:
                        failures.append(result)
                    self._log(f"{Path(path).name}: {result['message']}")
                    self._post_ui(self._set_file_status, path, result["status"].title())
                    self._post_ui(
                        self._set_progress,
                        8 + (completed / len(files)) * 57,
                        f"Read {completed} of {len(files)} statement files...",
                    )

            if failures:
                names = ", ".join(item["filename"] for item in failures[:5])
                more = "" if len(failures) <= 5 else f" and {len(failures) - 5} more"
                raise WorkflowError(
                    f"Could not read {len(failures)} statement file(s): {names}{more}. Open View activity for details."
                )
            if not records:
                raise WorkflowError("No invoices were found in the selected statements.")

            extracted_df = pd.DataFrame(records, columns=[
                "Date", "Invoice Number", "Amount( Original )", "Open amount",
                "Invoice Status", "Pending with", "Source File",
            ])
            self._post_ui(self._set_progress, 70, "Matching invoices with JDE and Cora...")
            final_df = create_reconciliation_file(
                extracted_df, cora, jde, logic, output, logger=self._log
            )
            if not os.path.isfile(output):
                raise WorkflowError("The final workbook could not be created.")
            elapsed = time.monotonic() - workflow_started
            report_file = token_report_path(output)
            report_text = build_token_report(
                api_usage, tokens, run_started_at, elapsed, output
            )
            self._write_token_report(report_file, report_text)
            self._post_ui(
                self._finish_success,
                final_df,
                output,
                tokens,
                len(files),
                len(missed),
                str(report_file),
                elapsed,
            )
        except Exception as exc:
            self._log(f"Run stopped: {exc}")
            self._post_ui(self._finish_error, str(exc), tokens)

    def _process_file(self, file_path):
        filename = os.path.basename(file_path)
        extension = Path(filename).suffix.lower()
        self._post_ui(self._set_file_status, file_path, "Reading")
        result = {
            "filename": filename,
            "status": "success",
            "records": [],
            "message": "",
            "tokens": {"prompt": 0, "output": 0, "total": 0},
            "api_called": False,
            "api_started_at": None,
            "api_seconds": 0,
        }
        api_started = None
        try:
            if extension == ".msg":
                parsed = extract_msg_contents(file_path)
            elif extension == ".eml":
                parsed = extract_eml_contents(file_path)
            elif extension in SUPPORTED_EXTENSIONS:
                parsed = extract_standalone_file_contents(file_path)
            else:
                result["status"] = "missed"
                result["message"] = f"Unsupported file type: {extension}"
                return result

            api_started = time.monotonic()
            result["api_called"] = True
            result["api_started_at"] = datetime.now()
            extracted, usage = call_llm_for_extraction(parsed)
            result["api_seconds"] = time.monotonic() - api_started
            result["tokens"] = {
                "prompt": usage.get("prompt_tokens", 0),
                "output": usage.get("output_tokens", 0),
                "total": usage.get("total_tokens", 0),
            }
            for item in extracted.line_items:
                source = filename
                if extension in {".msg", ".eml"}:
                    source = f"{filename}/{item.source_document or 'Email Body'}"
                result["records"].append({
                    "Date": item.invoice_date or "",
                    "Invoice Number": item.invoice_number or "",
                    "Amount( Original )": item.invoice_value if item.invoice_value is not None else "",
                    "Open amount": item.open_amount if item.open_amount is not None else "",
                    "Invoice Status": "",
                    "Pending with": "",
                    "Source File": source,
                })
            if result["records"]:
                result["message"] = f"Found {len(result['records'])} invoice row(s)."
            else:
                result["status"] = "missed"
                result["message"] = "No invoice rows found."
        except Exception as exc:
            if api_started is not None:
                result["api_seconds"] = time.monotonic() - api_started
            result["status"] = "failed"
            result["message"] = str(exc)
        return result

    @staticmethod
    def _write_token_report(report_file, report_text):
        report_path = Path(report_file)
        temporary_path = report_path.with_name(f".{report_path.name}.tmp")
        try:
            temporary_path.write_text(report_text, encoding="utf-8")
            os.replace(temporary_path, report_path)
        finally:
            if temporary_path.exists():
                temporary_path.unlink()

    def _finish_success(self, final_df, output, tokens, file_count, missed_count, report_file, elapsed):
        self._stop_timer()
        self.result_df = final_df
        review_prefixes = (
            "Missing in ledger", "Matched - Manual", "Need Manual", "Rejected", "Cancelled",
        )
        review_mask = final_df["Invoice Status"].str.startswith(review_prefixes, na=False)
        review_count = int(review_mask.sum())
        completed_count = len(final_df) - review_count
        suffix = f" • {missed_count} file(s) contained no invoices" if missed_count else ""
        self.result_summary.set(
            f"{len(final_df)} invoices  •  {completed_count} progressing normally  •  {review_count} need review{suffix}"
        )
        self._populate_results(final_df)
        self.progress_value.set(100)
        self.status_text.set(f"Complete — final workbook saved as {Path(output).name}")
        self.status_label.configure(foreground=self.colors["green"])
        self.processing_time_text.set(f"Processing time: {format_duration(elapsed)}")
        self._set_token_usage(tokens["total"])
        self.token_report_file = report_file
        self.save_location_text.set(
            f"Workbook: {output}\nToken report: {report_file}"
        )
        self._set_controls_running(False)
        self.cancel_button.configure(state="disabled")
        self.open_button.configure(state="normal")
        self._log(
            f"Complete. Files: {file_count}; invoices: {len(final_df)}; "
            f"needs review: {review_count}; time: {format_duration(elapsed)}; "
            f"AI tokens: {tokens['total']:,}; workbook: {output}; token report: {report_file}."
        )
        if self.auto_open.get():
            self.open_result()

    def _finish_error(self, message, tokens):
        elapsed = self._stop_timer()
        self.progress_value.set(0)
        self.status_text.set(f"Stopped — {message}")
        self.status_label.configure(foreground=self.colors["red"])
        self.result_summary.set("No new result was created. Correct the issue and run again.")
        self.processing_time_text.set(f"Processing time: {format_duration(elapsed)}")
        self._set_token_usage(tokens["total"])
        self._set_controls_running(False)
        self._set_inputs_collapsed(False)
        self.cancel_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        messagebox.showerror("Reconciliation stopped", message)

    def _populate_results(self, dataframe):
        self.result_tree.delete(*self.result_tree.get_children())
        for _, row in dataframe.iterrows():
            status = str(row.get("Invoice Status", ""))
            if status.startswith(("Paid in full", "Ready for Payment")):
                tag = "normal"
            elif status.startswith(("Pending Approval", "Under Processing", "Partially Paid")):
                tag = "progress"
            else:
                tag = "review"
            amount = row.get("Amount( Original )")
            amount_text = f"£{float(amount):,.2f}" if pd.notna(amount) and amount != "" else ""
            self.result_tree.insert("", "end", values=(
                row.get("Invoice Number", ""),
                amount_text,
                status,
                row.get("Pending with", ""),
                row.get("Source File", ""),
            ), tags=(tag,))

    def open_result(self):
        output = self.output_file.get().strip()
        try:
            open_file_externally(output)
        except Exception as exc:
            messagebox.showerror("Cannot open workbook", str(exc))

    def show_activity(self):
        if self.log_window and self.log_window.winfo_exists():
            self.log_window.lift()
            return
        self.log_window = tk.Toplevel(self)
        self.log_window.title("Reconciliation activity")
        self.log_window.geometry("760x440")
        self.log_window.minsize(560, 320)
        self.log_widget = scrolledtext.ScrolledText(
            self.log_window,
            wrap="word",
            font=("Consolas", 9),
            bg="#172B4D",
            fg="#EAF2F8",
            insertbackground="white",
            padx=10,
            pady=10,
        )
        self.log_widget.pack(fill="both", expand=True)
        self.log_widget.insert("1.0", "\n".join(self.log_lines) or "No activity yet.")
        self.log_widget.configure(state="disabled")
        self.log_window.protocol("WM_DELETE_WINDOW", self._close_activity)

    def _close_activity(self):
        if self.log_window:
            self.log_window.destroy()
        self.log_window = None
        self.log_widget = None

    def _set_progress(self, value, message):
        self.progress_value.set(value)
        self.status_text.set(message)
        self.status_label.configure(foreground=self.colors["muted"])

    def _set_token_usage(self, total_tokens):
        self.token_usage_text.set(f"AI tokens: {int(total_tokens):,}")

    def _schedule_timer_update(self):
        if self.started_at is None:
            return
        elapsed = time.monotonic() - self.started_at
        self.processing_time_text.set(f"Processing time: {format_duration(elapsed)}")
        self.timer_after_id = self.after(1000, self._schedule_timer_update)

    def _stop_timer(self):
        elapsed = time.monotonic() - self.started_at if self.started_at is not None else 0
        if self.timer_after_id is not None:
            try:
                self.after_cancel(self.timer_after_id)
            except tk.TclError:
                pass
        self.timer_after_id = None
        self.started_at = None
        return elapsed

    def _set_controls_running(self, running):
        state = "disabled" if running else "normal"
        self.run_button.configure(state=state)
        for control in self.input_controls:
            control.configure(state=state)

    def _log(self, message):
        self._post_ui(self._append_log, str(message))

    def _append_log(self, message):
        self.log_lines.append(message)
        if self.log_widget and self.log_widget.winfo_exists():
            self.log_widget.configure(state="normal")
            self.log_widget.insert(tk.END, message + "\n")
            self.log_widget.see(tk.END)
            self.log_widget.configure(state="disabled")

    def _post_ui(self, callback, *args):
        self.ui_queue.put((callback, args))

    def _drain_ui_queue(self):
        while True:
            try:
                callback, args = self.ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                callback(*args)
            except (tk.TclError, RuntimeError) as exc:
                print(f"UI update failed: {exc}")
        if self.winfo_exists():
            self.after(50, self._drain_ui_queue)

    def _on_close(self):
        self.cancel_event.set()
        if self.timer_after_id is not None:
            try:
                self.after_cancel(self.timer_after_id)
            except tk.TclError:
                pass
        self.destroy()


if __name__ == "__main__":
    ReconApp().mainloop()
