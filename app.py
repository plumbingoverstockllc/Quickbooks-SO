from __future__ import annotations

import json
import logging
import logging.handlers
import os
import re
import hashlib
import threading
import subprocess
import tempfile
import traceback
import urllib.parse
import tkinter as tk
import urllib.request
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from quickbooks_client import QuickBooksClient, _current_process_elevation
from transformer import OrderSettings, line_pricing_keys, load_source, transform_to_template, unique_brands, unique_skus


DEFAULT_SOURCE = r"C:\Users\QB-PC\Downloads\Project-LisaStrongDesign-EliezerLabkowski301NHighland (1).xls"
DEFAULT_TEMPLATE = r"C:\Users\QB-PC\Downloads\SaasAnt Template for David Meyer.xlsx"
DEFAULT_OUTPUT = r"C:\Users\QB-PC\Downloads\SaaSant Sales Order - Auto Filled.xlsx"
APP_NAME = "QB Sales Order Converter"
APP_VERSION = "v0.9.803 Beta"
UPDATE_API_URL = "https://api.github.com/repos/plumbingoverstockllc/Quickbooks-SO/releases/latest"
UPDATE_INFO_URL = "https://raw.githubusercontent.com/plumbingoverstockllc/Quickbooks-SO/main/releases/latest.json"
SETTINGS_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / APP_NAME
SETTINGS_PATH = SETTINGS_DIR / "settings.json"
LOG_PATH = SETTINGS_DIR / "app.log"


def _setup_logging() -> logging.Logger:
    SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("qb_so_app")
    logger.setLevel(logging.DEBUG)
    if not logger.handlers:
        fmt = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            "%Y-%m-%d %H:%M:%S",
        )
        handler = logging.handlers.RotatingFileHandler(
            LOG_PATH, maxBytes=512_000, backupCount=2, encoding="utf-8"
        )
        handler.setFormatter(fmt)
        logger.addHandler(handler)
    logger.propagate = False
    return logger


log = _setup_logging()
log.info("=" * 60)
log.info("App starting: %s %s (PID %s)", APP_NAME, APP_VERSION, os.getpid())
log.info("Settings dir: %s", SETTINGS_DIR)
log.info("Log file: %s", LOG_PATH)

UI = {
    "bg_window": "#F4F6FA",
    "bg_card": "#FFFFFF",
    "bg_subtle": "#EEF2F7",
    "bg_hover": "#E3EAF3",
    "bg_pressed": "#D6E0EE",
    "border": "#DCE3EE",
    "border_strong": "#BFC8D7",
    "border_inner_light": "#FFFFFF",
    "text_primary": "#0F172A",
    "text_secondary": "#475569",
    "text_tertiary": "#94A3B8",
    "accent": "#2563EB",
    "accent_light": "#3B82F6",
    "accent_hover": "#1D4ED8",
    "accent_pressed": "#1E40AF",
    "accent_dark": "#1E3A8A",
    "accent_bg": "#DBEAFE",
    "btn_light_top": "#FFFFFF",
    "btn_light_bottom": "#F4F6FA",
    "btn_light_border": "#DCE3EE",
    "success": "#15803D",
    "success_bg": "#E5F5EC",
    "warning": "#B45309",
    "warning_bg": "#FEF3C7",
    "danger": "#B91C1C",
    "danger_bg": "#FEE2E2",
}


class PricingRulesDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        source_df,
        pricing_mode: str = "brand",
        use_actual_cost: bool = False,
        default_value: float = 0.4,
        existing_brand_values: dict[str, float] | None = None,
        existing_item_values: dict[str, float] | None = None,
        existing_line_values: dict[str, float] | None = None,
    ) -> None:
        super().__init__(parent)
        self.title("Pricing Rules")
        self.geometry("560x620")
        self.resizable(True, True)
        self.configure(bg=UI["bg_window"])
        self.result = None
        self.source_df = source_df
        self.existing_brand_values = existing_brand_values or {}
        self.existing_item_values = existing_item_values or {}
        self.existing_line_values = existing_line_values or {}

        self.pricing_mode_var = tk.StringVar(value=pricing_mode)
        self.use_actual_cost_var = tk.BooleanVar(value=use_actual_cost)
        self.default_var = tk.StringVar(value=str(default_value))

        ttk.Label(self, text="Pricing Rules", style="Header.TLabel").pack(
            anchor="w", padx=18, pady=(18, 2)
        )
        ttk.Label(
            self,
            text="Choose how line item rates are calculated.",
            style="SubHeader.TLabel",
        ).pack(anchor="w", padx=18, pady=(0, 12))

        mode_frame = ttk.Frame(self)
        mode_frame.pack(fill="x", padx=18, pady=(0, 6))
        ttk.Label(mode_frame, text="Pricing Mode").pack(side="left")
        ttk.Radiobutton(
            mode_frame,
            text="Per Brand",
            value="brand",
            variable=self.pricing_mode_var,
            command=self._render_value_rows,
        ).pack(side="left", padx=(12, 8))
        ttk.Radiobutton(
            mode_frame,
            text="Per Item (SKU)",
            value="item",
            variable=self.pricing_mode_var,
            command=self._render_value_rows,
        ).pack(side="left", padx=(0, 8))
        ttk.Radiobutton(
            mode_frame,
            text="Per Line (Excel Row)",
            value="line",
            variable=self.pricing_mode_var,
            command=self._render_value_rows,
        ).pack(side="left")

        ttk.Checkbutton(
            self,
            text="Use actual cost (typed value is final rate, not multiplier)",
            variable=self.use_actual_cost_var,
            command=self._update_default_label,
        ).pack(anchor="w", padx=18, pady=(8, 10))

        self.default_label_var = tk.StringVar()
        self._update_default_label()
        ttk.Label(self, textvariable=self.default_label_var).pack(anchor="w", padx=18, pady=(0, 4))
        ttk.Entry(self, textvariable=self.default_var).pack(fill="x", padx=18, pady=(0, 14))

        wrapper = tk.Frame(
            self,
            bg=UI["bg_card"],
            highlightbackground=UI["border"],
            highlightthickness=1,
            bd=0,
        )
        wrapper.pack(fill="both", expand=True, padx=18, pady=(0, 14))
        canvas = tk.Canvas(wrapper, highlightthickness=0, bg=UI["bg_card"], bd=0)
        scroll = ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas, style="Card.TFrame")
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scroll.pack(side="right", fill="y", pady=8)

        self.value_vars: dict[str, tk.StringVar] = {}
        self._render_value_rows()

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", padx=18, pady=(0, 16))
        ttk.Button(buttons, text="Use Pricing Rules", command=self._submit, style="Primary.TButton").pack(
            side="right"
        )
        ttk.Button(buttons, text="Cancel", command=self._cancel, style="Quiet.TButton").pack(
            side="right", padx=(0, 8)
        )

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _update_default_label(self) -> None:
        if self.use_actual_cost_var.get():
            self.default_label_var.set("Default Cost/Rate:")
        else:
            self.default_label_var.set("Default Multiplier (MSRP x multiplier):")

    def _render_value_rows(self) -> None:
        for child in self.inner.winfo_children():
            child.destroy()

        self.value_vars = {}
        mode = self.pricing_mode_var.get()
        if mode == "brand":
            keys = [(k, k) for k in unique_brands(self.source_df)]
            values_source = self.existing_brand_values
        elif mode == "item":
            keys = [(k, k) for k in unique_skus(self.source_df)]
            values_source = self.existing_item_values
        else:
            keys = line_pricing_keys(self.source_df)
            values_source = self.existing_line_values

        for key, label in keys:
            row = ttk.Frame(self.inner, style="Card.TFrame")
            row.pack(fill="x", pady=3, padx=10)
            ttk.Label(row, text=label, width=34, style="Card.TLabel").pack(side="left")
            existing_value = ""
            if key in values_source:
                existing_value = str(values_source.get(key, ""))
            var = tk.StringVar(value=existing_value)
            self.value_vars[key] = var
            ttk.Entry(row, textvariable=var, width=16).pack(side="right")

    def _submit(self) -> None:
        try:
            default_value = float(self.default_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid input", "Default value must be a number.")
            return

        parsed_values = {}
        for key, var in self.value_vars.items():
            raw = var.get().strip()
            if not raw:
                parsed_values[key] = default_value
                continue
            try:
                parsed_values[key] = float(raw)
            except ValueError:
                messagebox.showerror("Invalid input", f"Invalid value for: {key}")
                return

        self.result = (
            self.pricing_mode_var.get(),
            self.use_actual_cost_var.get(),
            default_value,
            parsed_values,
        )
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class RowEditorDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, source_row: dict, output_row: dict, excel_line_number: int) -> None:
        super().__init__(parent)
        self.title(f"Edit Row (Excel Line {excel_line_number})")
        self.geometry("880x640")
        self.resizable(True, True)
        self.configure(bg=UI["bg_window"])
        self.result = None
        self._output_canvas: tk.Canvas | None = None
        self._source_canvas: tk.Canvas | None = None

        header_box = ttk.Frame(self, padding=(20, 18, 20, 4))
        header_box.pack(fill="x")
        ttk.Label(header_box, text=f"Edit Row · Excel line {excel_line_number}", style="Header.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            header_box,
            text=(
                f"SO {output_row.get('Sales Order No', '') or '—'}  ·  "
                f"Customer: {output_row.get('Customer', '') or '—'}"
            ),
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        wrapper = ttk.Frame(self, padding=(20, 10, 20, 0))
        wrapper.pack(fill="both", expand=True)

        notebook = ttk.Notebook(wrapper)
        notebook.pack(fill="both", expand=True)
        output_tab = ttk.Frame(notebook, style="Card.TFrame")
        source_tab = ttk.Frame(notebook, style="Card.TFrame")
        notebook.add(output_tab, text="Output Preview")
        notebook.add(source_tab, text="Source Fields")
        notebook.select(output_tab)

        self.source_vars = {}
        self.output_vars = {}

        ttk.Label(
            output_tab,
            text=f"Output values for Excel line {excel_line_number}",
            style="CardSubHeader.TLabel",
        ).pack(anchor="w", pady=(12, 10), padx=14)
        out_canvas = tk.Canvas(output_tab, highlightthickness=0, bg=UI["bg_card"], bd=0)
        out_scroll = ttk.Scrollbar(output_tab, orient="vertical", command=out_canvas.yview)
        out_inner = ttk.Frame(out_canvas, style="Card.TFrame")
        out_inner.bind("<Configure>", lambda e: out_canvas.configure(scrollregion=out_canvas.bbox("all")))
        out_canvas.create_window((0, 0), window=out_inner, anchor="nw")
        out_canvas.configure(yscrollcommand=out_scroll.set)
        out_canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
        out_scroll.pack(side="right", fill="y")
        self._output_canvas = out_canvas

        for col, value in output_row.items():
            value = output_row.get(col, "")
            row = ttk.Frame(out_inner, style="Card.TFrame")
            row.pack(fill="x", pady=3, padx=4)
            ttk.Label(row, text=col, width=34, style="Card.TLabel").pack(side="left")
            var = tk.StringVar(value="" if value is None else str(value))
            self.output_vars[col] = var
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

        ttk.Label(
            source_tab,
            text=f"Edit source fields for Excel line {excel_line_number}",
            style="CardSubHeader.TLabel",
        ).pack(anchor="w", pady=(12, 10), padx=14)
        canvas = tk.Canvas(source_tab, highlightthickness=0, bg=UI["bg_card"], bd=0)
        scrollbar = ttk.Scrollbar(source_tab, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="Card.TFrame")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
        scrollbar.pack(side="right", fill="y")
        self._source_canvas = canvas

        for col, value in source_row.items():
            row = ttk.Frame(inner, style="Card.TFrame")
            row.pack(fill="x", pady=3, padx=4)
            ttk.Label(row, text=col, width=34, style="Card.TLabel").pack(side="left")
            var = tk.StringVar(value="" if value is None else str(value))
            self.source_vars[col] = var
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

        btns = ttk.Frame(self, padding=(20, 12, 20, 16))
        btns.pack(fill="x")
        ttk.Button(btns, text="Save Changes", command=self._save, style="Primary.TButton").pack(side="right")
        ttk.Button(btns, text="Cancel", command=self._cancel, style="Quiet.TButton").pack(
            side="right", padx=(0, 8)
        )

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<MouseWheel>", self._on_mouse_wheel, add="+")
        self.bind("<Button-4>", self._on_mouse_wheel, add="+")
        self.bind("<Button-5>", self._on_mouse_wheel, add="+")

    def _resolve_target_canvas(self, widget) -> tk.Canvas | None:
        while widget is not None:
            if widget == self._output_canvas:
                return self._output_canvas
            if widget == self._source_canvas:
                return self._source_canvas
            widget = widget.master
        return None

    def _on_mouse_wheel(self, event):
        hovered_widget = self.winfo_containing(event.x_root, event.y_root)
        target_canvas = self._resolve_target_canvas(hovered_widget)
        if target_canvas is None:
            return
        if hasattr(event, "delta") and event.delta:
            step = int(-event.delta / 120)
        elif getattr(event, "num", None) == 4:
            step = -1
        elif getattr(event, "num", None) == 5:
            step = 1
        else:
            step = 0
        if step:
            target_canvas.yview_scroll(step, "units")
            return "break"

    def _save(self):
        source_values = {col: var.get().strip() for col, var in self.source_vars.items()}
        output_values = {col: var.get().strip() for col, var in self.output_vars.items()}
        try:
            out_qty_raw = output_values.get("Product/Service Quantity", "1")
            output_values["Product/Service Quantity"] = max(1, int(float(out_qty_raw or "1")))
            out_rate_raw = output_values.get("Product/Service Rate", "0")
            output_values["Product/Service Rate"] = round(float(out_rate_raw or "0"), 2)
        except ValueError:
            messagebox.showerror("Invalid row values", "Output Quantity/Rate must be numeric.")
            return
        self.result = {"source": source_values, "output": output_values}
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class SalesOrderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"QuickBooks Sales Order Converter {APP_VERSION}")
        self.root.geometry("1600x1040")
        try:
            self.root.minsize(1280, 800)
        except tk.TclError:
            pass
        self.root.minsize(1240, 820)
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self._configure_styles()
        self.settings = self._load_settings()

        self.source_df = None
        self.output_df = None
        self.preview_line_numbers: list[int] = []
        self.output_overrides: dict[str, dict] = {}
        self.source_preview_columns = ["Line #", "Brand", "SKU", "productName", "MSRP", "Qty", "Price", "Cost"]

        self.source_path_var = tk.StringVar(value=self.settings.get("source_path", DEFAULT_SOURCE))
        self.template_path_var = tk.StringVar(value=self.settings.get("template_path", DEFAULT_TEMPLATE))
        self.output_path_var = tk.StringVar(value=self.settings.get("output_path", DEFAULT_OUTPUT))
        self.customer_var = tk.StringVar(value=self.settings.get("customer", ""))
        self.sales_order_no_var = tk.StringVar(value=self.settings.get("sales_order_no", ""))
        self.sales_order_date_var = tk.StringVar(value=self._today_display_date())
        self.due_date_var = tk.StringVar(value=self._today_display_date())
        self.terms_var = tk.StringVar(value=self.settings.get("terms", "Prepaid"))
        self.shipping_method_var = tk.StringVar(value=self.settings.get("shipping_method", "Standard Ground"))
        self.memo_var = tk.StringVar(value=self.settings.get("memo", ""))
        self.currency_var = tk.StringVar(value=self.settings.get("currency", "USD"))
        self.tax_code_var = tk.StringVar(value=self.settings.get("tax_code", "TAX"))
        self.qb_company_file_var = tk.StringVar(value=self.settings.get("qb_company_file_path", ""))
        self.fallback_item_var = tk.StringVar(value=self.settings.get("fallback_item", ""))
        self.income_account_var = tk.StringVar(value=self.settings.get("income_account", ""))
        # Default ON: auto-connect on every launch. If the attach fails the
        # status pill just shows "Not Connected" and the user can click
        # Connect manually. The user can opt out by setting
        # "auto_connect_on_startup": false in settings.json.
        self.auto_connect_on_startup = bool(self.settings.get("auto_connect_on_startup", True))
        self.pricing_mode = self.settings.get("pricing_mode", "brand")
        self.use_actual_cost = bool(self.settings.get("use_actual_cost", False))
        self.default_pricing_value = float(self.settings.get("default_pricing_value", 0.4))
        self.brand_values: dict[str, float] = self.settings.get("brand_values", {})
        self.item_values: dict[str, float] = self.settings.get("item_values", {})
        self.line_values: dict[str, float] = self.settings.get("line_values", {})
        self.status_var = tk.StringVar(value="Ready. Load source file, then build preview.")
        self.qb_status_var = tk.StringVar(value="Not Connected")
        self.qb_status_label: ttk.Label | None = None

        self._build_layout()
        self._build_menu()
        self._set_qb_status("Not Connected", state="disconnected")
        self.root.after(400, self._warn_if_elevated)
        if self.auto_connect_on_startup:
            self.root.after(900, self._connect_quickbooks_on_startup)
        else:
            log.info("Auto-connect on startup disabled")
        self.root.after(1200, self.check_for_updates_on_startup)

    def _warn_if_elevated(self) -> None:
        """Show a one-time warning if the app is running elevated.

        QuickBooks Desktop is almost always launched as a standard user. If
        this app is elevated, QBXMLRP2 cannot attach across the UAC boundary
        and instead spawns a second elevated QuickBooks window, which fails
        with -2147220457. The user has to relaunch this app non-elevated to
        recover — this dialog tells them up front instead of making them
        diagnose it from the log.
        """
        try:
            elevation = _current_process_elevation()
        except Exception:
            log.exception("_warn_if_elevated: could not determine elevation")
            return
        log.info("Startup elevation check: %s", elevation)
        if elevation != "elevated":
            return
        log.warning("Startup: app is elevated — showing relaunch warning to user")
        self._set_qb_status("Elevated — relaunch as non-admin", state="disconnected")
        messagebox.showwarning(
            "Run as Administrator detected",
            "This app is running as Administrator.\n\n"
            "QuickBooks Desktop is typically running under your normal Windows user, "
            "and the SDK cannot attach across that UAC boundary — connection will fail "
            "and a second QuickBooks window will appear.\n\n"
            "Close this app, then reopen it normally (not 'Run as administrator'):\n"
            "  • Right-click the shortcut → Properties → Shortcut → Advanced…\n"
            "  • Make sure 'Run as administrator' is UNCHECKED → OK.\n\n"
            "Then relaunch the app and click Connect.",
        )

    def _configure_styles(self) -> None:
        c = UI
        self.root.configure(bg=c["bg_window"])

        self.style.configure(
            ".",
            font=("Segoe UI", 10),
            background=c["bg_window"],
            foreground=c["text_primary"],
        )

        self.style.configure("TFrame", background=c["bg_window"])
        self.style.configure("Card.TFrame", background=c["bg_card"])

        self.style.configure(
            "TLabel",
            background=c["bg_window"],
            foreground=c["text_primary"],
            font=("Segoe UI", 10),
        )
        self.style.configure(
            "Card.TLabel",
            background=c["bg_card"],
            foreground=c["text_primary"],
            font=("Segoe UI", 10),
        )
        self.style.configure(
            "Header.TLabel",
            background=c["bg_window"],
            foreground=c["accent"],
            font=("Segoe UI Semibold", 26),
        )
        self.style.configure(
            "AppTitle.TLabel",
            background=c["bg_window"],
            foreground=c["accent"],
            font=("Segoe UI Semibold", 28),
        )
        self.style.configure(
            "SubHeader.TLabel",
            background=c["bg_window"],
            foreground=c["text_secondary"],
            font=("Segoe UI", 11),
        )
        self.style.configure(
            "CardSubHeader.TLabel",
            background=c["bg_card"],
            foreground=c["text_secondary"],
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "FieldLabel.TLabel",
            background=c["bg_card"],
            foreground=c["text_secondary"],
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "Status.TLabel",
            background=c["bg_window"],
            foreground=c["text_secondary"],
            padding=(14, 4),
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "QbConnected.TLabel",
            background=c["bg_card"],
            foreground=c["success"],
            font=("Segoe UI Semibold", 9),
        )
        self.style.configure(
            "QbDisconnected.TLabel",
            background=c["bg_card"],
            foreground=c["danger"],
            font=("Segoe UI Semibold", 9),
        )
        self.style.configure(
            "QbPending.TLabel",
            background=c["bg_card"],
            foreground=c["warning"],
            font=("Segoe UI Semibold", 9),
        )

        self.style.configure(
            "Card.TLabelframe",
            background=c["bg_card"],
            bordercolor=c["border"],
            lightcolor=c["border_inner_light"],
            darkcolor=c["border"],
            borderwidth=1,
            relief="solid",
            padding=(16, 10, 16, 12),
        )
        self.style.configure(
            "Card.TLabelframe.Label",
            background=c["bg_card"],
            foreground=c["text_primary"],
            font=("Segoe UI Semibold", 13),
            padding=(10, 6),
        )

        self.style.configure(
            "TEntry",
            fieldbackground=c["bg_card"],
            foreground=c["text_primary"],
            bordercolor=c["border_strong"],
            lightcolor=c["border_strong"],
            darkcolor=c["border"],
            insertcolor=c["text_primary"],
            borderwidth=1,
            padding=(10, 8),
            relief="solid",
        )
        self.style.map(
            "TEntry",
            bordercolor=[("focus", c["accent"]), ("hover", c["accent_light"])],
            lightcolor=[("focus", c["accent"]), ("hover", c["accent_light"])],
            darkcolor=[("focus", c["accent"]), ("hover", c["accent_light"])],
            fieldbackground=[("disabled", c["bg_hover"])],
            foreground=[("disabled", c["text_tertiary"])],
        )

        self.style.configure(
            "TButton",
            padding=(14, 9),
            font=("Segoe UI", 10),
            background=c["bg_card"],
            foreground=c["text_primary"],
            bordercolor=c["btn_light_border"],
            lightcolor=c["btn_light_top"],
            darkcolor=c["btn_light_bottom"],
            borderwidth=1,
            focusthickness=0,
            relief="flat",
        )
        self.style.map(
            "TButton",
            background=[
                ("pressed", c["bg_pressed"]),
                ("active", c["bg_hover"]),
                ("disabled", c["bg_hover"]),
            ],
            foreground=[("disabled", c["text_tertiary"])],
            relief=[("pressed", "flat")],
            bordercolor=[("active", c["accent_light"]), ("pressed", c["accent"])],
        )

        self.style.configure(
            "Primary.TButton",
            padding=(20, 10),
            font=("Segoe UI Semibold", 10),
            background=c["accent"],
            foreground="#FFFFFF",
            bordercolor=c["accent_dark"],
            lightcolor=c["accent_light"],
            darkcolor=c["accent_dark"],
            borderwidth=1,
            focusthickness=0,
            relief="flat",
        )
        self.style.map(
            "Primary.TButton",
            background=[
                ("pressed", c["accent_pressed"]),
                ("active", c["accent_hover"]),
                ("disabled", c["border_strong"]),
            ],
            foreground=[("disabled", c["bg_card"])],
            lightcolor=[
                ("pressed", c["accent_dark"]),
                ("active", "#52A5EF"),
            ],
            darkcolor=[
                ("pressed", c["accent_light"]),
                ("active", c["accent_pressed"]),
            ],
            relief=[("pressed", "flat")],
            bordercolor=[
                ("pressed", c["accent_dark"]),
                ("active", c["accent_pressed"]),
            ],
        )

        self.style.configure(
            "Accent.TButton",
            padding=(16, 8),
            font=("Segoe UI Semibold", 10),
            background=c["bg_card"],
            foreground=c["accent"],
            bordercolor=c["accent"],
            lightcolor=c["btn_light_top"],
            darkcolor=c["accent_bg"],
            borderwidth=1,
            focusthickness=0,
            relief="flat",
        )
        self.style.map(
            "Accent.TButton",
            background=[
                ("pressed", c["accent_bg"]),
                ("active", c["bg_hover"]),
                ("disabled", c["bg_hover"]),
            ],
            foreground=[
                ("pressed", c["accent_pressed"]),
                ("active", c["accent_hover"]),
                ("disabled", c["text_tertiary"]),
            ],
            relief=[("pressed", "flat")],
            bordercolor=[
                ("pressed", c["accent_dark"]),
                ("active", c["accent_light"]),
                ("disabled", c["border"]),
            ],
        )

        self.style.configure(
            "Quiet.TButton",
            padding=(14, 8),
            font=("Segoe UI", 10),
            background=c["bg_card"],
            foreground=c["text_primary"],
            bordercolor=c["btn_light_border"],
            lightcolor=c["btn_light_top"],
            darkcolor=c["btn_light_bottom"],
            borderwidth=1,
            focusthickness=0,
            relief="flat",
        )
        self.style.map(
            "Quiet.TButton",
            background=[
                ("pressed", c["bg_pressed"]),
                ("active", c["bg_hover"]),
            ],
            foreground=[("disabled", c["text_tertiary"])],
            relief=[("pressed", "flat")],
            bordercolor=[("active", c["accent_light"]), ("pressed", c["accent"])],
        )

        self.style.configure(
            "Treeview",
            background=c["bg_card"],
            fieldbackground=c["bg_card"],
            foreground=c["text_primary"],
            rowheight=28,
            borderwidth=1,
            relief="solid",
            bordercolor=c["border"],
            font=("Segoe UI", 10),
        )
        self.style.configure(
            "Treeview.Heading",
            background=c["bg_subtle"],
            foreground=c["text_primary"],
            font=("Segoe UI Semibold", 9),
            padding=(12, 10),
            borderwidth=1,
            lightcolor=c["bg_card"],
            darkcolor=c["border"],
            bordercolor=c["border"],
            relief="flat",
        )
        self.style.map(
            "Treeview.Heading",
            background=[("active", c["bg_hover"])],
            foreground=[("active", c["text_primary"])],
        )
        self.style.map(
            "Treeview",
            background=[("selected", c["accent_bg"])],
            foreground=[("selected", c["text_primary"])],
        )

        self.style.configure(
            "TNotebook",
            background=c["bg_card"],
            borderwidth=0,
            tabmargins=(2, 4, 2, 0),
        )
        self.style.configure(
            "TNotebook.Tab",
            padding=(16, 8),
            background=c["bg_subtle"],
            foreground=c["text_secondary"],
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", c["bg_card"]), ("active", c["bg_hover"])],
            foreground=[("selected", c["accent"]), ("active", c["text_primary"])],
            font=[("selected", ("Segoe UI Semibold", 10))],
        )

        for orient in ("Vertical", "Horizontal"):
            self.style.configure(
                f"{orient}.TScrollbar",
                background=c["bg_card"],
                troughcolor=c["bg_card"],
                bordercolor=c["bg_card"],
                arrowcolor=c["text_tertiary"],
                gripcount=0,
                borderwidth=0,
                relief="flat",
                arrowsize=14,
            )
            self.style.map(
                f"{orient}.TScrollbar",
                background=[("active", c["border_strong"]), ("pressed", c["text_tertiary"])],
                arrowcolor=[("active", c["text_primary"])],
            )

        self.style.configure("TPanedwindow", background=c["bg_window"])
        self.style.configure(
            "Sash",
            sashthickness=8,
            gripcount=8,
            background=c["bg_window"],
            bordercolor=c["border"],
            lightcolor=c["bg_card"],
            darkcolor=c["border_strong"],
            handlesize=20,
        )

        self.style.configure(
            "TProgressbar",
            background=c["accent"],
            troughcolor=c["bg_hover"],
            bordercolor=c["bg_hover"],
            lightcolor=c["accent"],
            darkcolor=c["accent"],
            borderwidth=0,
            thickness=8,
        )

        self.style.configure(
            "TRadiobutton",
            background=c["bg_window"],
            foreground=c["text_primary"],
            font=("Segoe UI", 10),
            focuscolor=c["bg_window"],
        )
        self.style.map(
            "TRadiobutton",
            background=[("active", c["bg_window"])],
            foreground=[("active", c["text_primary"])],
        )
        self.style.configure(
            "TCheckbutton",
            background=c["bg_window"],
            foreground=c["text_primary"],
            font=("Segoe UI", 10),
            focuscolor=c["bg_window"],
        )
        self.style.map(
            "TCheckbutton",
            background=[("active", c["bg_window"])],
            foreground=[("active", c["text_primary"])],
        )

    def _build_menu(self) -> None:
        c = UI
        menu_bar = tk.Menu(self.root, bg=c["bg_card"], fg=c["text_primary"], borderwidth=0)
        help_menu = tk.Menu(
            menu_bar,
            tearoff=0,
            bg=c["bg_card"],
            fg=c["text_primary"],
            activebackground=c["accent"],
            activeforeground="#FFFFFF",
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 10),
        )
        help_menu.add_command(label="Check for Updates", command=self.check_for_updates)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", f"{APP_NAME} {APP_VERSION}"))
        menu_bar.add_cascade(label="Help", menu=help_menu)

        log_menu = tk.Menu(
            menu_bar,
            tearoff=0,
            bg=c["bg_card"],
            fg=c["text_primary"],
            activebackground=c["accent"],
            activeforeground="#FFFFFF",
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 10),
        )
        log_menu.add_command(label="Open Log File", command=self.open_log_file)
        log_menu.add_command(label="Show Log Folder", command=self.reveal_log_folder)
        log_menu.add_separator()
        log_menu.add_command(label="Clear Log", command=self.clear_log_file)
        menu_bar.add_cascade(label="Log", menu=log_menu)

        self.root.config(menu=menu_bar)

    def open_log_file(self) -> None:
        log.info("User opened log file via menu")
        try:
            LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            LOG_PATH.touch(exist_ok=True)
            # Use notepad explicitly. os.startfile would honor the user's .log
            # file association, which on some machines is wired to QuickBooks's
            # log viewer (or other tools) — we just want a plain text reader.
            subprocess.Popen(["notepad.exe", str(LOG_PATH)])
        except Exception as exc:
            log.exception("Failed to open log file")
            messagebox.showerror(
                "Log",
                f"Could not open the log file at:\n{LOG_PATH}\n\n{exc}",
            )

    def reveal_log_folder(self) -> None:
        log.info("User opened log folder via menu")
        try:
            LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            subprocess.Popen(["explorer", str(LOG_PATH.parent)])
        except Exception as exc:
            log.exception("Failed to open log folder")
            messagebox.showerror("Log", f"Could not open folder:\n{LOG_PATH.parent}\n\n{exc}")

    def clear_log_file(self) -> None:
        if not messagebox.askyesno("Clear Log", "Erase the log file? This cannot be undone."):
            return
        try:
            LOG_PATH.write_text("", encoding="utf-8")
            log.info("Log file cleared by user")
            messagebox.showinfo("Log", "Log file cleared.")
        except Exception as exc:
            log.exception("Failed to clear log file")
            messagebox.showerror("Log", f"Could not clear log:\n{exc}")

    def _version_tuple(self, version_str: str) -> tuple[int, int, int]:
        cleaned = version_str.lower().replace("v", "").replace("beta", "").strip()
        cleaned = cleaned.split("-")[0]
        parts = cleaned.split(".")
        nums = []
        for part in parts[:3]:
            num = "".join(ch for ch in part if ch.isdigit())
            nums.append(int(num) if num else 0)
        while len(nums) < 3:
            nums.append(0)
        return tuple(nums)

    def _today_display_date(self) -> str:
        return datetime.now().strftime("%m/%d/%Y")

    def _normalize_date_for_qb(self, date_text: str) -> str:
        date_text = (date_text or "").strip()
        if not date_text:
            return ""
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(date_text, fmt).strftime("%m/%d/%Y")
            except ValueError:
                continue
        raise ValueError(f"Invalid date format: {date_text}. Use MM/DD/YYYY.")

    def _normalize_date_for_display(self, date_text: str) -> str:
        date_text = (date_text or "").strip()
        if not date_text:
            return ""
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(date_text, fmt).strftime("%m/%d/%Y")
            except ValueError:
                continue
        raise ValueError(f"Invalid date format: {date_text}. Use MM/DD/YYYY.")

    def _fetch_release_info(self) -> dict:
        """Fetch latest release info from GitHub Releases API (uncached); fall back to raw JSON feed."""
        try:
            api_request = urllib.request.Request(
                UPDATE_API_URL,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": f"{APP_NAME}/{APP_VERSION}",
                },
            )
            with urllib.request.urlopen(api_request, timeout=10) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            tag = str(payload.get("tag_name", "")).strip()
            version = tag.lstrip("vV")
            body = str(payload.get("body", "")).strip()
            download_url = ""
            for asset in payload.get("assets", []) or []:
                name = str(asset.get("name", ""))
                if name.lower().endswith("-setup.exe"):
                    download_url = str(asset.get("browser_download_url", ""))
                    break
            sha_match = re.search(r"SHA256[:\s]+([0-9a-fA-F]{64})", body)
            sha256_hash = sha_match.group(1).lower() if sha_match else ""
            notes = re.sub(r"\s*SHA256[:\s]+[0-9a-fA-F]{64}\s*", "\n", body).strip()
            if version and download_url:
                return {"version": version, "url": download_url, "sha256": sha256_hash, "notes": notes}
        except Exception:
            pass

        cache_bust_url = f"{UPDATE_INFO_URL}?t={int(datetime.now().timestamp())}"
        request = urllib.request.Request(
            cache_bust_url,
            headers={
                "Cache-Control": "no-cache, no-store, max-age=0",
                "Pragma": "no-cache",
                "User-Agent": f"{APP_NAME}/{APP_VERSION}",
            },
        )
        with urllib.request.urlopen(request, timeout=10) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
        return {
            "version": str(payload.get("version", "")).strip(),
            "url": str(payload.get("url", "")).strip(),
            "sha256": str(payload.get("sha256", "")).strip().lower(),
            "notes": str(payload.get("notes", "")).strip(),
        }

    def check_for_updates(self, silent: bool = False) -> None:
        try:
            info = self._fetch_release_info()
            latest_version = info["version"]
            download_url = info["url"]
            sha256_hash = info["sha256"]
            notes = info["notes"]

            if not latest_version or not download_url:
                raise RuntimeError("Update feed is missing version or url.")

            current_version = self._version_tuple(APP_VERSION)
            available_version = self._version_tuple(latest_version)
            if available_version <= current_version:
                if not silent:
                    messagebox.showinfo("No Updates", f"You're up to date on {APP_VERSION}.")
                self._set_status(f"Update check complete: {APP_VERSION} is current.")
                return

            if not self._show_update_dialog(latest_version, notes):
                return

            self._download_and_run_update(download_url, sha256_hash)
        except Exception as exc:
            if not silent:
                messagebox.showerror("Update Check Failed", str(exc))
            self._set_status("Update check failed. Use 'Check for Updates' to retry.")

    def check_for_updates_on_startup(self) -> None:
        self._set_status("Checking for updates...")
        self.check_for_updates(silent=True)

    def _clean_release_notes(self, notes: str) -> str:
        """Strip technical noise (markdown markers, SHA256 line, headers) so
        the notes render as plain prose in the update dialog."""
        if not notes:
            return ""
        text = notes
        # Remove SHA256 lines entirely — verification is automated, the user
        # doesn't need to see the hash.
        text = re.sub(r"SHA-?256[^\n]*", "", text, flags=re.IGNORECASE)
        # Strip code spans, leaving the content.
        text = re.sub(r"`([^`]+)`", r"\1", text)
        # Strip bold markers.
        text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
        # Strip leading markdown headers (## Heading -> Heading).
        text = re.sub(r"^\s*#{1,6}\s+", "", text, flags=re.MULTILINE)
        # Collapse stray blank lines.
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    def _show_update_dialog(self, latest_version: str, notes: str) -> bool:
        """Present a clean Update Available dialog. Returns True if user wants
        to install, False otherwise."""
        c = UI
        dlg = tk.Toplevel(self.root)
        dlg.title("Update Available")
        dlg.configure(bg=c["bg_window"])
        dlg.geometry("520x440")
        try:
            dlg.minsize(440, 360)
        except tk.TclError:
            pass
        dlg.transient(self.root)
        dlg.grab_set()

        decision = {"install": False}

        accent_strip = tk.Frame(dlg, bg=c["accent"], height=3)
        accent_strip.pack(fill="x", side="top")

        body = tk.Frame(dlg, bg=c["bg_window"])
        body.pack(fill="both", expand=True, padx=22, pady=(18, 14))

        tk.Label(
            body,
            text=f"Version {latest_version} is available",
            bg=c["bg_window"],
            fg=c["accent"],
            font=("Segoe UI Semibold", 16),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            body,
            text=f"You're on {APP_VERSION}.",
            bg=c["bg_window"],
            fg=c["text_secondary"],
            font=("Segoe UI", 10),
            anchor="w",
        ).pack(fill="x", pady=(2, 12))

        cleaned = self._clean_release_notes(notes)
        if cleaned:
            tk.Label(
                body,
                text="What's new",
                bg=c["bg_window"],
                fg=c["text_primary"],
                font=("Segoe UI Semibold", 10),
                anchor="w",
            ).pack(fill="x")
            text_wrap = tk.Frame(body, bg=c["bg_card"], highlightbackground=c["border"], highlightthickness=1, bd=0)
            text_wrap.pack(fill="both", expand=True, pady=(4, 12))
            notes_text = tk.Text(
                text_wrap,
                wrap="word",
                bg=c["bg_card"],
                fg=c["text_primary"],
                font=("Segoe UI", 10),
                bd=0,
                relief="flat",
                padx=12,
                pady=10,
                height=10,
            )
            sb = ttk.Scrollbar(text_wrap, orient="vertical", command=notes_text.yview)
            notes_text.configure(yscrollcommand=sb.set)
            notes_text.insert("1.0", cleaned)
            notes_text.configure(state="disabled")
            notes_text.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

        button_row = tk.Frame(body, bg=c["bg_window"])
        button_row.pack(fill="x", side="bottom")

        def install():
            decision["install"] = True
            dlg.destroy()

        def cancel():
            decision["install"] = False
            dlg.destroy()

        ttk.Button(button_row, text="Not Now", command=cancel, style="Quiet.TButton").pack(side="right")
        ttk.Button(button_row, text="Install Update", command=install, style="Primary.TButton").pack(
            side="right", padx=(0, 10)
        )

        dlg.protocol("WM_DELETE_WINDOW", cancel)
        dlg.update_idletasks()
        # Center over the main window.
        try:
            self.root.update_idletasks()
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            dw = dlg.winfo_width()
            dh = dlg.winfo_height()
            dlg.geometry(f"+{rx + (rw - dw) // 2}+{ry + (rh - dh) // 2}")
        except tk.TclError:
            pass

        dlg.wait_window()
        return decision["install"]

    def _download_and_run_update(self, url: str, expected_sha256: str) -> None:
        temp_dir = Path(tempfile.gettempdir()) / "qb_so_updates"
        temp_dir.mkdir(parents=True, exist_ok=True)
        filename = Path(urllib.parse.urlparse(url).path).name or "QB-Sales-Order-Converter-Setup.exe"
        installer_path = temp_dir / filename
        progress_dialog = tk.Toplevel(self.root)
        progress_dialog.title("Installing Update")
        progress_dialog.geometry("460x170")
        progress_dialog.resizable(False, False)
        progress_dialog.configure(bg=UI["bg_window"])
        progress_dialog.transient(self.root)
        progress_dialog.grab_set()
        progress_dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        status_var = tk.StringVar(value="Downloading update...")
        progress_var = tk.DoubleVar(value=0.0)
        ttk.Label(progress_dialog, text="Updating app", style="Header.TLabel").pack(
            anchor="w", padx=20, pady=(18, 2)
        )
        ttk.Label(progress_dialog, textvariable=status_var, wraplength=420, style="SubHeader.TLabel").pack(
            anchor="w", padx=20, pady=(0, 10)
        )
        progress_bar = ttk.Progressbar(progress_dialog, mode="determinate", maximum=100, variable=progress_var)
        progress_bar.pack(fill="x", padx=20, pady=(0, 6))
        pct_label = ttk.Label(progress_dialog, text="0%", style="SubHeader.TLabel")
        pct_label.pack(anchor="e", padx=20)

        def update_ui(status: str | None = None, pct: float | None = None) -> None:
            def _apply():
                if status is not None:
                    status_var.set(status)
                if pct is not None:
                    bounded = max(0.0, min(100.0, pct))
                    progress_var.set(bounded)
                    pct_label.config(text=f"{int(round(bounded))}%")
            self.root.after(0, _apply)

        def fail_ui(message: str) -> None:
            def _apply():
                if progress_dialog.winfo_exists():
                    progress_dialog.destroy()
                messagebox.showerror("Update Failed", message)
                self._set_status("Update failed. Please try again.")
            self.root.after(0, _apply)

        def worker() -> None:
            try:
                update_ui("Downloading update...", 2)
                with urllib.request.urlopen(url, timeout=30) as response, installer_path.open("wb") as out_file:
                    total_bytes = int(response.headers.get("Content-Length", "0") or "0")
                    downloaded = 0
                    while True:
                        chunk = response.read(1024 * 256)
                        if not chunk:
                            break
                        out_file.write(chunk)
                        downloaded += len(chunk)
                        if total_bytes > 0:
                            pct = 2 + (downloaded / total_bytes) * 78
                            update_ui(pct=pct)

                update_ui("Verifying update package...", 84)
                if expected_sha256:
                    digest = hashlib.sha256(installer_path.read_bytes()).hexdigest().lower()
                    if digest != expected_sha256:
                        raise RuntimeError("Downloaded update failed checksum validation.")

                update_ui("Installing update... installer progress will stay visible.", 100)
                subprocess.Popen(
                    [str(installer_path), "/SILENT", "/NORESTART", "/CLOSEAPPLICATIONS", "/FORCECLOSEAPPLICATIONS"]
                )
                self.root.after(500, self.root.destroy)
            except Exception as exc:
                fail_ui(str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _load_settings(self) -> dict:
        if not SETTINGS_PATH.exists():
            return {}
        try:
            return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _persist_settings(self) -> None:
        payload = {
            "source_path": self.source_path_var.get().strip(),
            "template_path": self.template_path_var.get().strip(),
            "output_path": self.output_path_var.get().strip(),
            "customer": self.customer_var.get().strip(),
            "sales_order_no": self.sales_order_no_var.get().strip(),
            "sales_order_date": self.sales_order_date_var.get().strip(),
            "due_date": self.due_date_var.get().strip(),
            "terms": self.terms_var.get().strip(),
            "shipping_method": self.shipping_method_var.get().strip(),
            "memo": self.memo_var.get().strip(),
            "currency": self.currency_var.get().strip(),
            "tax_code": self.tax_code_var.get().strip(),
            "qb_company_file_path": self.qb_company_file_var.get().strip(),
            "fallback_item": self.fallback_item_var.get().strip(),
            "income_account": self.income_account_var.get().strip(),
            "auto_connect_on_startup": self.auto_connect_on_startup,
            "pricing_mode": self.pricing_mode,
            "use_actual_cost": self.use_actual_cost,
            "default_pricing_value": self.default_pricing_value,
            "brand_values": self.brand_values,
            "item_values": self.item_values,
            "line_values": self.line_values,
        }
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _build_layout(self) -> None:
        root = self.root
        c = UI

        accent_strip = tk.Frame(root, bg=c["accent"], height=4)
        accent_strip.pack(fill="x", side="top")

        header = ttk.Frame(root, padding=(32, 22, 32, 16))
        header.pack(fill="x")
        title_row = ttk.Frame(header)
        title_row.pack(fill="x")
        ttk.Label(title_row, text="QB Sales Order Converter", style="AppTitle.TLabel").pack(side="left")
        ttk.Label(
            header,
            text=f"{APP_VERSION} · Upload source data, review mapped rows, then export or upload to QuickBooks Desktop.",
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(6, 0))

        config = ttk.LabelFrame(root, text="  Configuration  ", style="Card.TLabelframe")
        config.pack(fill="x", padx=32, pady=(4, 10))

        self._path_row(config, "Source File", self.source_path_var, self._browse_source, 0)
        self._path_row(config, "SaaSant Template", self.template_path_var, self._browse_template, 1)
        self._path_row(config, "Output File", self.output_path_var, self._browse_output, 2)
        self._path_row(config, "QB Company File (.QBW)", self.qb_company_file_var, self._browse_qb_company_file, 3)

        form = ttk.Frame(config, style="Card.TFrame")
        form.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        for i in range(6):
            form.columnconfigure(i, weight=1)

        self._form_entry(form, "Customer", self.customer_var, 0, 0, 2)
        self._form_entry(form, "Sales Order No", self.sales_order_no_var, 0, 2, 1)
        ttk.Button(form, text="Fetch Next from QuickBooks", command=self.fetch_next_so, style="Quiet.TButton").grid(
            row=1, column=3, padx=4, sticky="ew"
        )
        self._form_entry(form, "Sales Order Date", self.sales_order_date_var, 0, 4, 1)
        self._form_entry(form, "Due Date", self.due_date_var, 0, 5, 1)
        ttk.Label(
            form,
            text="Date format: MM/DD/YYYY (example: 05/12/2026)",
            style="CardSubHeader.TLabel",
        ).grid(row=2, column=4, columnspan=2, sticky="w", padx=4, pady=(0, 2))

        self._form_entry(form, "Terms", self.terms_var, 3, 0, 1)
        self._form_entry(form, "Shipping Method", self.shipping_method_var, 3, 1, 1)
        self._form_entry(form, "Memo", self.memo_var, 3, 2, 2)
        self._form_entry(form, "Currency", self.currency_var, 3, 4, 1)
        self._form_entry(form, "Tax Code", self.tax_code_var, 3, 5, 1)

        self._form_entry(
            form,
            "Default Income Account (auto-creates missing SKUs under this account)",
            self.income_account_var,
            5,
            0,
            3,
        )
        self._form_entry(
            form,
            "Fallback Item (used only if Default Income Account is empty)",
            self.fallback_item_var,
            5,
            3,
            3,
        )

        qb_bar = ttk.Frame(config, style="Card.TFrame")
        qb_bar.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Button(
            qb_bar,
            text="Connect to QuickBooks Desktop",
            command=self.connect_quickbooks,
            style="Accent.TButton",
        ).pack(side="left")
        ttk.Button(
            qb_bar,
            text="QuickBooks Admin Setup",
            command=self.show_qb_admin_setup_guide,
            style="Quiet.TButton",
        ).pack(side="left", padx=8)
        ttk.Button(
            qb_bar,
            text="Check for Updates",
            command=self.check_for_updates,
            style="Quiet.TButton",
        ).pack(side="left")
        self.qb_status_pill = tk.Frame(
            qb_bar,
            bg=c["bg_hover"],
            highlightbackground=c["bg_hover"],
            highlightthickness=0,
            bd=0,
        )
        self.qb_status_pill.pack(side="left", padx=12)
        self.qb_status_inner = tk.Label(
            self.qb_status_pill,
            textvariable=self.qb_status_var,
            bg=c["bg_hover"],
            fg=c["text_secondary"],
            font=("Segoe UI Semibold", 9),
            padx=12,
            pady=4,
        )
        self.qb_status_inner.pack()
        self.qb_status_label = self.qb_status_inner

        actions = ttk.Frame(root, padding=(32, 4, 32, 10))
        actions.pack(fill="x")
        ttk.Button(actions, text="Build Preview", command=self.build_preview, style="Primary.TButton").pack(side="left")
        ttk.Button(
            actions,
            text="Change Pricing Rules",
            command=self.change_pricing_rules,
            style="Quiet.TButton",
        ).pack(side="left", padx=8)
        ttk.Button(
            actions,
            text="Export for SaaSant",
            command=self.export_saasant_template,
            style="Accent.TButton",
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            actions,
            text="Export Template (Custom Path)",
            command=self.export_file,
            style="Quiet.TButton",
        ).pack(side="left", padx=8)
        ttk.Button(
            actions,
            text="Upload to QuickBooks",
            command=self.upload_to_quickbooks,
            style="Accent.TButton",
        ).pack(side="left")

        content = ttk.Panedwindow(root, orient="vertical")
        content.pack(fill="both", expand=True, padx=32, pady=(0, 8))

        preview_frame = ttk.LabelFrame(content, text="  Preview  ", style="Card.TLabelframe")
        content.add(preview_frame, weight=12)

        self.preview_notebook = ttk.Notebook(preview_frame)
        self.preview_notebook.pack(fill="both", expand=True)
        output_tab = ttk.Frame(self.preview_notebook, style="Card.TFrame")
        source_tab = ttk.Frame(self.preview_notebook, style="Card.TFrame")
        self.preview_notebook.add(output_tab, text="Output Preview")
        self.preview_notebook.add(source_tab, text="Source Preview")

        self.preview_columns = [
            "Line #",
            "Sales Order No",
            "Customer",
            "Sales Order Date",
            "Due Date",
            "Terms",
            "Shipping Method",
            "Memo",
            "Product/Service",
            "Product/Service Description",
            "Product/Service Quantity",
            "Product/Service Rate",
            "Product/Service Sales Tax Code",
            "Currency",
        ]
        self.output_tree = ttk.Treeview(output_tab, columns=self.preview_columns, show="headings")
        for col in self.preview_columns:
            self.output_tree.heading(col, text=col)
            if col == "Line #":
                width = 70
            else:
                width = 420 if "Description" in col else 165
            self.output_tree.column(col, width=width, anchor="w")
        self.output_tree.tag_configure("odd", background=c["bg_subtle"])
        self.output_tree.tag_configure("even", background=c["bg_card"])

        out_scroll_y = ttk.Scrollbar(output_tab, orient="vertical", command=self.output_tree.yview)
        out_scroll_x = ttk.Scrollbar(output_tab, orient="horizontal", command=self.output_tree.xview)
        self.output_tree.configure(yscrollcommand=out_scroll_y.set, xscrollcommand=out_scroll_x.set)
        output_tab.columnconfigure(0, weight=1)
        output_tab.rowconfigure(0, weight=1)
        self.output_tree.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=(4, 0))
        out_scroll_y.grid(row=0, column=1, sticky="ns")
        out_scroll_x.grid(row=1, column=0, sticky="ew")
        self.output_tree.bind("<Double-1>", self._edit_output_row)

        self.source_tree = ttk.Treeview(source_tab, columns=self.source_preview_columns, show="headings")
        for col in self.source_preview_columns:
            self.source_tree.heading(col, text=col)
            width = 220 if col == "productName" else (80 if col == "Line #" else 120)
            self.source_tree.column(col, width=width, anchor="w")
        self.source_tree.tag_configure("odd", background=c["bg_subtle"])
        self.source_tree.tag_configure("even", background=c["bg_card"])
        src_scroll_y = ttk.Scrollbar(source_tab, orient="vertical", command=self.source_tree.yview)
        src_scroll_x = ttk.Scrollbar(source_tab, orient="horizontal", command=self.source_tree.xview)
        self.source_tree.configure(yscrollcommand=src_scroll_y.set, xscrollcommand=src_scroll_x.set)
        source_tab.columnconfigure(0, weight=1)
        source_tab.rowconfigure(0, weight=1)
        self.source_tree.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=(4, 0))
        src_scroll_y.grid(row=0, column=1, sticky="ns")
        src_scroll_x.grid(row=1, column=0, sticky="ew")
        self.source_tree.bind("<Double-1>", self._edit_source_row)

        error_frame = ttk.LabelFrame(content, text="  Validation Messages  ", style="Card.TLabelframe")
        content.add(error_frame, weight=1)
        self.error_text = tk.Text(
            error_frame,
            height=3,
            wrap="word",
            bg=c["bg_card"],
            fg=c["text_primary"],
            insertbackground=c["text_primary"],
            selectbackground=c["accent_bg"],
            selectforeground=c["text_primary"],
            borderwidth=0,
            highlightthickness=0,
            relief="flat",
            font=("Segoe UI", 10),
            padx=4,
            pady=4,
        )
        self.error_text.pack(fill="both", expand=True)

        status_bar = tk.Frame(root, bg=c["bg_window"], highlightthickness=0, bd=0)
        status_bar.pack(fill="x", side="bottom")
        separator = tk.Frame(status_bar, bg=c["border"], height=1)
        separator.pack(fill="x", side="top")
        ttk.Label(status_bar, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w")

    def _path_row(self, parent, label, var, browse_cmd, row_idx):
        ttk.Label(parent, text=label, style="Card.TLabel", width=18).grid(
            row=row_idx, column=0, sticky="w", pady=(2, 2)
        )
        ttk.Entry(parent, textvariable=var).grid(row=row_idx, column=1, sticky="ew", padx=10, pady=(2, 2))
        ttk.Button(parent, text="Browse", command=browse_cmd).grid(
            row=row_idx, column=2, sticky="e", pady=(2, 2)
        )
        parent.columnconfigure(1, weight=1)

    def _form_entry(self, parent, label, var, row, col, span):
        ttk.Label(parent, text=label, style="FieldLabel.TLabel").grid(
            row=row, column=col, sticky="w", padx=4, pady=(4, 1)
        )
        ttk.Entry(parent, textvariable=var).grid(
            row=row + 1, column=col, columnspan=span, sticky="ew", padx=4, pady=(0, 4)
        )

    def _tree_values_from_row(self, row, row_index: int | None = None) -> tuple:
        values = []
        for col in self.preview_columns:
            if col == "Line #":
                if row_index is not None and 0 <= row_index < len(self.preview_line_numbers):
                    values.append(self.preview_line_numbers[row_index])
                else:
                    values.append("")
            else:
                values.append(row.get(col, ""))
        return tuple(values)

    def _source_tree_values_from_row(self, source_row_index: int) -> tuple:
        row = self.source_df.loc[source_row_index]
        line_no = int(source_row_index) + 2
        return (
            line_no,
            row.get("Brand", ""),
            row.get("SKU", ""),
            row.get("productName", ""),
            row.get("MSRP", ""),
            row.get("Qty", ""),
            row.get("Price", ""),
            row.get("Cost", ""),
        )

    def _apply_output_overrides(self):
        if self.output_df is None or self.output_df.empty:
            return
        for row_idx, source_row_index in enumerate(self.source_df.index.tolist()):
            override = self.output_overrides.get(str(source_row_index))
            if not override:
                continue
            for key, value in override.items():
                if key in self.output_df.columns:
                    if key in ("Sales Order Date", "Due Date") and str(value).strip():
                        value = self._normalize_date_for_display(str(value))
                    self.output_df.at[row_idx, key] = value

    def _browse_source(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls *.xlsx")])
        if path:
            self.source_path_var.set(path)
            self._persist_settings()
            self._set_status(f"Selected source file: {Path(path).name}")

    def _browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.template_path_var.set(path)
            self._persist_settings()
            self._set_status(f"Selected template file: {Path(path).name}")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.output_path_var.set(path)
            self._persist_settings()
            self._set_status(f"Output path updated: {Path(path).name}")

    def _browse_qb_company_file(self):
        path = filedialog.askopenfilename(
            title="Select QuickBooks Company File",
            filetypes=[("QuickBooks Company File", "*.QBW *.qbw"), ("All Files", "*.*")],
        )
        if path:
            self.qb_company_file_var.set(path)
            self._persist_settings()
            self._set_status(f"QuickBooks company file set: {Path(path).name}")

    def _qb_client(self) -> QuickBooksClient:
        return QuickBooksClient(company_file_path=self.qb_company_file_var.get().strip())

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)

    def _set_qb_status(self, message: str, state: str = "disconnected") -> None:
        self.qb_status_var.set(message)
        c = UI
        if state == "connected":
            bg, fg = c["success_bg"], c["success"]
        elif state == "pending":
            bg, fg = c["warning_bg"], c["warning"]
        else:
            bg, fg = c["bg_hover"], c["text_secondary"]
        try:
            self.qb_status_pill.configure(bg=bg, highlightbackground=bg)
            self.qb_status_inner.configure(bg=bg, fg=fg)
        except (AttributeError, tk.TclError):
            pass

    def fetch_next_so(self):
        log.info("Fetch Next SO: clicked. company_file_path=%r", self.qb_company_file_var.get().strip())
        self._set_qb_status("Fetching next SO...", state="pending")
        self._set_status("Fetching next Sales Order number from QuickBooks (this may take a moment)...")

        client = self._qb_client()

        def worker():
            try:
                next_no = client.get_next_sales_order_number()
                log.info("Fetch Next SO: success, next=%s", next_no)
                self.root.after(0, lambda: self._fetch_next_so_done(next_no))
            except Exception as exc:
                log.error("Fetch Next SO: failed — %s", exc)
                log.debug("Fetch Next SO traceback:\n%s", traceback.format_exc())
                self.root.after(0, lambda e=exc: self._fetch_next_so_failed(e))

        threading.Thread(target=worker, daemon=True).start()

    def _fetch_next_so_done(self, next_no: str) -> None:
        self.sales_order_no_var.set(next_no)
        self._set_status(f"Fetched next sales order number: {next_no}")
        self._set_qb_status("Connected", state="connected")
        messagebox.showinfo("Success", f"Next Sales Order number: {next_no}")

    def _fetch_next_so_failed(self, exc: Exception) -> None:
        self._set_qb_status("Connection Failed", state="disconnected")
        self._set_status("Fetch Next SO failed. See the Log menu for details.")
        messagebox.showerror(
            "QuickBooks Error",
            f"{exc}\n\nSee the Log menu for details (file: {LOG_PATH}).",
        )

    def _connect_quickbooks_on_startup(self):
        log.info("Auto-connect on startup")
        self._set_qb_status("Connecting...", state="pending")
        self.connect_quickbooks(silent=True)

    def connect_quickbooks(self, silent: bool = False):
        log.info(
            "Connect QuickBooks: clicked (silent=%s). company_file_path=%r",
            silent,
            self.qb_company_file_var.get().strip(),
        )
        try:
            company_name = self._qb_client().test_connection()
            log.info("Connect QuickBooks: success, company=%r", company_name)
            self._set_qb_status(f"Connected: {company_name}", state="connected")
            self._set_status(f"Connected to QuickBooks company: {company_name}")
            if not self.auto_connect_on_startup:
                # First successful connect — remember it so future app launches
                # auto-connect without the user clicking the button.
                log.info("Connect QuickBooks: enabling auto-connect on future startups")
                self.auto_connect_on_startup = True
                try:
                    self._persist_settings()
                except Exception:
                    log.exception("Connect QuickBooks: failed to persist auto-connect flag")
            if not silent:
                messagebox.showinfo("QuickBooks Connected", f"Connected successfully.\nCompany: {company_name}")
        except Exception as exc:
            self._set_qb_status("Not Connected", state="disconnected")
            error_text = str(exc)
            log.error("Connect QuickBooks: failed — %s", error_text)
            log.debug("Connect QuickBooks traceback:\n%s", traceback.format_exc())
            if silent:
                self._set_status("QuickBooks auto-connect failed. You can click Connect to QuickBooks Desktop.")
                return
            if self._is_qb_admin_authorization_error(error_text):
                log.info("Connect QuickBooks: classified as admin authorization required")
                self._set_status("QuickBooks admin authorization required for this company file.")
                messagebox.showwarning(
                    "QuickBooks Admin Authorization Required",
                    "QuickBooks requires an Admin to authorize this app for this company file.\n\n"
                    "Open the QuickBooks Admin Setup guide and complete it once.",
                )
                self.show_qb_admin_setup_guide()
                return

            messagebox.showerror(
                "QuickBooks Connection Error",
                f"{error_text}\n\nTry this:\n"
                "1) Open QuickBooks Desktop manually\n"
                "2) Open the target company file\n"
                "3) Keep QuickBooks on the same Windows user session\n"
                "4) Re-click Connect to QuickBooks Desktop\n\n"
                f"See the Log menu for details (file: {LOG_PATH}).",
            )

    def _is_qb_admin_authorization_error(self, error_text: str) -> bool:
        lowered = error_text.lower()
        return (
            "has not accessed this quickbooks company data file before" in lowered
            or "-2147220456" in lowered
            or "quickbooks administrator must grant" in lowered
        )

    def show_qb_admin_setup_guide(self):
        messagebox.showinfo(
            "QuickBooks Admin Setup",
            "One-time setup by a QuickBooks Admin:\n\n"
            "1) Open QuickBooks Desktop as Admin.\n"
            "2) Open the target company file.\n"
            "3) In this app, click 'Connect to QuickBooks Desktop'.\n"
            "4) Approve the QuickBooks Application Certificate prompt.\n"
            "   - Recommended: 'Yes, always; allow access even if QuickBooks is not running'.\n"
            "5) In QuickBooks, verify under:\n"
            "   Edit > Preferences > Integrated Applications > Company Preferences.\n\n"
            "After this one-time approval, non-admin users can run imports from this app on the same setup.",
        )

    def _edit_output_row(self, event):
        if self.output_df is None or self.output_df.empty:
            return
        item_id = self.output_tree.identify_row(event.y)
        if not item_id:
            return
        row_index = int(item_id)
        source_row_index = self.source_df.index[row_index]
        self._open_source_row_editor(source_row_index)

    def _edit_source_row(self, event):
        if self.source_df is None or self.source_df.empty:
            return
        item_id = self.source_tree.identify_row(event.y)
        if not item_id:
            return
        source_row_index = int(item_id)
        self._open_source_row_editor(source_row_index)

    def _open_source_row_editor(self, source_row_index: int):
        source_row_data = self.source_df.loc[source_row_index].to_dict()
        output_row_index = int(self.source_df.index.get_loc(source_row_index))
        output_row_data = self.output_df.loc[output_row_index].to_dict()
        dlg = RowEditorDialog(self.root, source_row_data, output_row_data, int(source_row_index) + 2)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return
        try:
            self.source_df = self.source_df.astype(object)
            for col, value in dlg.result["source"].items():
                self.source_df.at[source_row_index, col] = value
            self.output_overrides[str(source_row_index)] = dlg.result["output"]
            self._rebuild_previews_from_source()
            self._set_status(f"Saved row updates for Excel line {int(source_row_index) + 2}.")
        except Exception as exc:
            messagebox.showerror(
                "Save Row Error",
                f"Could not save row changes.\n\n{exc}",
            )

    def _validate_settings(self):
        required = {
            "Source File": self.source_path_var.get().strip(),
            "Customer": self.customer_var.get().strip(),
            "Sales Order No": self.sales_order_no_var.get().strip(),
            "Sales Order Date": self.sales_order_date_var.get().strip(),
            "Due Date": self.due_date_var.get().strip(),
        }
        missing = [k for k, v in required.items() if not v]
        if missing:
            raise ValueError(f"Missing required fields: {', '.join(missing)}")
        if not Path(self.source_path_var.get()).exists():
            raise FileNotFoundError("Source file path does not exist.")

    def _ensure_source_loaded(self):
        if self.source_df is not None:
            return
        self.source_df = load_source(self.source_path_var.get().strip())

    def change_pricing_rules(self):
        try:
            self._validate_settings()
            self._ensure_source_loaded()
            dlg = PricingRulesDialog(
                self.root,
                self.source_df,
                pricing_mode=self.pricing_mode,
                use_actual_cost=self.use_actual_cost,
                default_value=self.default_pricing_value,
                existing_brand_values=self.brand_values,
                existing_item_values=self.item_values,
                existing_line_values=self.line_values,
            )
            self.root.wait_window(dlg)
            if dlg.result is None:
                return

            pricing_mode, use_actual_cost, default_value, values = dlg.result
            self.pricing_mode = pricing_mode
            self.use_actual_cost = use_actual_cost
            self.default_pricing_value = default_value
            if pricing_mode == "brand":
                self.brand_values = values
            elif pricing_mode == "item":
                self.item_values = values
            else:
                self.line_values = values
            self._persist_settings()
            mode_text = (
                "per brand"
                if self.pricing_mode == "brand"
                else ("per item (SKU)" if self.pricing_mode == "item" else "per line (Excel row)")
            )
            calc_text = "actual cost/rate values" if self.use_actual_cost else "MSRP multipliers"
            self._set_status(f"Pricing rules saved: {mode_text}, {calc_text}.")
            messagebox.showinfo("Pricing Rules Saved", f"Using {mode_text} with {calc_text}.")
        except Exception as exc:
            messagebox.showerror("Pricing Rules Error", str(exc))

    def _rebuild_previews_from_source(self):
        sales_date_text = self.sales_order_date_var.get().strip()
        due_date_text = self.due_date_var.get().strip()
        if sales_date_text:
            self.sales_order_date_var.set(self._normalize_date_for_display(sales_date_text))
        if due_date_text:
            self.due_date_var.set(self._normalize_date_for_display(due_date_text))
        settings = OrderSettings(
            customer_name=self.customer_var.get().strip(),
            sales_order_no=self.sales_order_no_var.get().strip(),
            sales_order_date=self.sales_order_date_var.get().strip(),
            due_date=self.due_date_var.get().strip(),
            terms=self.terms_var.get().strip() or "Prepaid",
            shipping_method=self.shipping_method_var.get().strip() or "Standard Ground",
            memo=self.memo_var.get().strip(),
            currency=self.currency_var.get().strip() or "USD",
            sales_tax_code=self.tax_code_var.get().strip() or "TAX",
        )
        self.output_df, errors = transform_to_template(
            self.source_df,
            settings,
            self.pricing_mode,
            self.use_actual_cost,
            (
                self.brand_values
                if self.pricing_mode == "brand"
                else (self.item_values if self.pricing_mode == "item" else self.line_values)
            ),
            self.default_pricing_value,
        )
        self._apply_output_overrides()
        self.preview_line_numbers = [int(i) + 2 for i in self.source_df.index.tolist()]
        self._persist_settings()

        for i in self.output_tree.get_children():
            self.output_tree.delete(i)
        for position, (row_idx, row) in enumerate(self.output_df.iterrows()):
            self.output_tree.insert(
                "",
                "end",
                iid=str(row_idx),
                values=self._tree_values_from_row(row, row_idx),
                tags=("odd" if position % 2 else "even",),
            )

        for i in self.source_tree.get_children():
            self.source_tree.delete(i)
        for position, source_row_index in enumerate(self.source_df.index.tolist()):
            self.source_tree.insert(
                "",
                "end",
                iid=str(source_row_index),
                values=self._source_tree_values_from_row(source_row_index),
                tags=("odd" if position % 2 else "even",),
            )

        self.error_text.delete("1.0", tk.END)
        if errors:
            self.error_text.insert(tk.END, "\n".join(errors))
            self._set_status(f"Preview generated with {len(errors)} validation warning(s).")
        else:
            self.error_text.insert(tk.END, "No validation errors found.")
            self._set_status(f"Preview ready: {len(self.output_df)} rows.")

    def build_preview(self):
        try:
            self._validate_settings()
            self.source_df = load_source(self.source_path_var.get().strip())
            self.output_overrides = {}
            if not self.brand_values and not self.item_values and not self.line_values:
                self.change_pricing_rules()
                if not self.brand_values and not self.item_values and not self.line_values:
                    return

            self._rebuild_previews_from_source()
            messagebox.showinfo("Preview Ready", f"Generated {len(self.output_df)} template rows.")
        except Exception as exc:
            messagebox.showerror("Build Preview Error", str(exc))

    def export_file(self):
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("No data", "Build preview first.")
            return
        output_path = self.output_path_var.get().strip()
        if not output_path:
            messagebox.showwarning("Missing output path", "Choose an output file path.")
            return
        self.output_df.to_excel(output_path, index=False, sheet_name="Sales Order")
        self._persist_settings()
        self._set_status(f"Exported template file: {Path(output_path).name}")
        self._show_post_export_actions(Path(output_path))

    def export_saasant_template(self):
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("No data", "Build preview first.")
            return

        downloads_dir = Path.home() / "Downloads"
        date_part = datetime.now().strftime("%m-%d-%Y")
        so_value = self.sales_order_no_var.get().strip() or "SO"
        safe_so_value = "".join(ch for ch in so_value if ch.isalnum() or ch in ("-", "_")) or "SO"
        output_path = downloads_dir / f"SalesOrder_'{safe_so_value}'_{date_part}.xlsx"

        self.output_df.to_excel(output_path, index=False, sheet_name="Sales Order")
        self.output_path_var.set(str(output_path))
        self._persist_settings()
        self._set_status(f"SaaSant export ready: {output_path.name}")
        self._show_post_export_actions(output_path, saasant_ready=True)

    def _show_post_export_actions(self, output_path: Path, saasant_ready: bool = False):
        msg = "File exported successfully."
        if saasant_ready:
            msg = "SaaSant template exported successfully."

        dlg = tk.Toplevel(self.root)
        dlg.title("Export Complete")
        dlg.geometry("580x240")
        dlg.resizable(False, False)
        dlg.configure(bg=UI["bg_window"])
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=(22, 20, 22, 18))
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=msg, style="Header.TLabel").pack(anchor="w")
        ttk.Label(frame, text=str(output_path), wraplength=520, style="SubHeader.TLabel").pack(
            anchor="w", pady=(6, 10)
        )
        if saasant_ready:
            ttk.Label(
                frame,
                text="Upload this file directly in SaaSant.",
                style="SubHeader.TLabel",
            ).pack(anchor="w", pady=(0, 8))

        btns = ttk.Frame(frame)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(
            btns,
            text="Open File",
            command=lambda: self._open_file(output_path),
            style="Accent.TButton",
        ).pack(side="left")
        ttk.Button(
            btns,
            text="Show in Explorer",
            command=lambda: self._reveal_file(output_path),
            style="Quiet.TButton",
        ).pack(side="left", padx=8)
        ttk.Button(btns, text="Close", command=dlg.destroy, style="Quiet.TButton").pack(side="right")

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    def _open_file(self, output_path: Path):
        os.startfile(str(output_path))

    def _reveal_file(self, output_path: Path):
        subprocess.Popen(["explorer", f"/select,{output_path}"])

    def upload_to_quickbooks(self):
        if self.output_df is None or self.output_df.empty:
            log.warning("Upload to QuickBooks: aborted — no preview data")
            messagebox.showwarning("No data", "Build preview first.")
            return
        log.info(
            "Upload to QuickBooks: starting — customer=%r so=%r lines=%d",
            self.customer_var.get().strip(),
            self.sales_order_no_var.get().strip(),
            len(self.output_df),
        )
        try:
            result = self._qb_client().upload_sales_order(
                customer_name=self.customer_var.get().strip(),
                sales_order_no=self.sales_order_no_var.get().strip(),
                txn_date=self._normalize_date_for_qb(self.sales_order_date_var.get().strip()),
                due_date=self._normalize_date_for_qb(self.due_date_var.get().strip()),
                terms=self.terms_var.get().strip() or "Prepaid",
                shipping_method=self.shipping_method_var.get().strip() or "Standard Ground",
                memo=self.memo_var.get().strip(),
                lines=self.output_df.to_dict(orient="records"),
                tax_code=self.tax_code_var.get().strip() or "TAX",
                fallback_item=self.fallback_item_var.get().strip(),
                income_account=self.income_account_var.get().strip(),
            )
            log.info("Upload to QuickBooks: success — %s", result)
            self._set_qb_status("Connected", state="connected")
            self._set_status("Sales order uploaded to QuickBooks successfully.")
            messagebox.showinfo("QuickBooks Upload", result)
        except Exception as exc:
            log.error("Upload to QuickBooks: failed — %s", exc)
            log.debug("Upload traceback:\n%s", traceback.format_exc())
            self._set_qb_status("Connection Failed", state="disconnected")
            messagebox.showerror(
                "QuickBooks Upload Error",
                f"{exc}\n\nSee the Log menu for details (file: {LOG_PATH}).",
            )


def main():
    root = tk.Tk()
    SalesOrderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
