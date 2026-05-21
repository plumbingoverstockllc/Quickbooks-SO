from __future__ import annotations

import json
import logging
import logging.handlers
import os
import re
import sys
import time
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
from transformer import (
    COST_COLUMN,
    OrderSettings,
    ROOM_COLUMN,
    line_pricing_keys,
    load_source,
    load_vendor_multipliers,
    transform_to_template,
    unique_brands,
    unique_skus,
)


DEFAULT_SOURCE = r"C:\Users\QB-PC\Downloads\Project-LisaStrongDesign-EliezerLabkowski301NHighland (1).xls"
DEFAULT_TEMPLATE = r"C:\Users\QB-PC\Downloads\SaasAnt Template for David Meyer.xlsx"
DEFAULT_OUTPUT = r"C:\Users\QB-PC\Downloads\SaaSant Sales Order - Auto Filled.xlsx"
APP_NAME = "DMQuotes"
APP_VERSION = "v1.054b"
# Features still being tested are gated on this flag. The version label is
# the single source of truth: any APP_VERSION ending in 'b' (the beta
# suffix convention used by this app) shows beta-only UI; stable builds
# hide it. The older "Beta" suffix is also recognized for backward compat.
IS_BETA = APP_VERSION.strip().lower().endswith("b") or "beta" in APP_VERSION.lower()
UPDATE_API_URL = "https://api.github.com/repos/plumbingoverstockllc/Quickbooks-SO/releases/latest"
UPDATE_INFO_URL = "https://raw.githubusercontent.com/plumbingoverstockllc/Quickbooks-SO/main/releases/latest.json"
BETA_UPDATE_INFO_URL = "https://raw.githubusercontent.com/plumbingoverstockllc/Quickbooks-SO/main/releases/beta.json"
# Where "Send Logs to Support" delivers. The log is uploaded to a paste
# service for a viewable link, and that link is emailed here via FormSubmit
# (a one-time "Activate Form" click on the first email enables delivery).
SUPPORT_EMAIL = "mosheyadelman@gmail.com"
# v1.025b: app was renamed from "QB Sales Order Converter" to "DMQuotes".
# Settings/log directory keeps the legacy name so existing users' saved
# configuration (brand multipliers, file paths, window geometry, etc.)
# carries over without a migration step.
LEGACY_SETTINGS_FOLDER = "QB Sales Order Converter"
SETTINGS_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / LEGACY_SETTINGS_FOLDER
SETTINGS_PATH = SETTINGS_DIR / "settings.json"
LOG_PATH = SETTINGS_DIR / "app.log"


# --- "Time saved" estimate -------------------------------------------------
# Rough model of how long it takes a person at an average pace to key a sales
# order into QuickBooks Desktop by hand instead of importing it. Used for the
# friendly "you just saved X min Y sec" message after a successful upload.
#   - SECONDS_PER_LINE: copy/paste + tab through SKU, description, qty, rate,
#     and tax code for one line, then fix the inevitable typo. 35s is a
#     conservative average — try doing 40 lines and you'll wish it were lower.
#   - HEADER_OVERHEAD_SECONDS: customer lookup, SO number, both dates, terms,
#     and shipping method on the order header before any line is even touched.
#   - MANUAL_SECONDS_PER_ROOM: each room group needs a header line the user
#     would otherwise have to type BY HAND — room name in ALL CAPS, wrapped
#     in ** markers, plus the blank separator lines between groups. Fiddly
#     and easy to typo, so it carries its own per-room penalty.
# The post-upload message compares this manual estimate against the *actual*
# measured import time, so there's no need to fudge an import-overhead figure.
MANUAL_SECONDS_PER_LINE = 35
MANUAL_HEADER_OVERHEAD_SECONDS = 45
MANUAL_SECONDS_PER_ROOM = 25


def _count_room_groups(lines) -> int:
    """Count how many **ROOM** header lines the upload inserts — i.e. the
    number of contiguous room groups, matching the grouping logic in
    quickbooks_client.upload_sales_order. Each one is a header the user
    would have had to type (in caps, with ** markers) by hand."""
    groups = 0
    current_room = None
    for line in lines or []:
        room = ""
        if isinstance(line, dict):
            room = str(line.get("Room", "") or "").strip()
        if current_room is None or room != current_room:
            groups += 1
            current_room = room
    return groups


def _estimate_manual_seconds(num_lines: int, num_rooms: int = 0) -> int:
    """Estimated seconds to enter `num_lines` line items (plus `num_rooms`
    hand-typed room headers) by hand, order header included."""
    return (
        max(0, num_lines) * MANUAL_SECONDS_PER_LINE
        + max(0, num_rooms) * MANUAL_SECONDS_PER_ROOM
        + MANUAL_HEADER_OVERHEAD_SECONDS
    )


def _format_duration(total_seconds: int) -> str:
    """Human phrasing like '6 minutes and 35 seconds' / '45 seconds'."""
    total_seconds = max(0, int(round(total_seconds)))
    minutes, seconds = divmod(total_seconds, 60)
    min_part = f"{minutes} minute{'s' if minutes != 1 else ''}"
    sec_part = f"{seconds} second{'s' if seconds != 1 else ''}"
    if minutes and seconds:
        return f"{min_part} and {sec_part}"
    if minutes:
        return min_part
    return sec_part


def _resource_path(filename: str) -> Path:
    """Resolve a bundled resource path.

    When running from a PyInstaller --onefile build, data files extracted
    into sys._MEIPASS. When running from source, the file sits next to
    app.py.
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / filename


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
    # v1.035: palette retuned to the QuickBooks Desktop Enterprise green
    # two-tone (dark forest sidebar + bright QB-green accents) so the app
    # stops reading like a doctor's office. Main content stays on a near-
    # white surface for form legibility; everything chrome (sidebar, buttons,
    # focus rings, accent strip) is green.
    "bg_window": "#FFFFFF",
    "bg_card": "#FFFFFF",
    "bg_subtle": "#F2F5F1",
    "bg_hover": "#E8F1E4",
    "bg_pressed": "#D1E5C8",
    "bg_sidebar": "#1A3A33",
    "sidebar_text": "#FFFFFF",
    "sidebar_text_muted": "#9CB5AE",
    "border": "#E1E6DE",
    "border_strong": "#C7CFC2",
    "border_inner_light": "#FFFFFF",
    "text_primary": "#0F1F1B",
    "text_secondary": "#3F5651",
    "text_tertiary": "#7A8E89",
    "accent": "#2CA01C",
    "accent_light": "#3DB52A",
    "accent_hover": "#26901A",
    "accent_pressed": "#1F7A15",
    "accent_dark": "#155A0E",
    "accent_bg": "#E3F4DF",
    "btn_light_top": "#FFFFFF",
    "btn_light_bottom": "#F1F5EE",
    "btn_light_border": "#D5DDD0",
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
        # Keys the user marked for deletion, and combine mappings
        # (source key -> target key) made in this dialog session.
        self.deleted: set[str] = set()
        self.combines: dict[str, str] = {}

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

        # Buttons FIRST, pinned to the bottom, so the scrolling brand list can
        # never push them off-screen.
        buttons = ttk.Frame(self)
        buttons.pack(fill="x", side="bottom", padx=18, pady=(0, 16))
        ttk.Button(buttons, text="Use Pricing Rules", command=self._submit, style="Primary.TButton").pack(
            side="right"
        )
        ttk.Button(buttons, text="Cancel", command=self._cancel, style="Quiet.TButton").pack(
            side="right", padx=(0, 8)
        )
        ttk.Button(buttons, text="Combine Brands…", command=self._combine_dialog, style="Quiet.TButton").pack(
            side="left"
        )

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

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda e: self._cancel())

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
        # v1.040: show the UNION of (a) keys present in the currently-loaded
        # source file and (b) keys already saved from previous sessions.
        # Previously only the current file's brands/SKUs were listed, so a
        # user who opened "Change Pricing Rules" with a different file (or no
        # file) loaded couldn't see — let alone edit — the multipliers they'd
        # already configured. They're persisted in settings.json the whole
        # time; they just weren't being displayed.
        if mode == "brand":
            values_source = self.existing_brand_values
            file_keys = unique_brands(self.source_df) if self.source_df is not None else []
            all_keys = sorted(set(file_keys) | set(values_source.keys()))
            keys = [(k, k) for k in all_keys]
        elif mode == "item":
            values_source = self.existing_item_values
            file_keys = unique_skus(self.source_df) if self.source_df is not None else []
            all_keys = sorted(set(file_keys) | set(values_source.keys()))
            keys = [(k, k) for k in all_keys]
        else:
            # Per-line keys are tied to the current file's Excel row numbers,
            # so they aren't meaningful across files — keep current-file rows
            # plus any saved line keys so nothing silently vanishes.
            values_source = self.existing_line_values
            keys = line_pricing_keys(self.source_df) if self.source_df is not None else []
            seen = {k for k, _ in keys}
            for saved_key in sorted(values_source.keys()):
                if saved_key not in seen:
                    keys.append((saved_key, f"Line {saved_key}"))
        if not keys:
            ttk.Label(
                self.inner,
                text=(
                    "No brands to configure yet. Load a source file (Step 1) or "
                    "saved pricing will appear here once you've set some."
                ),
                style="Card.TLabel",
                wraplength=440,
                justify="left",
            ).pack(anchor="w", padx=10, pady=10)
            return

        self._all_keys = [k for k, _ in keys]
        for key, label in keys:
            if key in self.deleted or key in self.combines:
                continue  # hidden: marked for delete or combined away
            row = ttk.Frame(self.inner, style="Card.TFrame")
            row.pack(fill="x", pady=3, padx=10)
            ttk.Label(row, text=label, width=30, style="Card.TLabel").pack(side="left")
            # Delete (✕) removes the brand entirely.
            ttk.Button(
                row, text="✕", width=2, style="Quiet.TButton",
                command=lambda k=key, lbl=label: self._mark_deleted(k, lbl),
            ).pack(side="right", padx=(6, 0))
            existing_value = ""
            if key in values_source:
                existing_value = str(values_source.get(key, ""))
            var = tk.StringVar(value=existing_value)
            self.value_vars[key] = var
            ttk.Entry(row, textvariable=var, width=14).pack(side="right")

    def _mark_deleted(self, key: str, label: str) -> None:
        if not messagebox.askyesno(
            "Delete Brand",
            f"Delete '{label}' from the pricing list?\n\n"
            "It won't be priced automatically anymore.",
            parent=self,
        ):
            return
        self.deleted.add(key)
        self.value_vars.pop(key, None)
        self._render_value_rows()

    def _combine_dialog(self) -> None:
        """Merge one brand into another: the source's quotes will use the
        target brand's pricing from now on (an alias)."""
        keys = sorted(getattr(self, "_all_keys", []))
        if len(keys) < 2:
            messagebox.showinfo("Combine Brands", "Need at least two brands to combine.", parent=self)
            return
        sub = tk.Toplevel(self)
        sub.title("Combine Brands")
        sub.configure(bg=UI["bg_window"])
        sub.geometry("460x220")
        sub.resizable(False, False)
        sub.transient(self)
        sub.grab_set()

        frame = ttk.Frame(sub, padding=(20, 18, 20, 16))
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Combine Brands", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text="The first brand will be merged into the second and use its pricing.",
            style="SubHeader.TLabel", wraplength=410, justify="left",
        ).pack(anchor="w", pady=(2, 12))

        a_var, b_var = tk.StringVar(), tk.StringVar()
        rowa = ttk.Frame(frame); rowa.pack(fill="x", pady=3)
        ttk.Label(rowa, text="Merge this brand:", width=18).pack(side="left")
        ttk.Combobox(rowa, textvariable=a_var, values=keys, state="readonly").pack(side="left", fill="x", expand=True)
        rowb = ttk.Frame(frame); rowb.pack(fill="x", pady=3)
        ttk.Label(rowb, text="Into this brand:", width=18).pack(side="left")
        ttk.Combobox(rowb, textvariable=b_var, values=keys, state="readonly").pack(side="left", fill="x", expand=True)

        def do_combine():
            a, b = a_var.get().strip(), b_var.get().strip()
            if not a or not b:
                messagebox.showerror("Combine Brands", "Pick both brands.", parent=sub)
                return
            if a == b:
                messagebox.showerror("Combine Brands", "Pick two different brands.", parent=sub)
                return
            self.combines[a] = b
            self.value_vars.pop(a, None)
            sub.destroy()
            self._render_value_rows()

        btns = ttk.Frame(frame); btns.pack(fill="x", side="bottom", pady=(14, 0))
        ttk.Button(btns, text="Combine", command=do_combine, style="Primary.TButton").pack(side="right")
        ttk.Button(btns, text="Cancel", command=sub.destroy, style="Quiet.TButton").pack(side="right", padx=(0, 8))
        sub.protocol("WM_DELETE_WINDOW", sub.destroy)

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
            set(self.deleted),
            dict(self.combines),
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

        # Buttons pinned to the bottom first so they're never pushed off by
        # the expanding tab content.
        btns = ttk.Frame(self, padding=(20, 12, 20, 16))
        btns.pack(fill="x", side="bottom")
        ttk.Button(btns, text="Save Changes", command=self._save, style="Primary.TButton").pack(side="right")
        ttk.Button(btns, text="Cancel", command=self._cancel, style="Quiet.TButton").pack(
            side="right", padx=(0, 8)
        )

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
        self.root.title(f"DMQuotes {APP_VERSION} · QuickBooks Sales Orders")
        # v1.034: set the window/taskbar icon to the DMQuotes logo. The .ico
        # is bundled next to the bootstrap by PyInstaller (--add-data) and
        # resolved via _resource_path. iconbitmap(default=...) makes Tk apply
        # the icon to every Toplevel as well, not just root. Wrap in try/except
        # because the .ico may be missing in dev runs and iconbitmap can fail
        # on some Windows configurations.
        try:
            icon_path = _resource_path("DMQuotes.ico")
            if icon_path.exists():
                self.root.iconbitmap(default=str(icon_path))
        except (tk.TclError, OSError):
            pass
        # v1.032: bump Tk's global scaling by 25% so every point-sized
        # font and text metric renders larger uniformly. Pixel-based
        # sizes (minsize, sidebar width, logo image) are scaled manually
        # below since Tk scaling doesn't touch those.
        try:
            base_scaling = self.root.tk.call("tk", "scaling")
            self.root.tk.call("tk", "scaling", base_scaling * 1.25)
        except tk.TclError:
            pass
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self._configure_styles()
        # Load settings BEFORE the geometry block — the saved window size
        # lives there.
        self.settings = self._load_settings()

        # Window geometry: prefer the value the user last resized to (saved
        # in settings.json as window_geometry), otherwise open at a
        # comfortable default of 80% width × 75% height, centered.
        try:
            self.root.update_idletasks()
            screen_w = self.root.winfo_screenwidth() or 1600
            screen_h = self.root.winfo_screenheight() or 1040
        except tk.TclError:
            screen_w, screen_h = 1600, 1040
        saved_geom = (self.settings.get("window_geometry") or "").strip()
        if saved_geom and self._geometry_fits_screen(saved_geom, screen_w, screen_h):
            self.root.geometry(saved_geom)
        else:
            win_w = max(1380, int(screen_w * 0.80))
            win_h = max(900, int(screen_h * 0.75))
            x = max(0, (screen_w - win_w) // 2)
            y = max(0, (screen_h - win_h) // 2)
            self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        try:
            # v1.032: minsize bumped to 1500×900 to match the 25%-larger
            # UI scale. The form's natural minimum width grew with the
            # Tk scaling factor; this floor keeps the entries from
            # clipping at small window sizes.
            self.root.minsize(1500, 900)
        except tk.TclError:
            pass
        # Persist window size + position on close so future launches
        # remember whatever the user resized to.
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

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
        self.income_account_var = tk.StringVar(value=self.settings.get("income_account", "Sales Non Inventory"))
        self.sales_tax_item_var = tk.StringVar(value=self.settings.get("sales_tax_item", ""))
        self.room_grouping_var = tk.BooleanVar(value=bool(self.settings.get("room_grouping_enabled", False)))
        # Default ON: auto-connect on every launch. If the attach fails the
        # status pill just shows "Not Connected" and the user can click
        # Connect manually. The user can opt out by setting
        # "auto_connect_on_startup": false in settings.json.
        self.auto_connect_on_startup = bool(self.settings.get("auto_connect_on_startup", True))
        # Version the user dismissed with "Not Now"; the hourly auto-check
        # won't re-prompt for it until a newer version appears.
        self._snoozed_update_version: str | None = None
        self.pricing_mode = self.settings.get("pricing_mode", "brand")
        self.use_actual_cost = bool(self.settings.get("use_actual_cost", False))
        self.default_pricing_value = float(self.settings.get("default_pricing_value", 0.4))
        self.brand_values: dict[str, float] = self.settings.get("brand_values", {})
        self.item_values: dict[str, float] = self.settings.get("item_values", {})
        self.line_values: dict[str, float] = self.settings.get("line_values", {})
        # Brands whose vendor-list multiplier is tiered/textual (not a single
        # number). We keep the raw note so the prompt can show it and the user
        # picks the right tier each time.
        self.vendor_notes: dict[str, str] = self.settings.get("vendor_notes", {})
        # Brand aliases: source-quote brand name -> the real brand name in the
        # system. Lets a wrong/abbreviated name on a quote (e.g. "WYC") map to
        # an existing brand's pricing, remembered for next time.
        self.brand_aliases: dict[str, str] = self.settings.get("brand_aliases", {})
        # Brands the user deleted; suppressed even if present in the bundle.
        self.deleted_brands: set[str] = set(self.settings.get("deleted_brands", []))
        # Vendor price database bundled with the app (no manual import needed).
        # vendor_clean: brand -> single multiplier (baseline, always available).
        # Tiered brands' notes are merged into vendor_notes. User-saved values
        # in brand_values still take precedence over these baselines.
        self.vendor_clean: dict[str, float] = {}
        self._load_bundled_vendor_pricing()
        # Per-order "variable" multipliers entered for tiered brands that the
        # user chose NOT to lock. Used for the current order only; not saved,
        # so those brands are asked again on the next order.
        self.session_brand_values: dict[str, float] = {}
        self._session_source_path: str | None = None
        self.status_var = tk.StringVar(value="Ready. Load source file, then build preview.")
        self.qb_status_var = tk.StringVar(value="Go to Setup → Connect to QuickBooks Desktop to connect")
        self.qb_status_label: ttk.Label | None = None

        self._build_layout()
        self._build_menu()
        self._set_qb_status(self._not_connected_message(), state="disconnected")
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
        # v1.039: "STEP N" eyebrow labels above each action area, so the
        # workflow is self-explanatory at a glance. Uppercase, accent-green,
        # tight letter-spacing simulated via Segoe UI Black at a small size.
        self.style.configure(
            "StepEyebrow.TLabel",
            background=c["bg_card"],
            foreground=c["accent"],
            font=("Segoe UI Black", 8),
        )
        self.style.configure(
            "StepTitle.TLabel",
            background=c["bg_card"],
            foreground=c["text_primary"],
            font=("Segoe UI Semibold", 11),
        )
        # Step badge variant for the action row, which sits on the window
        # background (not the card), so the bg matches bg_window.
        self.style.configure(
            "StepEyebrowOnWindow.TLabel",
            background=c["bg_window"],
            foreground=c["accent"],
            font=("Segoe UI Black", 8),
        )
        self.style.configure(
            "StepTitleOnWindow.TLabel",
            background=c["bg_window"],
            foreground=c["text_primary"],
            font=("Segoe UI Semibold", 11),
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
            bordercolor=c["border"],
            lightcolor=c["border"],
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

        # Error variant used when a required field is empty at submit time.
        # The border + background go red until the user types something or
        # focuses out into a valid state.
        self.style.configure(
            "Error.TEntry",
            fieldbackground=c["danger_bg"],
            foreground=c["text_primary"],
            bordercolor=c["danger"],
            lightcolor=c["danger"],
            darkcolor=c["danger"],
            insertcolor=c["text_primary"],
            borderwidth=2,
            padding=(10, 8),
            relief="solid",
        )
        self.style.map(
            "Error.TEntry",
            bordercolor=[("focus", c["danger"]), ("hover", c["danger"])],
            lightcolor=[("focus", c["danger"]), ("hover", c["danger"])],
            darkcolor=[("focus", c["danger"]), ("hover", c["danger"])],
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
            bordercolor=c["accent"],
            lightcolor=c["accent"],
            darkcolor=c["accent"],
            borderwidth=0,
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
                ("active", c["accent_light"]),
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

        # v1.022b: flat, clean buttons. EZTools-inspired look — single
        # 1px border, no shadow, just enough padding to breathe. Pressed
        # state uses the subtle pressed-bg; hover lifts to bg_hover.
        self.style.configure(
            "Accent.TButton",
            padding=(16, 9),
            font=("Segoe UI Semibold", 10),
            background=c["bg_card"],
            foreground=c["accent"],
            bordercolor=c["border"],
            lightcolor=c["border"],
            darkcolor=c["border"],
            borderwidth=1,
            focusthickness=0,
            relief="solid",
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
            relief=[("pressed", "sunken")],
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
            bordercolor=c["border"],
            lightcolor=c["border"],
            darkcolor=c["border"],
            borderwidth=1,
            focusthickness=0,
            relief="solid",
        )

        # Green/Success button for "commit" actions like Upload to QuickBooks.
        self.style.configure(
            "Success.TButton",
            padding=(20, 10),
            font=("Segoe UI Semibold", 10),
            background=c["success"],
            foreground="#FFFFFF",
            bordercolor=c["success"],
            lightcolor=c["success"],
            darkcolor=c["success"],
            borderwidth=0,
            focusthickness=0,
            relief="flat",
        )
        self.style.map(
            "Success.TButton",
            background=[
                ("pressed", "#0E5E2C"),
                ("active", "#127039"),
                ("disabled", c["border_strong"]),
            ],
            foreground=[("disabled", c["bg_card"])],
            relief=[("pressed", "sunken")],
        )
        self.style.map(
            "Quiet.TButton",
            background=[
                ("pressed", c["bg_pressed"]),
                ("active", c["bg_hover"]),
            ],
            foreground=[("disabled", c["text_tertiary"])],
            relief=[("pressed", "sunken")],
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

        setup_menu = tk.Menu(
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
        setup_menu.add_command(label="Connect to QuickBooks Desktop", command=self.connect_quickbooks)
        setup_menu.add_command(label="QuickBooks Admin Setup", command=self.show_qb_admin_setup_guide)
        setup_menu.add_separator()
        setup_menu.add_command(label="Import Vendor Price List…", command=self.import_vendor_price_list)
        setup_menu.add_command(label="Change Pricing Rules", command=self.change_pricing_rules)
        setup_menu.add_command(label="Export Template (Custom Path)", command=self.export_file)
        menu_bar.add_cascade(label="Setup", menu=setup_menu)

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
        help_menu.add_command(label="Check for Beta Update", command=self.check_for_beta_update)
        help_menu.add_separator()
        help_menu.add_command(
            label="About",
            command=lambda: messagebox.showinfo(
                "About",
                f"{APP_NAME} {APP_VERSION}\n"
                "QuickBooks Sales Order Uploader\n\n"
                "Created by Moshe Adelman\n"
                "For help email hello@shimiralabs.com\n"
                "Shimiralabs.com",
            ),
        )
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
        log_menu.add_command(label="Send Logs to Support", command=self.send_logs_to_support)
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

    def send_logs_to_support(self) -> None:
        """Upload the current log file to a paste service and email the link
        straight to support (SUPPORT_EMAIL) via FormSubmit, so issues can be
        diagnosed remotely. The link is also copied to the clipboard as a
        backup. No server of our own required.

        The log contains file paths, the QuickBooks company name, and recent
        SKUs/customer names — but no passwords. We say so and make it opt-in.
        """
        if not messagebox.askyesno(
            "Send Logs to Support",
            "This uploads your log file and emails it to support so they can "
            "look into the problem.\n\n"
            "The log includes file paths, your QuickBooks company name, and "
            "recent activity (SKUs, customers) — but no passwords.\n\n"
            "Send now?",
        ):
            return
        self._set_status("Sending log to support…")

        def worker() -> None:
            data = LOG_PATH.read_bytes() if LOG_PATH.exists() else b"(log file empty)"
            ua = f"{APP_NAME}/{APP_VERSION} (support log upload)"
            errors: list[str] = []

            # Primary: paste.rs — raw POST body, returns the URL as plain text.
            def via_paste_rs() -> str:
                req = urllib.request.Request(
                    "https://paste.rs/",
                    data=data,
                    headers={"User-Agent": ua, "Content-Type": "text/plain"},
                    method="POST",
                )
                with urllib.request.urlopen(req, timeout=30) as resp:
                    return resp.read().decode("utf-8", "replace").strip()

            # Fallback: 0x0.st — multipart/form-data.
            def via_0x0() -> str:
                boundary = "----DMQuotes" + os.urandom(8).hex()
                body = (
                    f"--{boundary}\r\n"
                    f'Content-Disposition: form-data; name="file"; '
                    f'filename="{APP_NAME}-{APP_VERSION}.log"\r\n'
                    f"Content-Type: text/plain\r\n\r\n"
                ).encode("utf-8") + data + f"\r\n--{boundary}--\r\n".encode("utf-8")
                req = urllib.request.Request(
                    "https://0x0.st",
                    data=body,
                    headers={
                        "Content-Type": f"multipart/form-data; boundary={boundary}",
                        "User-Agent": ua,
                    },
                    method="POST",
                )
                with urllib.request.urlopen(req, timeout=30) as resp:
                    return resp.read().decode("utf-8", "replace").strip()

            link = None
            for upload in (via_paste_rs, via_0x0):
                try:
                    result = upload()
                    if result.startswith("http"):
                        link = result
                        log.info("send_logs_to_support: uploaded log -> %s", link)
                        break
                    errors.append(f"{upload.__name__}: unexpected response {result[:120]!r}")
                except Exception as exc:
                    log.exception("send_logs_to_support: %s failed", upload.__name__)
                    errors.append(f"{upload.__name__}: {exc}")

            if not link:
                msg = "; ".join(errors)
                self.root.after(0, lambda: self._log_upload_failed(msg))
                return

            emailed = self._email_log_link(link)
            self.root.after(0, lambda l=link, e=emailed: self._show_log_link(l, e))

        threading.Thread(target=worker, daemon=True).start()

    def _email_log_link(self, link: str) -> bool:
        """Email the uploaded-log link to SUPPORT_EMAIL via FormSubmit's
        keyless endpoint. Returns True on a successful submission. (The very
        first email triggers a one-time 'Activate Form' confirmation to the
        support address; once clicked, all future sends arrive automatically.)
        """
        try:
            payload = json.dumps({
                "subject": f"{APP_NAME} log — {APP_VERSION}",
                "message": (
                    f"A {APP_NAME} user sent their log for support.\n\n"
                    f"Version: {APP_VERSION}\n"
                    f"Log link: {link}\n"
                ),
            }).encode("utf-8")
            req = urllib.request.Request(
                f"https://formsubmit.co/ajax/{SUPPORT_EMAIL}",
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                    # FormSubmit's AJAX endpoint requires a web Origin/Referer.
                    "Origin": "https://dmquotes.shimiralabs.com",
                    "Referer": "https://dmquotes.shimiralabs.com/",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0 Safari/537.36",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=30) as resp:
                body = resp.read().decode("utf-8", "replace")
            log.info("send_logs_to_support: email response %s", body[:200])
            return '"success":"true"' in body.replace(" ", "")
        except Exception:
            log.exception("send_logs_to_support: email step failed")
            return False

    def _show_log_link(self, link: str, emailed: bool = False) -> None:
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(link)
        except tk.TclError:
            pass
        if emailed:
            self._set_status("Log sent to support.")
            messagebox.showinfo(
                "Logs Sent",
                "Your log was sent to support. They'll take a look.\n\n"
                "A copy of the link is on your clipboard if you need it:\n\n"
                f"{link}",
            )
        else:
            self._set_status("Log uploaded. Link copied to clipboard.")
            messagebox.showinfo(
                "Log Uploaded",
                "Your log was uploaded (the email step didn't go through, but "
                "that's OK). Send this link to support — it's on your "
                "clipboard:\n\n"
                f"{link}",
            )

    def _log_upload_failed(self, error: str) -> None:
        self._set_status("Log upload failed.")
        messagebox.showerror(
            "Couldn't Send Logs",
            "The log couldn't be uploaded automatically:\n\n"
            f"{error}\n\n"
            "You can still send it manually — use Log → Show Log Folder and "
            f"email the file to hello@shimiralabs.com.\n\n(File: {LOG_PATH})",
        )

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

    def _version_tuple(self, version_str: str) -> tuple[int, int, int, int]:
        """Parse a version string into a tuple suitable for ordering.

        Returns (major, minor, patch, stability) where stability is 0 for a
        beta build and 1 for a stable build. Sorting tuples then naturally
        orders v1.002b *between* v1.001 and v1.002 (stable):
            v1.001    -> (1, 1, 0, 1)
            v1.002b   -> (1, 2, 0, 0)
            v1.002    -> (1, 2, 0, 1)
        Beta is recognized by the 'b' suffix OR the older "Beta" label.
        """
        raw = version_str.strip()
        is_beta = raw.lower().endswith("b") or "beta" in raw.lower()
        cleaned = raw.lower().replace("v", "").replace("beta", "").strip()
        cleaned = cleaned.split("-")[0]
        # Strip a trailing 'b' from the last numeric part (e.g. "1.002b").
        if cleaned.endswith("b"):
            cleaned = cleaned[:-1]
        parts = cleaned.split(".")
        nums: list[int] = []
        for part in parts[:3]:
            num = "".join(ch for ch in part if ch.isdigit())
            nums.append(int(num) if num else 0)
        while len(nums) < 3:
            nums.append(0)
        nums.append(0 if is_beta else 1)
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

    def _collect_cumulative_notes(
        self, current_tuple: tuple, latest_tuple: tuple, fallback_notes: str
    ) -> str:
        """Build a combined 'What's new' covering every stable release newer
        than the installed build, up to and including the latest. When a user
        is several versions behind, they see each version's notes — newest
        first — instead of only the latest one.

        Falls back to `fallback_notes` (the single latest release body) if the
        releases list can't be fetched.
        """
        api_url = (
            "https://api.github.com/repos/plumbingoverstockllc/Quickbooks-SO/releases"
            f"?per_page=100&t={int(datetime.now().timestamp())}"
        )
        try:
            request = urllib.request.Request(
                api_url,
                headers={
                    "Cache-Control": "no-cache",
                    "User-Agent": f"{APP_NAME}/{APP_VERSION}",
                    "Accept": "application/vnd.github+json",
                },
            )
            with urllib.request.urlopen(request, timeout=10) as resp:
                releases = json.loads(resp.read().decode("utf-8"))

            entries: list[tuple[tuple, str, str]] = []
            for rel in releases:
                if rel.get("prerelease") or rel.get("draft"):
                    continue
                tag = str(rel.get("tag_name", "")).strip()
                ver = tag.lstrip("vV")
                if not ver:
                    continue
                try:
                    vt = self._version_tuple(ver)
                except Exception:
                    continue
                # Strictly newer than what's installed, up to the latest.
                if vt <= current_tuple or vt > latest_tuple:
                    continue
                body = self._clean_release_notes(str(rel.get("body", "")).strip())
                entries.append((vt, ver, body))

            if not entries:
                return fallback_notes

            entries.sort(key=lambda e: e[0], reverse=True)
            blocks: list[str] = []
            for _, ver, body in entries:
                header = f"━━━━━━━━━  v{ver}  ━━━━━━━━━"
                blocks.append(f"{header}\n{body}" if body else header)
            return "\n\n".join(blocks)
        except Exception:
            log.exception("_collect_cumulative_notes: failed; using single-release notes")
            return fallback_notes

    def _fetch_beta_release_info(self) -> dict:
        """Find the latest beta release.

        Prefers the GitHub Releases API (no edge cache, sees prereleases)
        and falls back to the raw beta.json feed only when the API is
        unreachable. The raw feed lives on raw.githubusercontent.com which
        has a ~5-minute CDN cache, so betas pushed less than a few minutes
        ago can otherwise look invisible from inside the app.
        """
        api_url = (
            "https://api.github.com/repos/plumbingoverstockllc/Quickbooks-SO/releases"
            f"?per_page=20&t={int(datetime.now().timestamp())}"
        )
        try:
            request = urllib.request.Request(
                api_url,
                headers={
                    "Cache-Control": "no-cache",
                    "User-Agent": f"{APP_NAME}/{APP_VERSION}",
                    "Accept": "application/vnd.github+json",
                },
            )
            with urllib.request.urlopen(request, timeout=10) as resp:
                releases = json.loads(resp.read().decode("utf-8"))
            prereleases = [r for r in releases if r.get("prerelease")]
            if prereleases:
                latest = prereleases[0]
                tag = (latest.get("tag_name") or "").lstrip("vV")
                exe_asset = None
                for asset in latest.get("assets", []):
                    name = asset.get("name", "")
                    if name.lower().endswith(".exe"):
                        exe_asset = asset
                        break
                if tag and exe_asset and exe_asset.get("browser_download_url"):
                    body = (latest.get("body") or "").strip()
                    return {
                        "version": tag,
                        "url": exe_asset["browser_download_url"],
                        "sha256": "",
                        "notes": body,
                    }
        except Exception:
            log.exception("_fetch_beta_release_info: API path failed; falling back to raw")

        cache_bust_url = f"{BETA_UPDATE_INFO_URL}?t={int(datetime.now().timestamp())}"
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

    def check_for_beta_update(self) -> None:
        """Manually-triggered beta update check. Offers the beta even when
        its version is lower than the current installed build, so users can
        roll forward to a beta from a stable in the same line."""
        self._set_status("Checking for beta update...")
        try:
            info = self._fetch_beta_release_info()
        except Exception as exc:
            log.exception("Beta update check failed")
            messagebox.showerror("Beta Update Check Failed", str(exc))
            self._set_status("Beta update check failed.")
            return
        latest_version = info["version"]
        download_url = info["url"]
        sha256_hash = info["sha256"]
        notes = info["notes"]
        if not latest_version or not download_url:
            messagebox.showerror(
                "Beta Update Check Failed",
                "The beta release feed didn't include a version or download URL.",
            )
            return
        try:
            current = self._version_tuple(APP_VERSION)
            available = self._version_tuple(latest_version)
        except Exception:
            current, available = (0, 0, 0, 0), (0, 0, 0, 0)
        if available == current:
            messagebox.showinfo(
                "Beta Update",
                f"You're already on v{latest_version}.",
            )
            return
        if available < current:
            messagebox.showinfo(
                "Beta Update",
                f"You're on {APP_VERSION}, which is already newer than the "
                f"latest beta (v{latest_version}). Nothing to install.",
            )
            return
        if not self._show_update_dialog(latest_version, notes, is_beta=True):
            return
        self._download_and_run_update(download_url, sha256_hash)

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

            # On automatic (silent) checks, don't re-nag for a version the
            # user already dismissed with "Not Now" this session. A manual
            # check (Help menu) always shows. A genuinely newer version than
            # the one snoozed will still prompt.
            if silent and latest_version == getattr(self, "_snoozed_update_version", None):
                self._set_status(f"Update available ({latest_version}); reminder snoozed.")
                return

            # When the user is several versions behind, show every version's
            # notes between what they have and the newest — not just the
            # latest release body.
            notes = self._collect_cumulative_notes(current_version, available_version, notes)

            if not self._show_update_dialog(latest_version, notes):
                # Remember the decline so the hourly check doesn't keep
                # popping the same version.
                self._snoozed_update_version = latest_version
                self._set_status(f"Update {latest_version} available — install later from the Help menu.")
                return

            self._download_and_run_update(download_url, sha256_hash)
        except Exception as exc:
            if not silent:
                messagebox.showerror("Update Check Failed", str(exc))
            self._set_status("Update check failed. Use 'Check for Updates' to retry.")

    def check_for_updates_on_startup(self) -> None:
        self._set_status("Checking for updates...")
        self.check_for_updates(silent=True)
        # Kick off the recurring hourly check.
        self._schedule_hourly_update_check()

    def _schedule_hourly_update_check(self) -> None:
        """(Re)arm the once-an-hour background update check."""
        try:
            self.root.after(3_600_000, self._hourly_update_check)
        except tk.TclError:
            pass

    def _hourly_update_check(self) -> None:
        """Fires every hour: quietly check for an update and prompt only if a
        new one (that hasn't already been snoozed) is found, then re-arm."""
        try:
            self.check_for_updates(silent=True)
        except Exception:
            log.exception("_hourly_update_check failed")
        finally:
            self._schedule_hourly_update_check()

    def _show_created_items_report(self, upload_message: str, created: list[dict]) -> None:
        """Display a report listing every item the upload had to create in
        QuickBooks. Triggered only on beta builds.
        """
        c = UI
        dlg = tk.Toplevel(self.root)
        dlg.title("Sales Order Uploaded")
        dlg.configure(bg=c["bg_window"])
        dlg.geometry("640x460")
        try:
            dlg.minsize(520, 380)
        except tk.TclError:
            pass
        dlg.transient(self.root)
        dlg.grab_set()

        accent = tk.Frame(dlg, bg=c["accent"], height=3)
        accent.pack(fill="x", side="top")

        body = tk.Frame(dlg, bg=c["bg_window"])
        body.pack(fill="both", expand=True, padx=22, pady=(18, 14))

        tk.Label(
            body,
            text=upload_message,
            bg=c["bg_window"],
            fg=c["accent"],
            font=("Segoe UI Semibold", 14),
            anchor="w",
            wraplength=580,
            justify="left",
        ).pack(fill="x")
        tk.Label(
            body,
            text=f"{len(created)} new item(s) were created in QuickBooks during this upload:",
            bg=c["bg_window"],
            fg=c["text_secondary"],
            font=("Segoe UI", 10),
            anchor="w",
            wraplength=580,
            justify="left",
        ).pack(fill="x", pady=(6, 10))

        def copy_to_clipboard() -> None:
            text = "SKU\tDescription\n" + "\n".join(
                f"{item.get('sku', '')}\t{item.get('description', '')}" for item in created
            )
            self.root.clipboard_clear()
            self.root.clipboard_append(text)

        # Buttons pinned to the bottom BEFORE the expanding table, so they're
        # always visible.
        button_row = tk.Frame(body, bg=c["bg_window"])
        button_row.pack(fill="x", side="bottom", pady=(12, 0))
        ttk.Button(button_row, text="Close", command=dlg.destroy, style="Quiet.TButton").pack(side="right")
        ttk.Button(button_row, text="Copy as Table", command=copy_to_clipboard, style="Quiet.TButton").pack(
            side="right", padx=(0, 8)
        )

        table_wrap = tk.Frame(body, bg=c["bg_card"], highlightbackground=c["border"], highlightthickness=1, bd=0)
        table_wrap.pack(fill="both", expand=True)

        cols = ("sku", "description")
        tree = ttk.Treeview(table_wrap, columns=cols, show="headings", height=10)
        tree.heading("sku", text="SKU")
        tree.heading("description", text="Description")
        tree.column("sku", width=180, anchor="w", stretch=False)
        tree.column("description", width=420, anchor="w", stretch=True)
        for item in created:
            tree.insert("", "end", values=(item.get("sku", ""), item.get("description", "")))
        sb = ttk.Scrollbar(table_wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        dlg.update_idletasks()
        try:
            self.root.update_idletasks()
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            dw, dh = dlg.winfo_width(), dlg.winfo_height()
            dlg.geometry(f"+{rx + (rw - dw) // 2}+{ry + (rh - dh) // 2}")
        except tk.TclError:
            pass

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

    def _show_update_dialog(self, latest_version: str, notes: str, is_beta: bool = False) -> bool:
        """Present an Update Available dialog. Returns True if user wants
        to install, False otherwise.

        Updates are optional: both stable and beta dialogs offer a "Not Now"
        escape hatch alongside the install button.
        """
        c = UI
        dlg = tk.Toplevel(self.root)
        dlg.title("Beta Update Available" if is_beta else "Update Available")
        dlg.configure(bg=c["bg_window"])
        dlg.geometry("540x460")
        try:
            dlg.minsize(460, 380)
        except tk.TclError:
            pass
        dlg.transient(self.root)
        dlg.grab_set()

        decision = {"install": False}

        strip_color = c["warning"] if is_beta else c["accent"]
        accent_strip = tk.Frame(dlg, bg=strip_color, height=3)
        accent_strip.pack(fill="x", side="top")

        body = tk.Frame(dlg, bg=c["bg_window"])
        body.pack(fill="both", expand=True, padx=22, pady=(18, 14))

        heading_color = c["warning"] if is_beta else c["accent"]
        title_text = (
            f"Beta v{latest_version} available"
            if is_beta
            else f"Version {latest_version} is available"
        )
        tk.Label(
            body,
            text=title_text,
            bg=c["bg_window"],
            fg=heading_color,
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
        ).pack(fill="x", pady=(2, 4))
        if is_beta:
            tk.Label(
                body,
                text=(
                    "This is an experimental build that may have rough edges. "
                    "Use only if you're OK testing new features."
                ),
                bg=c["bg_window"],
                fg=c["warning"],
                font=("Segoe UI", 10),
                anchor="w",
                wraplength=480,
                justify="left",
            ).pack(fill="x", pady=(0, 12))
        else:
            tk.Frame(body, bg=c["bg_window"], height=8).pack(fill="x")

        def install():
            decision["install"] = True
            dlg.destroy()

        def cancel():
            decision["install"] = False
            dlg.destroy()

        # CRITICAL: pack the button row (side=bottom) BEFORE the expanding
        # notes box. Tk allocates space in pack order, so if the notes box
        # (expand=True) is packed first it claims the whole cavity and the
        # buttons get squeezed off the bottom — which is exactly the bug
        # where "Install / Not Now" weren't visible. Reserving the bottom
        # first guarantees the buttons always show; the notes fill what's left.
        button_row = tk.Frame(body, bg=c["bg_window"])
        button_row.pack(fill="x", side="bottom", pady=(10, 0))
        ttk.Button(button_row, text="Not Now", command=cancel, style="Quiet.TButton").pack(side="right")
        install_label = "Install Beta" if is_beta else "Install Update"
        ttk.Button(
            button_row, text=install_label, command=install, style="Primary.TButton",
        ).pack(side="right", padx=(0, 10))

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
            text_wrap.pack(fill="both", expand=True, pady=(4, 0))
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
                height=8,
                width=10,
            )
            sb = ttk.Scrollbar(text_wrap, orient="vertical", command=notes_text.yview)
            notes_text.configure(yscrollcommand=sb.set)
            notes_text.insert("1.0", cleaned)
            notes_text.configure(state="disabled")
            notes_text.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

        dlg.protocol("WM_DELETE_WINDOW", cancel)
        dlg.bind("<Return>", lambda e: install())
        dlg.bind("<Escape>", lambda e: cancel())
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

    def _geometry_fits_screen(self, geom: str, screen_w: int, screen_h: int) -> bool:
        """Return True when a saved geometry string ("WxH+X+Y") still lands
        on the current monitor. Guards against a window restoring off-
        screen if the user resized on a now-disconnected second display."""
        try:
            size_part, _, pos_part = geom.partition("+")
            if "x" not in size_part:
                return False
            w_s, h_s = size_part.split("x", 1)
            w_val, h_val = int(w_s), int(h_s)
            if pos_part:
                xs, _, ys = pos_part.partition("+")
                x_val, y_val = int(xs), int(ys) if ys else 0
            else:
                x_val, y_val = 0, 0
            if w_val < 600 or h_val < 400:
                return False
            if x_val < -50 or y_val < -50:
                return False
            if x_val + 100 > screen_w or y_val + 100 > screen_h:
                return False
            if w_val > screen_w + 100 or h_val > screen_h + 100:
                return False
            return True
        except Exception:
            return False

    def _on_close(self) -> None:
        """Save the window's current geometry before tearing down so the
        next launch can reopen at the same size + position."""
        try:
            geom = self.root.winfo_geometry()
        except tk.TclError:
            geom = ""
        if geom:
            self.settings["window_geometry"] = geom
            try:
                self._persist_settings()
            except Exception:
                log.exception("_on_close: failed to persist window geometry")
        self.root.destroy()

    def _persist_settings(self) -> None:
        payload = {
            "window_geometry": self.settings.get("window_geometry", ""),
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
            "sales_tax_item": self.sales_tax_item_var.get().strip(),
            "room_grouping_enabled": bool(self.room_grouping_var.get()),
            "auto_connect_on_startup": self.auto_connect_on_startup,
            "pricing_mode": self.pricing_mode,
            "use_actual_cost": self.use_actual_cost,
            "default_pricing_value": self.default_pricing_value,
            "brand_values": self.brand_values,
            "item_values": self.item_values,
            "line_values": self.line_values,
            "vendor_notes": self.vendor_notes,
            "brand_aliases": self.brand_aliases,
            "deleted_brands": sorted(self.deleted_brands),
        }
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _build_layout(self) -> None:
        root = self.root
        c = UI

        # Beta builds get an orange strip up top so the build channel is
        # visible at a glance; stable builds keep the blue accent.
        strip_color = c["warning"] if IS_BETA else c["accent"]
        accent_strip = tk.Frame(root, bg=strip_color, height=4)
        accent_strip.pack(fill="x", side="top")

        # Body splits into sidebar (left, two-tone) and main content (right).
        # Status bar (packed below) and accent strip (packed above) frame
        # this body section vertically.
        body = tk.Frame(root, bg=c["bg_window"], highlightthickness=0, bd=0)
        body.pack(fill="both", expand=True)

        sidebar = tk.Frame(body, bg=c["bg_sidebar"], width=225, highlightthickness=0, bd=0)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        sidebar_border = tk.Frame(body, bg=c["border"], width=1)
        sidebar_border.pack(side="left", fill="y")
        self._build_sidebar(sidebar)

        # Wrap everything that used to be at root level in `main_content`.
        main_content = tk.Frame(body, bg=c["bg_window"], highlightthickness=0, bd=0)
        main_content.pack(side="left", fill="both", expand=True)

        # A slimmer header in the main area now -- the brand block lives in
        # the sidebar so the top of the content can just show the subtitle.
        header = ttk.Frame(main_content, padding=(24, 16, 24, 8))
        header.pack(fill="x")
        ttk.Label(
            header,
            text="Upload source data, review mapped rows, then export or upload to QuickBooks Desktop.",
            style="SubHeader.TLabel",
        ).pack(anchor="w")

        config = ttk.LabelFrame(main_content, text="  Configuration  ", style="Card.TLabelframe")
        config.pack(fill="x", padx=24, pady=(4, 8))
        # Two columns: form fields on the left, validation messages on the
        # right. v1.017b shifts the ratio to 3:1 so the form has enough
        # room and Validation Messages still gets a comfortable column.
        config.columnconfigure(0, weight=3, uniform="cfg")
        config.columnconfigure(1, weight=1, uniform="cfg")

        config_left = ttk.Frame(config, style="Card.TFrame")
        config_left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        config_right = ttk.Frame(config, style="Card.TFrame")
        config_right.grid(row=0, column=1, sticky="nsew")
        config_right.rowconfigure(1, weight=1)
        config_right.columnconfigure(0, weight=1)
        ttk.Label(
            config_right,
            text="Validation Messages",
            style="FieldLabel.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=4, pady=(2, 4))
        error_wrap = tk.Frame(
            config_right,
            bg=c["bg_card"],
            highlightbackground=c["border"],
            highlightthickness=1,
            bd=0,
        )
        error_wrap.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        self.error_text = tk.Text(
            error_wrap,
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
            padx=6,
            pady=6,
            height=6,
            # Explicit small width so the Text widget doesn't claim the
            # default 80 characters and push the right column past the
            # config card's allocation. The fill="both"/expand=True on the
            # pack still stretches it to fill whatever space is granted.
            width=10,
        )
        self.error_text.pack(fill="both", expand=True)

        # v1.039 layout: same unified form grid, but each functional block
        # gets a "STEP N — …" eyebrow + title above it so first-time users
        # see the workflow without having to ask. Step 1 sits over the
        # Source File column; Step 2 spans the customer/SO/date/config
        # cluster. Steps 3 and 4 live below the form on the action row.
        #
        # Grid rows (shifted down by 2 vs v1.038 to make room for the
        # step-header row at the top):
        #   Row 0/1: STEP 1 / STEP 2 eyebrow + title headers
        #   Row 2/3: Source File [Browse] | Customer | Sales Order No | Fetch
        #   Row 4/5: Sales Order Date [📅] | Due Date [📅]
        #   Row 7/8: Terms | Shipping Method | Currency | Tax Code | Default Income Account
        form = ttk.Frame(config_left, style="Card.TFrame")
        form.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        FORM_COLS = 16
        for i in range(FORM_COLS):
            form.columnconfigure(i, weight=1)

        # --- Row 0/1: STEP headers ---
        ttk.Label(form, text="STEP 1", style="StepEyebrow.TLabel").grid(
            row=0, column=0, columnspan=5, sticky="w", padx=4, pady=(2, 0)
        )
        ttk.Label(
            form, text="Pick the source quote file", style="StepTitle.TLabel"
        ).grid(row=1, column=0, columnspan=5, sticky="w", padx=4, pady=(0, 6))
        ttk.Label(form, text="STEP 2", style="StepEyebrow.TLabel").grid(
            row=0, column=5, columnspan=11, sticky="w", padx=4, pady=(2, 0)
        )
        ttk.Label(
            form,
            text="Fill in the customer & order details",
            style="StepTitle.TLabel",
        ).grid(row=1, column=5, columnspan=11, sticky="w", padx=4, pady=(0, 6))

        # --- Row 2/3: Source File, Customer, Sales Order No, Fetch ---
        # Reshuffled v1.033 to make Customer noticeably longer. Spans
        # now sum to 16 as: Source 5 | Customer 5 | SO No 3 | Fetch 3.
        src_label = ttk.Label(form, text="Source File", style="FieldLabel.TLabel")
        src_label.grid(row=2, column=0, columnspan=5, sticky="w", padx=4, pady=(4, 1))
        src_row = ttk.Frame(form, style="Card.TFrame")
        src_row.grid(row=3, column=0, columnspan=5, sticky="ew", padx=4, pady=(0, 4))
        ttk.Button(
            src_row, text="Browse", command=self._browse_source, style="Quiet.TButton"
        ).pack(side="right", padx=(6, 0))
        # width=20 — long enough to show ~20 chars of a path before
        # truncation kicks in. sticky/fill still lets it grow.
        ttk.Entry(src_row, textvariable=self.source_path_var, width=20).pack(
            side="left", fill="x", expand=True
        )

        customer_entry = self._form_entry(
            form, "Customer", self.customer_var, 2, 5, 5, min_chars=24
        )
        so_no_entry = self._form_entry(
            form, "Sales Order No", self.sales_order_no_var, 2, 10, 3, min_chars=8
        )
        ttk.Button(
            form,
            text="Fetch Next",
            command=self.fetch_next_so,
            style="Quiet.TButton",
        ).grid(row=3, column=13, columnspan=3, padx=4, sticky="ew", pady=(0, 4))

        # --- Row 4/5: Sales Order Date, Due Date ---
        so_date_entry = self._date_entry(
            form, "Sales Order Date", self.sales_order_date_var, 4, 0, 6, min_chars=11
        )
        due_date_entry = self._date_entry(
            form, "Due Date", self.due_date_var, 4, 6, 6, min_chars=11
        )

        # Required fields used by both Build Preview and Upload. Each entry
        # gets a write trace on its var so typing clears the red error
        # state immediately.
        self._required_fields = [
            ("Customer", self.customer_var, customer_entry),
            ("Sales Order No", self.sales_order_no_var, so_no_entry),
            ("Sales Order Date", self.sales_order_date_var, so_date_entry),
            ("Due Date", self.due_date_var, due_date_entry),
        ]
        for _, var, entry in self._required_fields:
            var.trace_add("write", lambda *_a, e=entry: self._clear_field_error(e))

        # --- Row 7/8: Terms, Shipping Method, Currency, Tax Code, Income Account ---
        # (Row 6 is a small visual gap before the bottom config strip.)
        self._form_entry(form, "Terms", self.terms_var, 7, 0, 2, min_chars=10)
        self._form_entry(form, "Shipping Method", self.shipping_method_var, 7, 2, 4, min_chars=14)
        self._form_entry(form, "Currency", self.currency_var, 7, 6, 2, min_chars=6)
        self._form_entry(form, "Tax Code", self.tax_code_var, 7, 8, 2, min_chars=6)
        self._form_entry(form, "Default Income Account", self.income_account_var, 7, 10, 6, min_chars=18)

        # Bottom bar of the Configuration card — now just the QB status pill,
        # since Connect/Admin Setup/Update commands moved to the Setup and
        # Help menus in v1.013b.
        qb_bar = ttk.Frame(config_left, style="Card.TFrame")
        qb_bar.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        self.qb_status_pill = tk.Frame(
            qb_bar,
            bg=c["bg_hover"],
            highlightbackground=c["bg_hover"],
            highlightthickness=0,
            bd=0,
        )
        self.qb_status_pill.pack(side="left")
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

        # v1.039 action row: each primary action gets its own "STEP N" badge
        # stacked above the button so the workflow reads top-to-bottom even
        # though the buttons themselves are laid out left-to-right. Export
        # for SaaSant stays alongside Build Preview under Step 3 — it's an
        # alternate format of "preview output," not a separate step.
        actions = ttk.Frame(main_content, padding=(24, 4, 24, 10))
        actions.pack(fill="x")

        step3 = ttk.Frame(actions, style="TFrame")
        step3.pack(side="left")
        ttk.Label(step3, text="STEP 3", style="StepEyebrowOnWindow.TLabel").pack(
            anchor="w", padx=2
        )
        ttk.Label(
            step3,
            text="Build the preview & verify the lines",
            style="StepTitleOnWindow.TLabel",
        ).pack(anchor="w", padx=2, pady=(0, 4))
        step3_buttons = ttk.Frame(step3, style="TFrame")
        step3_buttons.pack(anchor="w")
        ttk.Button(
            step3_buttons,
            text="Build Preview",
            command=self.build_preview,
            style="Primary.TButton",
        ).pack(side="left")
        ttk.Button(
            step3_buttons,
            text="Export for SaaSant",
            command=self.export_saasant_template,
            style="Accent.TButton",
        ).pack(side="left", padx=(8, 0))

        step4 = ttk.Frame(actions, style="TFrame")
        step4.pack(side="left", padx=(28, 0))
        ttk.Label(step4, text="STEP 4", style="StepEyebrowOnWindow.TLabel").pack(
            anchor="w", padx=2
        )
        ttk.Label(
            step4,
            text="Upload the sales order to QuickBooks",
            style="StepTitleOnWindow.TLabel",
        ).pack(anchor="w", padx=2, pady=(0, 4))
        ttk.Button(
            step4,
            text="Upload to QuickBooks",
            command=self.upload_to_quickbooks,
            style="Success.TButton",
        ).pack(anchor="w")

        content = ttk.Panedwindow(main_content, orient="vertical")
        content.pack(fill="both", expand=True, padx=24, pady=(0, 8))

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

        # Validation Messages now lives on the right side of the
        # Configuration card (see config_right above). The old bottom panel
        # was removed in v1.014b.

        status_bar = tk.Frame(root, bg=c["bg_window"], highlightthickness=0, bd=0)
        status_bar.pack(fill="x", side="bottom")
        separator = tk.Frame(status_bar, bg=c["border"], height=1)
        separator.pack(fill="x", side="top")
        ttk.Label(status_bar, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w")

    def _build_sidebar(self, sidebar: tk.Frame) -> None:
        """DMQuotes brand block. Loads the actual logo PNG bundled with
        the app (which contains the document/arrow/ring graphic plus
        the wordmark and tagline) and pins Shimiralabs + version at the
        bottom.
        """
        c = UI

        # Top spacer so the logo doesn't kiss the accent strip.
        tk.Frame(sidebar, bg=c["bg_sidebar"], height=18).pack(side="top")

        logo_path = _resource_path("Logo_New.png")
        self._sidebar_logo_image = None
        loaded = False
        if logo_path.exists():
            try:
                from PIL import Image, ImageTk
                pil_img = Image.open(str(logo_path))
                # Resize to fit comfortably in the sidebar width.
                target = 188
                pil_img = pil_img.convert("RGBA")
                pil_img.thumbnail((target, target), Image.LANCZOS)
                # v1.036: Logo_New.png ships with a proper alpha channel, so
                # it composites cleanly onto the dark-green sidebar with no
                # chroma-key trick required.
                self._sidebar_logo_image = ImageTk.PhotoImage(pil_img)
                tk.Label(
                    sidebar,
                    image=self._sidebar_logo_image,
                    bg=c["bg_sidebar"],
                    borderwidth=0,
                ).pack(side="top")
                loaded = True
            except Exception:
                log.exception("_build_sidebar: failed to load logo via PIL; falling back to tk.PhotoImage")
                try:
                    self._sidebar_logo_image = tk.PhotoImage(file=str(logo_path))
                    # tk.PhotoImage can only scale via integer subsample.
                    # Original is ~1254 px; subsample by 8 -> ~157 px.
                    self._sidebar_logo_image = self._sidebar_logo_image.subsample(8, 8)
                    tk.Label(
                        sidebar,
                        image=self._sidebar_logo_image,
                        bg=c["bg_sidebar"],
                        borderwidth=0,
                    ).pack(side="top")
                    loaded = True
                except Exception:
                    log.exception("_build_sidebar: tk.PhotoImage fallback also failed")

        if not loaded:
            # Final fallback: plain text wordmark so the sidebar isn't blank.
            # v1.035: sidebar is now dark green, so the wordmark flips to
            # white + bright QB-green for contrast.
            wordmark = tk.Frame(sidebar, bg=c["bg_sidebar"])
            wordmark.pack(side="top", pady=(20, 0))
            tk.Label(
                wordmark, text="DM",
                bg=c["bg_sidebar"], fg=c["sidebar_text"],
                font=("Segoe UI Black", 22),
            ).pack(side="left")
            tk.Label(
                wordmark, text="Quotes",
                bg=c["bg_sidebar"], fg=c["accent_light"],
                font=("Segoe UI Black", 22),
            ).pack(side="left")
            tk.Label(
                sidebar, text="INTO QUICKBOOKS",
                bg=c["bg_sidebar"], fg=c["sidebar_text_muted"],
                font=("Segoe UI Semibold", 9),
            ).pack(side="top", pady=(4, 0))

        # Version pinned near the bottom — muted light on dark green.
        tk.Label(
            sidebar, text=APP_VERSION,
            bg=c["bg_sidebar"], fg=c["sidebar_text_muted"],
            font=("Segoe UI", 9),
        ).pack(side="bottom", pady=(0, 12))

        # Subtle "made by" line above the version.
        tk.Label(
            sidebar, text="Shimiralabs",
            bg=c["bg_sidebar"], fg=c["sidebar_text_muted"],
            font=("Segoe UI", 8),
        ).pack(side="bottom")

    def _draw_rounded_rect(self, canvas: tk.Canvas, x1: int, y1: int, x2: int, y2: int, radius: int, fill: str) -> None:
        """Approximate a filled rounded rectangle on a Tk canvas using four
        corner arcs plus two overlapping rectangles."""
        d = radius * 2
        canvas.create_arc(x1, y1, x1 + d, y1 + d, start=90, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x2 - d, y1, x2, y1 + d, start=0, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x1, y2 - d, x1 + d, y2, start=180, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x2 - d, y2 - d, x2, y2, start=270, extent=90, fill=fill, outline=fill)
        canvas.create_rectangle(x1 + radius, y1, x2 - radius, y2, fill=fill, outline=fill)
        canvas.create_rectangle(x1, y1 + radius, x2, y2 - radius, fill=fill, outline=fill)

    def _clear_field_error(self, entry) -> None:
        """Reset an entry's style back to default TEntry. Safe to call on
        an already-default entry."""
        try:
            entry.configure(style="TEntry")
        except tk.TclError:
            pass

    def _validate_required_fields(self) -> list[str]:
        """Mark blank required entries in red and return the list of label
        names that failed. Empty list = all good."""
        missing: list[str] = []
        for label, var, entry in getattr(self, "_required_fields", []):
            value = (var.get() or "").strip()
            if value:
                self._clear_field_error(entry)
            else:
                try:
                    entry.configure(style="Error.TEntry")
                except tk.TclError:
                    pass
                missing.append(label)
        return missing

    def _path_row(self, parent, label, var, browse_cmd, row_idx):
        ttk.Label(parent, text=label, style="Card.TLabel", width=18).grid(
            row=row_idx, column=0, sticky="w", pady=(2, 2)
        )
        ttk.Entry(parent, textvariable=var).grid(row=row_idx, column=1, sticky="ew", padx=10, pady=(2, 2))
        ttk.Button(parent, text="Browse", command=browse_cmd).grid(
            row=row_idx, column=2, sticky="e", pady=(2, 2)
        )
        parent.columnconfigure(1, weight=1)

    def _path_row_inline(self, parent, label, var, browse_cmd):
        """Compact path row: label on top, entry + Browse button below.

        Uses pack so the Browse button always claims its natural width
        first (right side) and the Entry expands to fill whatever's left.
        Setting width=1 on the Entry removes its 20-character minimum so
        the row shrinks gracefully when the parent is narrow.
        """
        ttk.Label(parent, text=label, style="FieldLabel.TLabel").pack(
            anchor="w", padx=4, pady=(2, 1)
        )
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x", padx=4, pady=(0, 2))
        ttk.Button(
            row, text="Browse", command=browse_cmd, style="Quiet.TButton"
        ).pack(side="right", padx=(6, 0))
        ttk.Entry(row, textvariable=var, width=1).pack(
            side="left", fill="x", expand=True
        )

    def _form_entry(self, parent, label, var, row, col, span, min_chars=10):
        """Grid a labelled text entry. `min_chars` is the entry's minimum
        character width — sticky="ew" lets it grow when the cell is wider,
        but won't let it shrink below this floor. Stops fields from
        collapsing to "Pr..." / "C:/Us..." on a narrow window."""
        ttk.Label(parent, text=label, style="FieldLabel.TLabel").grid(
            row=row, column=col, sticky="w", padx=4, pady=(4, 1), columnspan=span
        )
        entry = ttk.Entry(parent, textvariable=var, width=min_chars)
        entry.grid(
            row=row + 1, column=col, columnspan=span, sticky="ew", padx=4, pady=(0, 4)
        )
        return entry

    def _date_entry(self, parent, label, var, row, col, span, min_chars=12):
        """Form entry with a small 📅 picker button on the right. The
        picker packs first (right) so it always reserves its width;
        the entry packs second (left) and grows to fill the rest, but
        won't shrink narrower than `min_chars` characters."""
        ttk.Label(parent, text=label, style="FieldLabel.TLabel").grid(
            row=row, column=col, sticky="w", padx=4, pady=(4, 1), columnspan=span
        )
        wrap = ttk.Frame(parent, style="Card.TFrame")
        wrap.grid(row=row + 1, column=col, columnspan=span, sticky="ew", padx=4, pady=(0, 4))
        ttk.Button(
            wrap,
            text="📅",
            width=3,
            command=lambda v=var: self._open_date_picker(v),
            style="Quiet.TButton",
        ).pack(side="right", padx=(4, 0))
        entry = ttk.Entry(wrap, textvariable=var, width=min_chars)
        entry.pack(side="left", fill="x", expand=True)
        return entry

    def _open_date_picker(self, var: tk.StringVar) -> None:
        """Tiny calendar popup. Reads the current value to seed the month,
        renders day buttons, and on click writes MM/DD/YYYY back into var."""
        c = UI
        try:
            seed = datetime.strptime(var.get().strip(), "%m/%d/%Y")
        except Exception:
            seed = datetime.now()
        state = {"year": seed.year, "month": seed.month}

        dlg = tk.Toplevel(self.root)
        dlg.title("Pick a date")
        dlg.configure(bg=c["bg_window"])
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        accent = tk.Frame(dlg, bg=c["accent"], height=3)
        accent.pack(fill="x")

        outer = tk.Frame(dlg, bg=c["bg_window"])
        outer.pack(padx=14, pady=12)

        nav = tk.Frame(outer, bg=c["bg_window"])
        nav.pack(fill="x")
        title_var = tk.StringVar()
        title_label = tk.Label(
            nav,
            textvariable=title_var,
            bg=c["bg_window"],
            fg=c["text_primary"],
            font=("Segoe UI Semibold", 12),
            width=18,
        )

        grid_frame = tk.Frame(outer, bg=c["bg_window"])
        grid_frame.pack(pady=(8, 0))

        def render():
            import calendar
            for child in grid_frame.winfo_children():
                child.destroy()
            year, month = state["year"], state["month"]
            title_var.set(datetime(year, month, 1).strftime("%B %Y"))
            for i, name in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
                tk.Label(
                    grid_frame, text=name, bg=c["bg_window"], fg=c["text_secondary"],
                    font=("Segoe UI", 9), width=4,
                ).grid(row=0, column=i, padx=1, pady=1)
            cal = calendar.Calendar(firstweekday=0)
            today = datetime.now().date()
            for r, week in enumerate(cal.monthdatescalendar(year, month), start=1):
                for c_idx, day in enumerate(week):
                    in_month = (day.month == month)
                    is_today = (day == today)
                    label_fg = c["text_primary"] if in_month else c["text_tertiary"]
                    bg = c["accent_bg"] if is_today and in_month else c["bg_card"]
                    btn = tk.Button(
                        grid_frame,
                        text=str(day.day),
                        width=4,
                        bg=bg,
                        fg=label_fg,
                        relief="flat",
                        bd=0,
                        font=("Segoe UI", 9, "bold" if is_today else "normal"),
                        activebackground=c["accent"],
                        activeforeground="#FFFFFF",
                        command=lambda d=day: pick(d),
                    )
                    btn.grid(row=r, column=c_idx, padx=1, pady=1, sticky="nsew")

        def pick(day):
            var.set(day.strftime("%m/%d/%Y"))
            dlg.destroy()

        def shift(delta):
            m = state["month"] + delta
            y = state["year"]
            while m < 1:
                m += 12; y -= 1
            while m > 12:
                m -= 12; y += 1
            state["month"] = m; state["year"] = y
            render()

        ttk.Button(nav, text="◀", width=3, command=lambda: shift(-1), style="Quiet.TButton").pack(side="left")
        title_label.pack(side="left", padx=8)
        ttk.Button(nav, text="▶", width=3, command=lambda: shift(1), style="Quiet.TButton").pack(side="left")
        ttk.Button(nav, text="Today", command=lambda: pick(datetime.now().date()), style="Quiet.TButton").pack(
            side="right"
        )

        render()

        dlg.update_idletasks()
        try:
            self.root.update_idletasks()
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            dw, dh = dlg.winfo_width(), dlg.winfo_height()
            dlg.geometry(f"+{rx + (rw - dw) // 2}+{ry + (rh - dh) // 2}")
        except tk.TclError:
            pass

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
        # PDF showroom-quote import is a beta-only feature for now; stable
        # builds only offer Excel.
        if IS_BETA:
            filetypes = [
                ("Quote files", "*.xls *.xlsx *.pdf"),
                ("Excel Files", "*.xls *.xlsx"),
                ("Showroom PDF (Deluxe Vanity)", "*.pdf"),
                ("All Files", "*.*"),
            ]
        else:
            filetypes = [("Excel Files", "*.xls *.xlsx")]
        path = filedialog.askopenfilename(filetypes=filetypes)
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

    def _not_connected_message(self) -> str:
        """Status pill text when QuickBooks isn't connected. Points users at
        the Setup menu so they know where the Connect action moved to."""
        return "Go to Setup → Connect to QuickBooks Desktop to connect"

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
        log.info("Auto-connect on startup (background)")
        self._set_qb_status("Connecting...", state="pending")
        threading.Thread(target=self._silent_connect_worker, daemon=True).start()

    def _silent_connect_worker(self):
        """Run the auto-connect QuickBooks attempt off the main thread so
        the UI stays responsive during the SDK call (which can take 5+
        seconds). UI updates are marshaled back via root.after."""
        try:
            company_name = self._qb_client().test_connection()
        except Exception as exc:
            log.warning("Auto-connect: silent attempt failed — %s", exc)
            self.root.after(0, self._on_silent_connect_failed)
            return
        log.info("Auto-connect: silent success, company=%r", company_name)
        self.root.after(0, lambda n=company_name: self._on_silent_connect_success(n))

    def _on_silent_connect_success(self, company_name: str) -> None:
        self._set_qb_status(f"Connected: {company_name}", state="connected")
        self._set_status(f"Connected to QuickBooks company: {company_name}")
        if not self.auto_connect_on_startup:
            log.info("Auto-connect: enabling auto-connect on future startups")
            self.auto_connect_on_startup = True
            try:
                self._persist_settings()
            except Exception:
                log.exception("Auto-connect: failed to persist auto-connect flag")

    def _on_silent_connect_failed(self) -> None:
        self._set_qb_status(self._not_connected_message(), state="disconnected")
        self._set_status("QuickBooks auto-connect failed. Use Setup → Connect to QuickBooks Desktop to retry.")

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
            self._set_qb_status(self._not_connected_message(), state="disconnected")
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

    def _load_bundled_vendor_pricing(self) -> None:
        """Load the vendor price database that ships with the app so brand
        multipliers work out of the box — no manual import required. Single
        multipliers populate self.vendor_clean; tiered/textual entries are
        merged into self.vendor_notes (without clobbering user-set notes)."""
        try:
            path = _resource_path("vendor_pricing.json")
            if not path.exists():
                log.info("_load_bundled_vendor_pricing: no bundled vendor_pricing.json")
                return
            data = json.loads(path.read_text(encoding="utf-8"))
            clean = data.get("clean", {}) or {}
            tiered = data.get("tiered", {}) or {}
            # Honor brands the user explicitly deleted — don't resurrect them
            # from the bundle on the next launch.
            deleted = self.deleted_brands
            self.vendor_clean = {
                str(k): float(v) for k, v in clean.items() if str(k) not in deleted
            }
            for brand, note in tiered.items():
                if str(brand) in deleted:
                    continue
                self.vendor_notes.setdefault(str(brand), str(note))
            log.info(
                "_load_bundled_vendor_pricing: loaded %d clean, %d tiered brands",
                len(self.vendor_clean), len(tiered),
            )
        except Exception:
            log.exception("_load_bundled_vendor_pricing failed")

    def _known_brands(self) -> set:
        """Brands we already have a usable multiplier for and so won't prompt:
        the bundled baseline, anything the user saved, plus any brand aliased
        to one of those."""
        direct = set(self.vendor_clean) | set(self.brand_values) | set(self.session_brand_values)
        for src, tgt in self.brand_aliases.items():
            if tgt in direct:
                direct.add(src)
        return direct

    def _effective_brand_values(self) -> dict:
        """Resolve the multiplier each brand should use for THIS order.
        Precedence: per-order 'variable' values > user-saved overrides >
        bundled vendor baseline. Aliased brands inherit their target's
        multiplier."""
        merged = dict(self.vendor_clean)
        merged.update(self.brand_values)
        merged.update(self.session_brand_values)
        for src, tgt in self.brand_aliases.items():
            if tgt in merged and src not in merged:
                merged[src] = merged[tgt]
        return merged

    def _load_source_file(self, path: str):
        """Load a source quote, gating PDF import to beta builds only."""
        if path.lower().endswith(".pdf") and not IS_BETA:
            raise RuntimeError(
                "Reading showroom PDF quotes is currently a beta-only feature.\n\n"
                "Switch to the beta build (Setup → Check for Beta Update) to "
                "import Deluxe Vanity & Kitchen PDFs."
            )
        return load_source(path)

    def _ensure_source_loaded(self):
        if self.source_df is not None:
            return
        self.source_df = self._load_source_file(self.source_path_var.get().strip())

    def import_vendor_price_list(self):
        """Build the brand→multiplier database from the vendor price list
        Excel. Column F (Cost Multiplier) drives it: brands with a single
        clean number are stored as their multiplier; brands with tiered or
        textual pricing (multiple lines) are recorded as notes so the user is
        asked to pick the right multiplier when that brand comes up — we never
        silently guess a default for them."""
        path = filedialog.askopenfilename(
            title="Select the Vendor Price List",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            clean, ambiguous = load_vendor_multipliers(path)
        except Exception as exc:
            log.exception("import_vendor_price_list: failed to read %s", path)
            messagebox.showerror(
                "Vendor Price List",
                f"Couldn't read the vendor price list:\n\n{exc}\n\n"
                "Make sure the first tab has Brand in column A and Cost "
                "Multiplier in column F.",
            )
            return

        if not clean and not ambiguous:
            messagebox.showwarning(
                "Vendor Price List",
                "No brands with multipliers were found on the first tab. "
                "Check that Brand is in column A and Cost Multiplier in column F.",
            )
            return

        # Single-value multipliers become the brand database. Merge so any
        # multipliers the user typed by hand for brands not in this sheet
        # survive.
        self.brand_values.update(clean)
        # Tiered/textual brands: store the note, and make sure they are NOT in
        # brand_values so they get prompted (with the note shown) each time.
        for brand, note in ambiguous.items():
            self.vendor_notes[brand] = note
            self.brand_values.pop(brand, None)
        # From now on pricing is per-brand from the vendor list, so switch the
        # mode to brand.
        self.pricing_mode = "brand"
        self.use_actual_cost = False
        self._persist_settings()

        log.info(
            "Imported vendor price list: %d clean, %d tiered/ambiguous from %s",
            len(clean), len(ambiguous), path,
        )
        self._set_status(
            f"Vendor price list imported: {len(clean)} brand multipliers, "
            f"{len(ambiguous)} need a choice when used."
        )
        sample = ", ".join(sorted(ambiguous)[:8])
        messagebox.showinfo(
            "Vendor Price List Imported",
            f"Loaded {len(clean)} brand multipliers from column F.\n\n"
            f"{len(ambiguous)} brand(s) have tiered or text pricing (multiple "
            f"lines in the cell). You'll be asked to enter the multiplier for "
            f"those when a quote uses them, with the vendor note shown to help "
            f"you pick.\n\n"
            + (f"Tiered brands include: {sample}{'…' if len(ambiguous) > 8 else ''}" if ambiguous else ""),
        )

    def change_pricing_rules(self):
        try:
            # v1.040: don't hard-require a complete order (_validate_settings)
            # or a loaded source file. Reviewing/editing saved pricing should
            # work anytime — this was the bug that hid previously-saved brands
            # on a machine that didn't have a source file + full order entered.
            # If a valid source file IS set, load it so the dialog can also
            # list that file's brands; otherwise just show the saved ones.
            try:
                src = self.source_path_var.get().strip()
                if src and Path(src).exists():
                    self._ensure_source_loaded()
            except Exception:
                log.info("change_pricing_rules: no source file loaded; showing saved pricing only")
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

            pricing_mode, use_actual_cost, default_value, values, deleted, combines = dlg.result
            self.pricing_mode = pricing_mode
            self.use_actual_cost = use_actual_cost
            self.default_pricing_value = default_value
            # v1.040: merge (update) rather than replace so multipliers for
            # brands/SKUs that weren't shown this time (e.g. they live in a
            # different source file) are preserved instead of wiped.
            if pricing_mode == "brand":
                self.brand_values.update(values)
                # Combines: alias source -> target, drop the source's own data.
                for src, tgt in combines.items():
                    self.brand_aliases[src] = tgt
                    self.brand_values.pop(src, None)
                    self.session_brand_values.pop(src, None)
                    self.vendor_clean.pop(src, None)
                # Deletes: remove everywhere and remember so the bundle won't
                # bring them back next launch.
                for key in deleted:
                    self.deleted_brands.add(key)
                    self.brand_values.pop(key, None)
                    self.session_brand_values.pop(key, None)
                    self.vendor_clean.pop(key, None)
                    self.vendor_notes.pop(key, None)
                    self.brand_aliases.pop(key, None)
            elif pricing_mode == "item":
                self.item_values.update(values)
                for key in deleted:
                    self.item_values.pop(key, None)
            else:
                self.line_values.update(values)
                for key in deleted:
                    self.line_values.pop(key, None)
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
                # In brand mode, fold the bundled vendor baseline + any
                # per-order "variable" values into the multiplier lookup.
                self._effective_brand_values()
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

        # Show the time-savings estimate in the validation panel as soon as
        # the preview is ready — appended below any validation messages — so
        # the payoff is visible before the user even uploads.
        try:
            records = self.output_df.to_dict(orient="records")
            num_lines = len(records)
            num_rooms = _count_room_groups(records)
            est_lines = self._time_saved_message_lines(num_lines, num_rooms, None)
            if est_lines:
                self.error_text.insert(tk.END, "\n\n" + "\n".join(est_lines))
        except Exception:
            log.exception("post-preview time-saved estimate failed; skipping")

    def build_preview(self):
        try:
            missing = self._validate_required_fields()
            if missing:
                messagebox.showwarning(
                    "Required fields are blank",
                    "Please fill in the following before building the preview:\n\n  - "
                    + "\n  - ".join(missing),
                )
                return
            self._validate_settings()
            source_path = self.source_path_var.get().strip()
            # A new source file = a new order, so clear any per-order
            # "variable" multipliers from a previous order.
            if source_path != self._session_source_path:
                self.session_brand_values = {}
                self._session_source_path = source_path
            self.source_df = self._load_source_file(source_path)
            self.output_overrides = {}

            # When pricing by brand, prompt for any brand we don't yet have a
            # multiplier for. "Known" includes the bundled vendor database, so
            # the 147 single-multiplier brands never prompt. Tiered brands
            # (and anything not in the list) are asked — variable by default,
            # with the option to lock the cost.
            if self.pricing_mode == "brand":
                source_brands = unique_brands(self.source_df)
                known = self._known_brands()
                missing = [b for b in source_brands if b and b not in known]
                if missing:
                    cancelled = self._prompt_missing_brand_multipliers(missing)
                    if cancelled:
                        return

            self._rebuild_previews_from_source()
            messagebox.showinfo("Preview Ready", f"Generated {len(self.output_df)} template rows.")
        except Exception as exc:
            messagebox.showerror("Build Preview Error", str(exc))

    def _prompt_missing_brand_multipliers(self, missing_brands: list[str]) -> bool:
        """Show a dialog listing brands not yet in self.brand_values and
        ask the user to enter a multiplier (or actual cost) for each. The
        entered values are written back to self.brand_values and persisted.

        Returns True when the user cancelled (preview build should abort),
        False when they entered values and saved.
        """
        c = UI
        dlg = tk.Toplevel(self.root)
        dlg.title("Brand Pricing")
        dlg.configure(bg=c["bg_window"])
        dlg.geometry("700x560")
        try:
            dlg.minsize(560, 420)
        except tk.TclError:
            pass
        dlg.transient(self.root)
        dlg.grab_set()

        outcome = {"cancelled": True}

        accent = tk.Frame(dlg, bg=c["accent"], height=3)
        accent.pack(fill="x", side="top")

        body = tk.Frame(dlg, bg=c["bg_window"])
        body.pack(fill="both", expand=True, padx=22, pady=(18, 14))

        tk.Label(
            body,
            text=f"{len(missing_brands)} brand(s) need pricing",
            bg=c["bg_window"],
            fg=c["accent"],
            font=("Segoe UI Semibold", 11),
            anchor="w",
        ).pack(fill="x")
        if self.use_actual_cost:
            sub = "Enter the actual cost/rate, or match the brand to one already in the system."
        else:
            sub = (
                "Enter the cost multiplier, or match the brand to one already in "
                "the system if it's the same vendor under a different name. Leave "
                "\"Variable\" checked to be asked each order; uncheck to lock it."
            )
        tk.Label(
            body,
            text=sub,
            bg=c["bg_window"],
            fg=c["text_secondary"],
            font=("Segoe UI", 9),
            anchor="w",
            wraplength=640,
            justify="left",
        ).pack(fill="x", pady=(3, 10))

        # Sorted list of every brand already known to the app, used as the
        # "match to existing brand" options.
        match_options = [""] + sorted(
            set(self.vendor_clean) | set(self.brand_values) | set(self.vendor_notes)
        )

        # Buttons are packed at the BOTTOM first, so the scrollable list can
        # never push them off-screen (the bug where Save wasn't visible).
        button_row = tk.Frame(body, bg=c["bg_window"])
        button_row.pack(fill="x", side="bottom", pady=(10, 0))

        wrapper = tk.Frame(body, bg=c["bg_card"], highlightbackground=c["border"], highlightthickness=1, bd=0)
        wrapper.pack(fill="both", expand=True)
        canvas = tk.Canvas(wrapper, highlightthickness=0, bg=c["bg_card"], bd=0)
        scroll = ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="Card.TFrame")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scroll.pack(side="right", fill="y", pady=8)

        entry_vars: dict[str, tk.StringVar] = {}
        variable_vars: dict[str, tk.BooleanVar] = {}
        match_vars: dict[str, tk.StringVar] = {}
        for brand in missing_brands:
            note = self.vendor_notes.get(brand)
            is_tiered = bool(note)
            row = ttk.Frame(inner, style="Card.TFrame")
            row.pack(fill="x", padx=10, pady=6)

            # Left column: brand name, vendor note, and a "match to existing
            # brand" dropdown for when the source used a wrong/alias name.
            left = ttk.Frame(row, style="Card.TFrame")
            left.pack(side="left", fill="x", expand=True)
            ttk.Label(left, text=brand, style="Card.TLabel", font=("Segoe UI Semibold", 9)).pack(anchor="w")
            if note:
                ttk.Label(
                    left, text=note, style="Card.TLabel",
                    foreground=UI["accent"], wraplength=360, justify="left",
                    font=("Segoe UI", 8),
                ).pack(anchor="w")
            match_row = ttk.Frame(left, style="Card.TFrame")
            match_row.pack(anchor="w", pady=(2, 0))
            ttk.Label(
                match_row, text="Match to existing brand:", style="Card.TLabel",
                font=("Segoe UI", 8),
            ).pack(side="left", padx=(0, 6))
            mvar = tk.StringVar(value="")
            match_vars[brand] = mvar
            ttk.Combobox(
                match_row, textvariable=mvar, values=match_options,
                state="readonly", width=26, font=("Segoe UI", 8),
            ).pack(side="left")

            # Right: multiplier entry + Variable toggle. Tiered brands default
            # to Variable (ask each order); unknown brands default to locked.
            var = tk.StringVar(value="")
            entry_vars[brand] = var
            ttk.Entry(row, textvariable=var, width=7).pack(side="right", padx=(8, 0))
            vbar = tk.BooleanVar(value=is_tiered)
            variable_vars[brand] = vbar
            ttk.Checkbutton(row, text="Variable", variable=vbar, style="TCheckbutton").pack(side="right")

        def cancel():
            outcome["cancelled"] = True
            dlg.destroy()

        def save():
            aliases: dict[str, str] = {}
            parsed: dict[str, float] = {}
            try:
                for brand in missing_brands:
                    match = match_vars[brand].get().strip()
                    if match:
                        # Mapped to an existing brand — inherit its pricing.
                        aliases[brand] = match
                        continue
                    text = entry_vars[brand].get().strip()
                    if not text:
                        raise ValueError(
                            f"{brand}: enter a multiplier or pick a brand to match it to."
                        )
                    parsed[brand] = float(text)
            except ValueError as exc:
                messagebox.showerror("Invalid Value", str(exc), parent=dlg)
                return
            # Aliases: remember the mapping so this name auto-resolves next time.
            for brand, target in aliases.items():
                self.brand_aliases[brand] = target
                self.brand_values.pop(brand, None)
                self.session_brand_values.pop(brand, None)
            # Variable -> per-order only; locked -> saved in brand_values.
            for brand, value in parsed.items():
                self.brand_aliases.pop(brand, None)
                if variable_vars[brand].get():
                    self.session_brand_values[brand] = value
                    self.brand_values.pop(brand, None)
                else:
                    self.brand_values[brand] = value
                    self.session_brand_values.pop(brand, None)
            self._persist_settings()
            outcome["cancelled"] = False
            dlg.destroy()

        ttk.Button(button_row, text="Cancel", command=cancel, style="Quiet.TButton").pack(side="right")
        ttk.Button(button_row, text="Save and Continue", command=save, style="Primary.TButton").pack(
            side="right", padx=(0, 8)
        )

        dlg.protocol("WM_DELETE_WINDOW", cancel)
        dlg.bind("<Return>", lambda e: save())
        dlg.bind("<Escape>", lambda e: cancel())
        dlg.update_idletasks()
        try:
            self.root.update_idletasks()
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            dw, dh = dlg.winfo_width(), dlg.winfo_height()
            dlg.geometry(f"+{rx + (rw - dw) // 2}+{ry + (rh - dh) // 2}")
        except tk.TclError:
            pass

        dlg.wait_window()
        return outcome["cancelled"]

    def _export_df(self):
        """Output dataframe with internal-only columns stripped.

        Room and Cost are used internally by the QuickBooks upload for
        grouping and item-create cost. Neither belongs in the SaaSant
        Excel export.
        """
        if self.output_df is None:
            return self.output_df
        drop_cols = [c for c in (ROOM_COLUMN, COST_COLUMN) if c in self.output_df.columns]
        return self.output_df.drop(columns=drop_cols) if drop_cols else self.output_df

    def _ask_export_path(self, title: str) -> "Path | None":
        """Pop a Save-As dialog for the export file. Returns the chosen
        path or None if the user cancelled.

        Pre-fills the dialog with:
        - initialdir: the directory of the last saved export if it still
          exists, otherwise the user's Downloads folder.
        - initialfile: SalesOrder_<so>_<date>.xlsx so the user just has
          to confirm or tweak.
        """
        date_part = datetime.now().strftime("%m-%d-%Y")
        so_value = self.sales_order_no_var.get().strip() or "SO"
        safe_so = "".join(ch for ch in so_value if ch.isalnum() or ch in ("-", "_")) or "SO"
        default_name = f"SalesOrder_{safe_so}_{date_part}.xlsx"

        downloads = Path.home() / "Downloads"
        saved = (self.output_path_var.get() or "").strip()
        initial_dir = ""
        if saved:
            saved_parent = Path(saved).parent
            if saved_parent.exists():
                initial_dir = str(saved_parent)
        if not initial_dir and downloads.exists():
            initial_dir = str(downloads)

        chosen = filedialog.asksaveasfilename(
            title=title,
            defaultextension=".xlsx",
            initialdir=initial_dir or None,
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
        )
        return Path(chosen) if chosen else None

    def export_file(self):
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("No data", "Build preview first.")
            return
        output_path = self._ask_export_path("Export Template")
        if output_path is None:
            return
        self._export_df().to_excel(output_path, index=False, sheet_name="Sales Order")
        self.output_path_var.set(str(output_path))
        self._persist_settings()
        self._set_status(f"Exported template file: {output_path.name}")
        self._show_post_export_actions(output_path)

    def export_saasant_template(self):
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("No data", "Build preview first.")
            return
        output_path = self._ask_export_path("Export for SaaSant")
        if output_path is None:
            return
        self._export_df().to_excel(output_path, index=False, sheet_name="Sales Order")
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
        missing = self._validate_required_fields()
        if missing:
            log.warning("Upload to QuickBooks: aborted — blank required fields: %s", missing)
            messagebox.showwarning(
                "Required fields are blank",
                "Please fill in the following before uploading:\n\n  - " + "\n  - ".join(missing),
            )
            return
        upload_kwargs = dict(
            customer_name=self.customer_var.get().strip(),
            sales_order_no=self.sales_order_no_var.get().strip(),
            txn_date=self._normalize_date_for_qb(self.sales_order_date_var.get().strip()),
            due_date=self._normalize_date_for_qb(self.due_date_var.get().strip()),
            terms=self.terms_var.get().strip() or "Prepaid",
            shipping_method=self.shipping_method_var.get().strip() or "Standard Ground",
            memo=self.memo_var.get().strip(),
            lines=self.output_df.to_dict(orient="records"),
            tax_code=self.tax_code_var.get().strip() or "TAX",
            sales_tax_item="CA Tax",
            fallback_item=self.fallback_item_var.get().strip(),
            income_account=self.income_account_var.get().strip(),
            expense_account="COGS Non Inventory",
            group_by_room=True,
        )
        log.info(
            "Upload to QuickBooks: starting — customer=%r so=%r lines=%d",
            upload_kwargs["customer_name"],
            upload_kwargs["sales_order_no"],
            len(self.output_df),
        )
        self._run_upload_with_progress(upload_kwargs)

    def _run_upload_with_progress(self, upload_kwargs: dict) -> None:
        """Open a streaming progress dialog and run the SalesOrderAdd in a
        worker thread. Each progress callback from the SDK appends a line
        to the log Text widget so the user sees what's happening line by
        line. When the worker finishes, the dialog flips into a "Done"
        state with a single button to close.
        """
        c = UI
        dlg = tk.Toplevel(self.root)
        dlg.title("Uploading to QuickBooks")
        dlg.configure(bg=c["bg_window"])
        dlg.geometry("700x520")
        try:
            dlg.minsize(560, 420)
        except tk.TclError:
            pass
        dlg.transient(self.root)
        dlg.grab_set()
        # Disable closing during the upload — re-enabled on completion.
        dlg.protocol("WM_DELETE_WINDOW", lambda: None)

        accent_strip = tk.Frame(dlg, bg=c["accent"], height=4)
        accent_strip.pack(fill="x", side="top")

        body = tk.Frame(dlg, bg=c["bg_window"])
        body.pack(fill="both", expand=True, padx=22, pady=(18, 14))

        header = tk.Label(
            body,
            text="Uploading sales order...",
            bg=c["bg_window"],
            fg=c["accent"],
            font=("Segoe UI Semibold", 16),
            anchor="w",
        )
        header.pack(fill="x")

        sub_text = (
            f"SO #{upload_kwargs.get('sales_order_no', '?')}  ·  "
            f"{upload_kwargs.get('customer_name', '')}"
        )
        subheader = tk.Label(
            body,
            text=sub_text,
            bg=c["bg_window"],
            fg=c["text_secondary"],
            font=("Segoe UI", 10),
            anchor="w",
        )
        subheader.pack(fill="x", pady=(2, 12))

        # Pin the button row + progress bar to the bottom BEFORE the
        # expanding log, so "Done" is always visible.
        button_row = tk.Frame(body, bg=c["bg_window"])
        button_row.pack(fill="x", side="bottom", pady=(12, 0))
        progress_bar = ttk.Progressbar(body, mode="indeterminate")
        progress_bar.pack(fill="x", side="bottom", pady=(10, 0))
        progress_bar.start(12)

        log_wrap = tk.Frame(
            body, bg=c["bg_card"],
            highlightbackground=c["border"], highlightthickness=1, bd=0,
        )
        log_wrap.pack(fill="both", expand=True)
        log_text = tk.Text(
            log_wrap,
            wrap="none",
            bg=c["bg_card"],
            fg=c["text_primary"],
            font=("Consolas", 9),
            bd=0,
            relief="flat",
            padx=10,
            pady=8,
            height=14,
            width=10,  # let pack expand handle width; default 80 chars overflows
        )
        log_text.tag_configure("ok", foreground=c["success"])
        log_text.tag_configure("err", foreground=c["danger"])
        log_text.tag_configure("muted", foreground=c["text_tertiary"])
        sb = ttk.Scrollbar(log_wrap, orient="vertical", command=log_text.yview)
        log_text.configure(yscrollcommand=sb.set)
        log_text.configure(state="disabled")
        log_text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        state: dict = {"result": None, "error": None, "client": None, "import_seconds": 0.0}

        def append_log(msg: str, tag: str = "") -> None:
            log_text.configure(state="normal")
            if tag:
                log_text.insert("end", msg + "\n", tag)
            else:
                log_text.insert("end", msg + "\n")
            log_text.see("end")
            log_text.configure(state="disabled")

        def on_progress(msg: str) -> None:
            # SDK worker thread -> Tk main thread.
            self.root.after(0, append_log, msg)

        def finish() -> None:
            if dlg.winfo_exists():
                dlg.destroy()
            if state["error"]:
                self._set_qb_status("Connection Failed", state="disconnected")
                messagebox.showerror(
                    "QuickBooks Upload Error",
                    f"{state['error']}\n\nSee the Log menu for details (file: {LOG_PATH}).",
                )
                return
            self._set_qb_status("Connected", state="connected")
            self._set_status("Sales order uploaded to QuickBooks successfully.")
            created = list(getattr(state["client"], "last_created_items", []) or [])
            if created:
                self._show_created_items_report(state["result"], created)
            # Offer to clean up the downloaded source file now that the order
            # is safely in QuickBooks.
            self._prompt_delete_source_file()

        done_btn = ttk.Button(
            button_row,
            text="Done",
            style="Primary.TButton",
            command=finish,
            state="disabled",
        )
        done_btn.pack(side="right")

        def on_complete() -> None:
            try:
                progress_bar.stop()
                progress_bar.configure(mode="determinate", maximum=100, value=100)
            except tk.TclError:
                pass
            if state["error"]:
                header.config(text="Upload failed", fg=c["danger"])
                append_log("", tag="muted")
                append_log(f"ERROR: {state['error']}", tag="err")
                done_btn.configure(text="Close")
            else:
                header.config(text="Done — sales order uploaded", fg=c["success"])
                append_log("", tag="muted")
                append_log(state["result"] or "Upload complete.", tag="ok")
                # The whole point of the app: prove how much hand-keying it
                # just saved. Shown for every order — yes Spencer, even the
                # little ones under 10 lines. Also mirrored into the
                # Validation Messages panel on the main window.
                try:
                    lines = upload_kwargs.get("lines") or []
                    num_lines = len(lines)
                    num_rooms = _count_room_groups(lines)
                    saved_lines = self._time_saved_message_lines(
                        num_lines, num_rooms, state.get("import_seconds", 0.0)
                    )
                    if saved_lines:
                        append_log("", tag="muted")
                        for ln in saved_lines:
                            append_log(ln, tag="ok")
                        self._show_time_saved_in_validation(saved_lines)
                except Exception:
                    log.exception("time-saved estimate failed; skipping")
            done_btn.configure(state="normal")
            dlg.protocol("WM_DELETE_WINDOW", finish)
            dlg.bind("<Return>", lambda e: finish())
            dlg.bind("<Escape>", lambda e: finish())

        def worker() -> None:
            start = time.monotonic()
            try:
                client = self._qb_client()
                state["client"] = client
                result = client.upload_sales_order(
                    progress_cb=on_progress,
                    **upload_kwargs,
                )
                state["result"] = result
                log.info("Upload to QuickBooks: success — %s", result)
            except Exception as exc:
                state["error"] = str(exc)
                log.error("Upload to QuickBooks: failed — %s", exc)
                log.debug("Upload traceback:\n%s", traceback.format_exc())
            finally:
                state["import_seconds"] = time.monotonic() - start
                self.root.after(0, on_complete)

        # Center the dialog over the main window.
        dlg.update_idletasks()
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

        threading.Thread(target=worker, daemon=True).start()

    def _time_saved_message_lines(
        self, num_lines: int, num_rooms: int, import_seconds: float | None = None
    ) -> list[str]:
        """Build the cheeky time-comparison blurb, addressed to Spencer.

        Used in two places:
          - After Build Preview (import_seconds=None): an *estimate* — "by
            hand this would take about X; importing it takes seconds."
          - After a successful upload (import_seconds set): the real number —
            "by hand this would have taken X; the import took Y."

        Returns [] when there's nothing meaningful to show.
        """
        if num_lines <= 0:
            return []
        manual = _estimate_manual_seconds(num_lines, num_rooms)

        line_word = "line" if num_lines == 1 else "lines"
        if num_rooms > 0:
            room_word = "room" if num_rooms == 1 else "rooms"
            detail = (
                f"({num_lines} {line_word}, plus {num_rooms} {room_word} typed in "
                f"ALL CAPS and wrapped in ** yourself)"
            )
        else:
            detail = f"({num_lines} {line_word})"

        if import_seconds is None:
            # Estimate-only (post-preview): use future tense, no real timing.
            return [
                f"Hey Spencer — entering this order into QuickBooks by hand would take "
                f"you about {_format_duration(manual)} {detail}.",
                "This import will take you seconds.",
                "Still think the small orders aren't worth it? :)",
            ]

        # Clamp the real duration to a 1s floor so a sub-second round-trip
        # doesn't read as "0 seconds".
        import_secs = max(1, int(round(import_seconds)))
        return [
            f"Hey Spencer — entering this order by hand would have taken you about "
            f"{_format_duration(manual)} {detail}.",
            f"This import took you {_format_duration(import_secs)}.",
            "Still think the small orders aren't worth it? :)",
        ]

    def _show_time_saved_in_validation(self, saved_lines: list[str]) -> None:
        """Mirror the time-saved blurb into the Validation Messages panel on
        the main window, after the upload succeeds."""
        if not saved_lines:
            return
        try:
            self.error_text.delete("1.0", tk.END)
            self.error_text.insert(tk.END, "\n".join(saved_lines))
        except tk.TclError:
            log.debug("_show_time_saved_in_validation: error_text not available")

    def _prompt_delete_source_file(self) -> None:
        """After a successful upload, ask whether to delete the downloaded
        source Excel file (the order from David Meyer) or keep it. Deletes
        it for the user when they choose to."""
        src = (self.source_path_var.get() or "").strip()
        if not src:
            return
        src_path = Path(src)
        if not src_path.exists():
            return
        do_delete = messagebox.askyesno(
            "Delete the downloaded order file?",
            "This order is now in QuickBooks.\n\n"
            "Do you want to delete the downloaded Excel file from David Meyer, "
            "or keep it?\n\n"
            f"{src_path.name}\n\n"
            "Yes = delete it for me     No = keep it",
            icon="question",
        )
        if not do_delete:
            return
        try:
            src_path.unlink()
            log.info("Deleted source file after upload: %s", src_path)
            self.source_df = None
            self.source_path_var.set("")
            self._persist_settings()
            self._set_status(f"Deleted downloaded order file: {src_path.name}")
        except Exception as exc:
            log.exception("Failed to delete source file %s", src_path)
            messagebox.showerror(
                "Could Not Delete File",
                f"Couldn't delete the file:\n\n{src_path}\n\n{exc}",
            )


def _close_pyi_splash() -> None:
    """Dismiss the PyInstaller --splash image once the main window is up.
    pyi_splash is a runtime module that ONLY exists in --splash-built
    binaries; in plain `python app.py` dev runs the import simply fails
    and we move on.
    """
    try:
        import pyi_splash  # type: ignore[import-not-found]
    except Exception:
        return
    try:
        if pyi_splash.is_alive():
            pyi_splash.close()
    except Exception:
        log.exception("_close_pyi_splash: pyi_splash.close() raised")


def main():
    root = tk.Tk()
    app = SalesOrderApp(root)
    # Paint the main window once before tearing down the splash so the
    # transition is splash -> visible UI, not splash -> blank flash -> UI.
    try:
        root.update_idletasks()
        root.update()
    except tk.TclError:
        pass
    _close_pyi_splash()
    root.mainloop()


if __name__ == "__main__":
    main()
