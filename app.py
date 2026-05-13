from __future__ import annotations

import json
import os
import hashlib
import threading
import subprocess
import tempfile
import urllib.parse
import tkinter as tk
import urllib.request
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from quickbooks_client import QuickBooksClient
from transformer import OrderSettings, line_pricing_keys, load_source, transform_to_template, unique_brands, unique_skus


DEFAULT_SOURCE = r"C:\Users\QB-PC\Downloads\Project-LisaStrongDesign-EliezerLabkowski301NHighland (1).xls"
DEFAULT_TEMPLATE = r"C:\Users\QB-PC\Downloads\SaasAnt Template for David Meyer.xlsx"
DEFAULT_OUTPUT = r"C:\Users\QB-PC\Downloads\SaaSant Sales Order - Auto Filled.xlsx"
APP_NAME = "QB Sales Order Converter"
APP_VERSION = "v0.9.53 Beta"
UPDATE_INFO_URL = "https://raw.githubusercontent.com/plumbingoverstockllc/Quickbooks-SO/main/releases/latest.json"
SETTINGS_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / APP_NAME
SETTINGS_PATH = SETTINGS_DIR / "settings.json"


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
        self.geometry("520x580")
        self.resizable(True, True)
        self.result = None
        self.source_df = source_df
        self.existing_brand_values = existing_brand_values or {}
        self.existing_item_values = existing_item_values or {}
        self.existing_line_values = existing_line_values or {}

        self.pricing_mode_var = tk.StringVar(value=pricing_mode)
        self.use_actual_cost_var = tk.BooleanVar(value=use_actual_cost)
        self.default_var = tk.StringVar(value=str(default_value))

        mode_frame = ttk.Frame(self)
        mode_frame.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(mode_frame, text="Pricing Mode:").pack(side="left")
        ttk.Radiobutton(mode_frame, text="Per Brand", value="brand", variable=self.pricing_mode_var, command=self._render_value_rows).pack(side="left", padx=(10, 6))
        ttk.Radiobutton(mode_frame, text="Per Item (SKU)", value="item", variable=self.pricing_mode_var, command=self._render_value_rows).pack(side="left", padx=(0, 6))
        ttk.Radiobutton(mode_frame, text="Per Line (Excel Row)", value="line", variable=self.pricing_mode_var, command=self._render_value_rows).pack(side="left")

        ttk.Checkbutton(
            self,
            text="Use actual cost (typed value is final rate, not multiplier)",
            variable=self.use_actual_cost_var,
            command=self._update_default_label,
        ).pack(anchor="w", padx=12, pady=(0, 8))

        self.default_label_var = tk.StringVar()
        self._update_default_label()
        ttk.Label(self, textvariable=self.default_label_var).pack(anchor="w", padx=12, pady=(0, 4))
        ttk.Entry(self, textvariable=self.default_var).pack(fill="x", padx=12, pady=(0, 10))

        wrapper = ttk.Frame(self)
        wrapper.pack(fill="both", expand=True, padx=12, pady=4)
        canvas = tk.Canvas(wrapper, highlightthickness=0)
        scroll = ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.value_vars: dict[str, tk.StringVar] = {}
        self._render_value_rows()

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", padx=12, pady=12)
        ttk.Button(buttons, text="Cancel", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text="Use Pricing Rules", command=self._submit).pack(side="right")

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
            row = ttk.Frame(self.inner)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=label, width=34).pack(side="left")
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
        self.geometry("840x600")
        self.resizable(True, True)
        self.result = None
        self._output_canvas: tk.Canvas | None = None
        self._source_canvas: tk.Canvas | None = None

        wrapper = ttk.Frame(self, padding=12)
        wrapper.pack(fill="both", expand=True)

        notebook = ttk.Notebook(wrapper)
        notebook.pack(fill="both", expand=True)
        output_tab = ttk.Frame(notebook)
        source_tab = ttk.Frame(notebook)
        notebook.add(output_tab, text="Output Preview")
        notebook.add(source_tab, text="Source Fields")
        notebook.select(output_tab)

        self.source_vars = {}
        self.output_vars = {}
        ttk.Label(output_tab, text=f"Output values for Excel line {excel_line_number}", style="SubHeader.TLabel").pack(anchor="w", pady=(8, 8), padx=8)
        out_canvas = tk.Canvas(output_tab, highlightthickness=0)
        out_scroll = ttk.Scrollbar(output_tab, orient="vertical", command=out_canvas.yview)
        out_inner = ttk.Frame(out_canvas)
        out_inner.bind("<Configure>", lambda e: out_canvas.configure(scrollregion=out_canvas.bbox("all")))
        out_canvas.create_window((0, 0), window=out_inner, anchor="nw")
        out_canvas.configure(yscrollcommand=out_scroll.set)
        out_canvas.pack(side="left", fill="both", expand=True)
        out_scroll.pack(side="right", fill="y")
        self._output_canvas = out_canvas

        header_line = ttk.Frame(out_inner)
        header_line.pack(fill="x", pady=(0, 8))
        ttk.Label(header_line, text=f"Line #: {excel_line_number}", width=20).pack(side="left")
        ttk.Label(header_line, text=f"SO: {output_row.get('Sales Order No', '')}").pack(side="left")
        ttk.Label(header_line, text=f"Customer: {output_row.get('Customer', '')}").pack(side="left", padx=(10, 0))

        for col, value in output_row.items():
            value = output_row.get(col, "")
            row = ttk.Frame(out_inner)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=col, width=34).pack(side="left")
            var = tk.StringVar(value="" if value is None else str(value))
            self.output_vars[col] = var
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

        canvas = tk.Canvas(source_tab, highlightthickness=0)
        scrollbar = ttk.Scrollbar(source_tab, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._source_canvas = canvas

        ttk.Label(inner, text=f"Edit source fields for Excel line {excel_line_number}", style="SubHeader.TLabel").pack(anchor="w", pady=(0, 8))
        for col, value in source_row.items():
            row = ttk.Frame(inner)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=col, width=34).pack(side="left")
            var = tk.StringVar(value="" if value is None else str(value))
            self.source_vars[col] = var
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

        btns = ttk.Frame(self, padding=(12, 8, 12, 12))
        btns.pack(fill="x")
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(btns, text="Save Changes", command=self._save).pack(side="right")

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
        self.root.geometry("1360x860")
        self.root.minsize(1200, 760)
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
        self.sales_order_date_var = tk.StringVar(
            value=self._normalize_date_for_display(
                self.settings.get("sales_order_date", self._today_display_date())
            )
        )
        self.due_date_var = tk.StringVar(
            value=self._normalize_date_for_display(
                self.settings.get("due_date", self._today_display_date())
            )
        )
        self.terms_var = tk.StringVar(value=self.settings.get("terms", "Prepaid"))
        self.shipping_method_var = tk.StringVar(value=self.settings.get("shipping_method", "Standard Ground"))
        self.memo_var = tk.StringVar(value=self.settings.get("memo", ""))
        self.currency_var = tk.StringVar(value=self.settings.get("currency", "USD"))
        self.tax_code_var = tk.StringVar(value=self.settings.get("tax_code", "TAX"))
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
        self.root.after(900, self._connect_quickbooks_on_startup)
        self.root.after(1200, self.check_for_updates_on_startup)

    def _configure_styles(self) -> None:
        self.style.configure(".", font=("Segoe UI", 10))
        self.style.configure("TLabelframe", padding=8)
        self.style.configure("Header.TLabel", font=("Segoe UI Semibold", 18))
        self.style.configure("SubHeader.TLabel", foreground="#4b5563")
        self.style.configure("Primary.TButton", padding=(13, 9), font=("Segoe UI", 10, "bold"))
        self.style.configure("Accent.TButton", padding=(13, 9))
        self.style.configure("Quiet.TButton", padding=(10, 8))
        self.style.configure("Status.TLabel", padding=(10, 8), foreground="#1f2937")
        self.style.configure("QbConnected.TLabel", foreground="#065f46", font=("Segoe UI", 9, "bold"))
        self.style.configure("QbDisconnected.TLabel", foreground="#b91c1c", font=("Segoe UI", 9, "bold"))
        self.style.configure("QbPending.TLabel", foreground="#92400e", font=("Segoe UI", 9, "bold"))

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self.root)
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Check for Updates", command=self.check_for_updates)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", f"{APP_NAME} {APP_VERSION}"))
        menu_bar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menu_bar)

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
        now = datetime.now()
        return f"{now.month}/{now.day}/{now.year}"

    def _normalize_date_for_qb(self, date_text: str) -> str:
        date_text = (date_text or "").strip()
        if not date_text:
            return ""
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(date_text, fmt)
                return f"{dt.month}/{dt.day}/{dt.year}"
            except ValueError:
                continue
        raise ValueError(f"Invalid date format: {date_text}. Use M/D/YYYY.")

    def _normalize_date_for_display(self, date_text: str) -> str:
        date_text = (date_text or "").strip()
        if not date_text:
            return ""
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(date_text, fmt)
                return f"{dt.month}/{dt.day}/{dt.year}"
            except ValueError:
                continue
        raise ValueError(f"Invalid date format: {date_text}. Use M/D/YYYY.")

    def check_for_updates(self, silent: bool = False) -> None:
        try:
            cache_bust_url = f"{UPDATE_INFO_URL}?t={int(datetime.now().timestamp())}"
            request = urllib.request.Request(
                cache_bust_url,
                headers={
                    "Cache-Control": "no-cache, no-store, max-age=0",
                    "Pragma": "no-cache",
                },
            )
            with urllib.request.urlopen(request, timeout=10) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            latest_version = str(payload.get("version", "")).strip()
            download_url = str(payload.get("url", "")).strip()
            sha256_hash = str(payload.get("sha256", "")).strip().lower()
            notes = str(payload.get("notes", "")).strip()

            if not latest_version or not download_url:
                raise RuntimeError("Update feed is missing version or url.")

            current_version = self._version_tuple(APP_VERSION)
            available_version = self._version_tuple(latest_version)
            if available_version <= current_version:
                if not silent:
                    messagebox.showinfo("No Updates", f"You're up to date on {APP_VERSION}.")
                self._set_status(f"Update check complete: {APP_VERSION} is current.")
                return

            prompt = f"New version available: v{latest_version}\nCurrent: {APP_VERSION}\n\nInstall now?"
            if notes:
                prompt = f"{prompt}\n\nRelease notes:\n{notes}"
            if not messagebox.askyesno("Update Available", prompt):
                return

            self._download_and_run_update(download_url, sha256_hash)
        except Exception as exc:
            if not silent:
                messagebox.showerror("Update Check Failed", str(exc))
            self._set_status("Update check failed. Use 'Check for Updates' to retry.")

    def check_for_updates_on_startup(self) -> None:
        self._set_status("Checking for updates...")
        self.check_for_updates(silent=True)

    def _download_and_run_update(self, url: str, expected_sha256: str) -> None:
        temp_dir = Path(tempfile.gettempdir()) / "qb_so_updates"
        temp_dir.mkdir(parents=True, exist_ok=True)
        filename = Path(urllib.parse.urlparse(url).path).name or "QB-Sales-Order-Converter-Setup.exe"
        installer_path = temp_dir / filename
        progress_dialog = tk.Toplevel(self.root)
        progress_dialog.title("Installing Update")
        progress_dialog.geometry("430x150")
        progress_dialog.resizable(False, False)
        progress_dialog.transient(self.root)
        progress_dialog.grab_set()
        progress_dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        status_var = tk.StringVar(value="Downloading update...")
        progress_var = tk.DoubleVar(value=0.0)
        ttk.Label(progress_dialog, textvariable=status_var, wraplength=390).pack(anchor="w", padx=18, pady=(18, 8))
        progress_bar = ttk.Progressbar(progress_dialog, mode="determinate", maximum=100, variable=progress_var)
        progress_bar.pack(fill="x", padx=18, pady=(0, 8))
        pct_label = ttk.Label(progress_dialog, text="0%")
        pct_label.pack(anchor="e", padx=18)

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

        header = ttk.Frame(root, padding=12)
        header.pack(fill="x")
        ttk.Label(header, text=f"Sales Order File Converter + QuickBooks Upload ({APP_VERSION})", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Upload source file, review mapped template rows, then export or upload directly to QuickBooks Desktop Enterprise.",
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        config = ttk.LabelFrame(root, text="Configuration", padding=12)
        config.pack(fill="x", padx=12, pady=8)

        self._path_row(config, "Source File", self.source_path_var, self._browse_source, 0)
        self._path_row(config, "SaaSant Template", self.template_path_var, self._browse_template, 1)
        self._path_row(config, "Output File", self.output_path_var, self._browse_output, 2)

        form = ttk.Frame(config)
        form.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        for i in range(6):
            form.columnconfigure(i, weight=1)

        self._form_entry(form, "Customer", self.customer_var, 0, 0, 2)
        self._form_entry(form, "Sales Order No", self.sales_order_no_var, 0, 2, 1)
        ttk.Button(form, text="Fetch Next from QuickBooks", command=self.fetch_next_so, style="Quiet.TButton").grid(row=1, column=3, padx=4, sticky="w")
        self._form_entry(form, "Sales Order Date", self.sales_order_date_var, 0, 4, 1)
        self._form_entry(form, "Due Date", self.due_date_var, 0, 5, 1)
        ttk.Label(form, text="Date format: M/D/YYYY (example: 5/12/2026)", style="SubHeader.TLabel").grid(
            row=2, column=4, columnspan=2, sticky="w", padx=4, pady=(0, 6)
        )

        self._form_entry(form, "Terms", self.terms_var, 2, 0, 1)
        self._form_entry(form, "Shipping Method", self.shipping_method_var, 2, 1, 1)
        self._form_entry(form, "Memo", self.memo_var, 2, 2, 2)
        self._form_entry(form, "Currency", self.currency_var, 2, 4, 1)
        self._form_entry(form, "Tax Code", self.tax_code_var, 2, 5, 1)

        qb_bar = ttk.Frame(config)
        qb_bar.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Button(qb_bar, text="Connect to QuickBooks Desktop", command=self.connect_quickbooks, style="Accent.TButton").pack(side="left")
        ttk.Button(qb_bar, text="QuickBooks Admin Setup", command=self.show_qb_admin_setup_guide, style="Quiet.TButton").pack(side="left", padx=8)
        ttk.Button(qb_bar, text="Check for Updates", command=self.check_for_updates, style="Quiet.TButton").pack(side="left", padx=8)
        self.qb_status_label = ttk.Label(qb_bar, textvariable=self.qb_status_var, style="QbDisconnected.TLabel")
        self.qb_status_label.pack(side="left", padx=10)

        actions = ttk.Frame(root, padding=(12, 0, 12, 10))
        actions.pack(fill="x")
        ttk.Button(actions, text="Build Preview", command=self.build_preview, style="Primary.TButton").pack(side="left")
        ttk.Button(actions, text="Change Pricing Rules", command=self.change_pricing_rules, style="Quiet.TButton").pack(side="left", padx=8)
        ttk.Button(actions, text="Export for SaaSant", command=self.export_saasant_template, style="Accent.TButton").pack(side="left", padx=8)
        ttk.Button(actions, text="Export Template File (Custom Path)", command=self.export_file, style="Quiet.TButton").pack(side="left", padx=8)
        ttk.Button(actions, text="Upload to QuickBooks", command=self.upload_to_quickbooks, style="Accent.TButton").pack(side="left")

        content = ttk.Panedwindow(root, orient="vertical")
        content.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        preview_frame = ttk.LabelFrame(content, text="Preview")
        content.add(preview_frame, weight=3)

        self.preview_notebook = ttk.Notebook(preview_frame)
        self.preview_notebook.pack(fill="both", expand=True)
        output_tab = ttk.Frame(self.preview_notebook)
        source_tab = ttk.Frame(self.preview_notebook)
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

        out_scroll_y = ttk.Scrollbar(output_tab, orient="vertical", command=self.output_tree.yview)
        out_scroll_x = ttk.Scrollbar(output_tab, orient="horizontal", command=self.output_tree.xview)
        self.output_tree.configure(yscrollcommand=out_scroll_y.set, xscrollcommand=out_scroll_x.set)
        output_tab.columnconfigure(0, weight=1)
        output_tab.rowconfigure(0, weight=1)
        self.output_tree.grid(row=0, column=0, sticky="nsew")
        out_scroll_y.grid(row=0, column=1, sticky="ns")
        out_scroll_x.grid(row=1, column=0, sticky="ew")
        self.output_tree.bind("<Double-1>", self._edit_output_row)

        self.source_tree = ttk.Treeview(source_tab, columns=self.source_preview_columns, show="headings")
        for col in self.source_preview_columns:
            self.source_tree.heading(col, text=col)
            width = 220 if col == "productName" else (80 if col == "Line #" else 120)
            self.source_tree.column(col, width=width, anchor="w")
        src_scroll_y = ttk.Scrollbar(source_tab, orient="vertical", command=self.source_tree.yview)
        src_scroll_x = ttk.Scrollbar(source_tab, orient="horizontal", command=self.source_tree.xview)
        self.source_tree.configure(yscrollcommand=src_scroll_y.set, xscrollcommand=src_scroll_x.set)
        source_tab.columnconfigure(0, weight=1)
        source_tab.rowconfigure(0, weight=1)
        self.source_tree.grid(row=0, column=0, sticky="nsew")
        src_scroll_y.grid(row=0, column=1, sticky="ns")
        src_scroll_x.grid(row=1, column=0, sticky="ew")
        self.source_tree.bind("<Double-1>", self._edit_source_row)

        error_frame = ttk.LabelFrame(content, text="Validation Messages")
        content.add(error_frame, weight=1)
        self.error_text = tk.Text(error_frame, height=8, wrap="word")
        self.error_text.pack(fill="both", expand=True)

        status_bar = ttk.Frame(root)
        status_bar.pack(fill="x", side="bottom")
        ttk.Label(status_bar, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w")

    def _path_row(self, parent, label, var, browse_cmd, row_idx):
        ttk.Label(parent, text=label).grid(row=row_idx, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row_idx, column=1, sticky="ew", padx=8)
        ttk.Button(parent, text="Browse", command=browse_cmd).grid(row=row_idx, column=2, sticky="e")
        parent.columnconfigure(1, weight=1)

    def _form_entry(self, parent, label, var, row, col, span):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=4, pady=(6, 0))
        ttk.Entry(parent, textvariable=var).grid(row=row + 1, column=col, columnspan=span, sticky="ew", padx=4, pady=(0, 6))

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

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)

    def _set_qb_status(self, message: str, state: str = "disconnected") -> None:
        self.qb_status_var.set(message)
        if self.qb_status_label is None:
            return
        style_name = "QbDisconnected.TLabel"
        if state == "connected":
            style_name = "QbConnected.TLabel"
        elif state == "pending":
            style_name = "QbPending.TLabel"
        self.qb_status_label.configure(style=style_name)

    def fetch_next_so(self):
        try:
            next_no = QuickBooksClient().get_next_sales_order_number()
            self.sales_order_no_var.set(next_no)
            self._set_status(f"Fetched next sales order number: {next_no}")
            self._set_qb_status("Connected", state="connected")
            messagebox.showinfo("Success", f"Next Sales Order number: {next_no}")
        except Exception as exc:
            self._set_qb_status("Connection Failed", state="disconnected")
            messagebox.showerror("QuickBooks Error", str(exc))

    def _connect_quickbooks_on_startup(self):
        self._set_qb_status("Connecting...", state="pending")
        self.connect_quickbooks(silent=True)

    def connect_quickbooks(self, silent: bool = False):
        try:
            company_name = QuickBooksClient().test_connection()
            self._set_qb_status(f"Connected: {company_name}", state="connected")
            self._set_status(f"Connected to QuickBooks company: {company_name}")
            if not silent:
                messagebox.showinfo("QuickBooks Connected", f"Connected successfully.\nCompany: {company_name}")
        except Exception as exc:
            self._set_qb_status("Not Connected", state="disconnected")
            error_text = str(exc)
            if silent:
                self._set_status("QuickBooks auto-connect failed. You can click Connect to QuickBooks Desktop.")
                return
            if self._is_qb_admin_authorization_error(error_text):
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
                "4) Re-click Connect to QuickBooks Desktop",
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
        for row_idx, row in self.output_df.iterrows():
            self.output_tree.insert("", "end", iid=str(row_idx), values=self._tree_values_from_row(row, row_idx))

        for i in self.source_tree.get_children():
            self.source_tree.delete(i)
        for source_row_index in self.source_df.index.tolist():
            self.source_tree.insert("", "end", iid=str(source_row_index), values=self._source_tree_values_from_row(source_row_index))

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
        today = datetime.now()
        date_part = f"{today.month}-{today.day}-{today.year}"
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
        dlg.geometry("560x220")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=14)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=msg, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(frame, text=str(output_path), wraplength=520).pack(anchor="w", pady=(8, 10))
        if saasant_ready:
            ttk.Label(frame, text="Upload this file directly in SaaSant.").pack(anchor="w", pady=(0, 8))

        btns = ttk.Frame(frame)
        btns.pack(fill="x", pady=(6, 0))
        ttk.Button(btns, text="Open File", command=lambda: self._open_file(output_path)).pack(side="left")
        ttk.Button(btns, text="Go to File (Explorer)", command=lambda: self._reveal_file(output_path)).pack(side="left", padx=8)
        ttk.Button(btns, text="Close", command=dlg.destroy).pack(side="right")

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    def _open_file(self, output_path: Path):
        os.startfile(str(output_path))

    def _reveal_file(self, output_path: Path):
        subprocess.Popen(["explorer", f"/select,{output_path}"])

    def upload_to_quickbooks(self):
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("No data", "Build preview first.")
            return
        try:
            result = QuickBooksClient().upload_sales_order(
                customer_name=self.customer_var.get().strip(),
                sales_order_no=self.sales_order_no_var.get().strip(),
                txn_date=self._normalize_date_for_qb(self.sales_order_date_var.get().strip()),
                due_date=self._normalize_date_for_qb(self.due_date_var.get().strip()),
                terms=self.terms_var.get().strip() or "Prepaid",
                shipping_method=self.shipping_method_var.get().strip() or "Standard Ground",
                memo=self.memo_var.get().strip(),
                lines=self.output_df.to_dict(orient="records"),
                tax_code=self.tax_code_var.get().strip() or "TAX",
            )
            self._set_qb_status("Connected", state="connected")
            self._set_status("Sales order uploaded to QuickBooks successfully.")
            messagebox.showinfo("QuickBooks Upload", result)
        except Exception as exc:
            self._set_qb_status("Connection Failed", state="disconnected")
            messagebox.showerror("QuickBooks Upload Error", str(exc))


def main():
    root = tk.Tk()
    SalesOrderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
