from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple

import pandas as pd


DOCUMENT_SALES_ORDER = "sales_order"
DOCUMENT_ESTIMATE = "estimate"
DOCUMENT_TYPES = (DOCUMENT_SALES_ORDER, DOCUMENT_ESTIMATE)

TEMPLATE_COLUMNS = [
    "Sales Order No",
    "Customer",
    "Sales Order Date",
    "Due Date",
    "Terms",
    "Ship Date",
    "Shipping Address Line 1",
    "Billing Address Line 1",
    "Shipping Method",
    "Memo",
    "Product/Service",
    "Product/Service Description",
    "Product/Service Quantity",
    "Product/Service Rate",
    "Unit of Measure",
    "Product/Service Sales Tax Code",
    "Sales Tax Item",
    "Customer Sales Tax Code",
    "Currency",
]

# Extra column carried internally so the QuickBooks upload can group by room.
# Source files put the room/area in column L (the 12th column, 0-indexed 11).
ROOM_COLUMN = "Room"
ROOM_SOURCE_COLUMN_INDEX = 11

# Extra column carried internally so the upload can stamp the per-line cost
# onto any auto-created Non-Inventory items in QuickBooks. The source file's
# column I holds the wholesale cost.
COST_COLUMN = "Cost"

# Per-line note carried internally from the source file's "Notes" column so
# the QuickBooks upload can write it into the line's NOTES custom field.
# Like Room/Cost, it's stripped before the SaaSant export.
NOTES_COLUMN = "LineNotes"

# Deluxe Vanity & Kitchen pricing relationship. Their PO's NET price is what
# they pay us, which is MSRP * 0.45. Our cost is MSRP * 0.40 (the standard
# default multiplier). So from the NET on the PO we back out the implied MSRP
# (NET / 0.45); the normal cost-side multiplier (0.40) then yields our cost,
# which works out to ~12.5% below what we charge Deluxe.
DELUXE_SALE_MULTIPLIER = 0.45


@dataclass
class OrderSettings:
    customer_name: str
    sales_order_no: str
    sales_order_date: str
    due_date: str
    terms: str
    shipping_method: str
    memo: str
    currency: str
    sales_tax_code: str


def _as_number(value, default: float = 0.0) -> float:
    if pd.isna(value):
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def load_source(source_path: str) -> pd.DataFrame:
    """Load a source quote into the common line-item DataFrame.

    Dispatches on file extension:
      - .pdf  -> a showroom purchase order (currently Deluxe Vanity & Kitchen).
      - .xls / .xlsx -> the David Meyer Excel export (original format).
    Both return a DataFrame with at least SKU / Brand / MSRP / Price / Qty /
    productName columns so transform_to_template can consume either one.
    """
    if str(source_path).lower().endswith(".pdf"):
        return load_source_pdf(source_path)

    source_df = pd.read_excel(source_path, sheet_name="Sheet1")
    optional_col = source_df.get("Optional")
    if optional_col is not None:
        source_df = source_df[optional_col.astype(str).str.lower().ne("yes")]
    source_df = source_df[source_df["SKU"].notna()].copy()
    source_df["Brand"] = source_df["Brand"].fillna("").astype(str).str.strip()
    return source_df


def load_vendor_multipliers(source_path: str, sheet_name=0) -> tuple[dict, dict]:
    """Read the vendor price list and split brands into two buckets.

    The 'Vendor Price List for POS' sheet has Brand in column A and the Cost
    Multiplier in column F. A cell holds either:
      - a single clean number (e.g. 0.4, .405)  -> goes in `clean`
      - multi-line / tiered / textual pricing (e.g. "0.3848 China\\n0.45
        Others", "Net Pricing", "Single Orders = 0.405\\nBulk ...")
        -> goes in `ambiguous` as a one-line note, because no single
        multiplier can be auto-applied and the user must choose.

    Returns (clean: {brand: float}, ambiguous: {brand: note}).
    """
    df = pd.read_excel(source_path, sheet_name=sheet_name, header=None)

    # Locate the header row + the Brand / Cost Multiplier columns by label so
    # we don't hard-depend on exact positions (defaults: A=brand, F=mult).
    brand_col, mult_col, header_row = 0, 5, 0
    for r in range(min(10, len(df))):
        row = [str(df.iat[r, c]).strip().lower() if not pd.isna(df.iat[r, c]) else "" for c in range(df.shape[1])]
        if "brand" in row:
            header_row = r
            brand_col = row.index("brand")
            for c, label in enumerate(row):
                if "cost multiplier" in label or label == "multiplier":
                    mult_col = c
                    break
            break

    clean: dict[str, float] = {}
    ambiguous: dict[str, str] = {}
    for i in range(header_row + 1, len(df)):
        brand = df.iat[i, brand_col] if brand_col < df.shape[1] else None
        f = df.iat[i, mult_col] if mult_col < df.shape[1] else None
        if pd.isna(brand) or not str(brand).strip():
            continue
        brand = str(brand).strip()
        if pd.isna(f) or not str(f).strip():
            continue  # empty multiplier -> not in DB; will be prompted on demand
        raw = str(f).strip()
        if "\n" not in raw:
            try:
                clean[brand] = float(raw)
                continue
            except ValueError:
                pass
        ambiguous[brand] = " | ".join(ln.strip() for ln in raw.splitlines() if ln.strip())
    return clean, ambiguous


def _money_to_float(text: str) -> float:
    """'$729.000' / '1,299.50' -> float. Returns 0.0 on anything unparseable."""
    if text is None:
        return 0.0
    cleaned = str(text).replace("$", "").replace(",", "").strip()
    try:
        return float(cleaned)
    except (TypeError, ValueError):
        return 0.0


def load_source_pdf(source_path: str) -> pd.DataFrame:
    """Parse a Deluxe Vanity & Kitchen purchase-order PDF into the common
    line-item DataFrame.

    The PO has a bordered line-items table with the header row
    LN | MFG | PART NUMBER | DESCRIPTION | QTY | NET | U/M | EXT. NET.
    We locate that table on any page, map columns by their header label
    (cells can be split by merged-column gaps that show up as None), and
    pull one row per line item.

    Mapping into the common schema:
      - SKU         <- PART NUMBER
      - Brand       <- MFG
      - productName <- DESCRIPTION (newlines flattened to spaces)
      - Qty         <- QTY
      - Price       <- NET            (the per-unit sale rate on the order)
      - MSRP        <- NET / 0.45     (implied MSRP: Deluxe's NET = MSRP*0.45,
                                       so dividing recovers the MSRP. The
                                       standard 0.40 cost multiplier then
                                       yields our cost, ~12.5% under the NET.)
    """
    try:
        import pdfplumber
    except Exception as exc:  # pragma: no cover - bundled in the build
        # Log the REAL underlying error (e.g. a missing transitive native
        # dependency) so a support log shows what's actually wrong instead of
        # just "pdfplumber not available".
        import logging
        logging.getLogger("qb_so_app").exception(
            "load_source_pdf: pdfplumber import failed"
        )
        raise RuntimeError(
            "Reading PDF quotes failed to load the PDF engine: "
            f"{type(exc).__name__}: {exc}\n\n"
            "Reinstall the latest DMQuotes. (Details are in the log — "
            "Log → Open Log File.)"
        ) from exc

    rows: list[dict] = []
    with pdfplumber.open(source_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                header_idx, col_map = _find_line_item_header(table)
                if header_idx is None:
                    continue
                for raw in table[header_idx + 1:]:
                    parsed = _parse_pdf_line_item(raw, col_map)
                    if parsed is not None:
                        rows.append(parsed)

    if not rows:
        raise RuntimeError(
            "No line items could be read from this PDF. It may not be a "
            "Deluxe Vanity & Kitchen purchase order, or it may be a scanned "
            "image rather than a text PDF."
        )

    df = pd.DataFrame(rows, columns=["SKU", "Brand", "MSRP", "Price", "Qty", "productName"])
    df["Brand"] = df["Brand"].fillna("").astype(str).str.strip()
    df = df[df["SKU"].notna() & (df["SKU"].astype(str).str.strip() != "")].copy()
    df.reset_index(drop=True, inplace=True)
    return df


def _find_line_item_header(table: list) -> tuple:
    """Return (header_row_index, {canonical_field: column_index}) for the
    line-items table, or (None, {}) if this table isn't it. Recognized by a
    row that contains both a PART NUMBER and a QTY column."""
    aliases = {
        "LN": "LN",
        "MFG": "MFG",
        "PART NUMBER": "PART",
        "DESCRIPTION": "DESC",
        "QTY": "QTY",
        "NET": "NET",
        "U/M": "UM",
        "EXT. NET": "EXT",
    }
    for ridx, row in enumerate(table):
        col_map: dict[str, int] = {}
        for cidx, cell in enumerate(row):
            if cell is None:
                continue
            label = str(cell).strip().upper()
            if label in aliases and aliases[label] not in col_map:
                col_map[aliases[label]] = cidx
        if "PART" in col_map and "QTY" in col_map:
            return ridx, col_map
    return None, {}


def _parse_pdf_line_item(row: list, col_map: dict) -> dict | None:
    """Turn one raw table row into a common-schema line dict, or None when the
    row isn't a line item (totals, blank rows, continuation text)."""
    def cell(field: str) -> str:
        idx = col_map.get(field)
        if idx is None or idx >= len(row):
            return ""
        val = row[idx]
        return "" if val is None else str(val).strip()

    part = cell("PART")
    if not part:
        return None
    # A real line item has a numeric LN and a numeric QTY. This filters out
    # the "PO TOTAL" footer and any stray rows.
    ln = cell("LN")
    if ln and not ln.replace(".", "").isdigit():
        return None

    qty_text = cell("QTY").replace(",", "")
    try:
        qty = int(float(qty_text)) if qty_text else 1
    except ValueError:
        qty = 1
    qty = max(1, qty)

    net = _money_to_float(cell("NET"))
    desc = " ".join(cell("DESC").split())  # flatten internal newlines/spaces

    # NET is what Deluxe pays us (= MSRP * 0.45). Recover the implied MSRP so
    # the cost-side 0.40 multiplier produces our actual cost. NET stays the
    # sale rate on the order.
    implied_msrp = round(net / DELUXE_SALE_MULTIPLIER, 2) if net > 0 else 0.0

    return {
        "SKU": part,
        "Brand": cell("MFG"),
        "MSRP": implied_msrp,
        "Price": net,
        "Qty": qty,
        "productName": desc,
    }


def unique_brands(source_df: pd.DataFrame) -> List[str]:
    return sorted({b for b in source_df["Brand"].tolist() if b})


def unique_skus(source_df: pd.DataFrame) -> List[str]:
    return sorted({str(s).strip() for s in source_df["SKU"].tolist() if str(s).strip()})


def line_pricing_keys(source_df: pd.DataFrame) -> List[tuple[str, str]]:
    keys: List[tuple[str, str]] = []
    for idx, row in source_df.iterrows():
        excel_line = str(int(idx) + 2)
        sku = str(row.get("SKU", "")).strip()
        brand = str(row.get("Brand", "")).strip()
        label = f"Line {excel_line} - {brand} {sku}".strip()
        keys.append((excel_line, label))
    return keys


def transform_to_template(
    source_df: pd.DataFrame,
    settings: OrderSettings,
    pricing_mode: str,
    use_actual_cost: bool,
    pricing_values: Dict[str, float],
    default_value: float,
) -> Tuple[pd.DataFrame, List[str]]:
    errors: List[str] = []
    rows = []

    # Resolve the room column by positional index (column L) so the source
    # file's header name doesn't matter.
    room_series = None
    if source_df.shape[1] > ROOM_SOURCE_COLUMN_INDEX:
        room_series = source_df.iloc[:, ROOM_SOURCE_COLUMN_INDEX]

    for idx, row in source_df.iterrows():
        brand = str(row.get("Brand", "")).strip()
        sku = str(row.get("SKU", "")).strip()
        if not sku:
            errors.append(f"Row {idx + 2}: Missing SKU, skipped.")
            continue
        if room_series is not None:
            raw_room = room_series.loc[idx]
            room_value = "" if pd.isna(raw_room) else str(raw_room).strip()
        else:
            room_value = ""

        qty = max(1, int(math.ceil(_as_number(row.get("Qty"), default=1))))
        msrp = _as_number(row.get("MSRP"), default=0.0)
        source_price = _as_number(row.get("Price"), default=0.0)
        if msrp <= 0:
            errors.append(f"Row {idx + 2} ({sku}): MSRP is 0 or missing.")

        if pricing_mode == "brand":
            key = brand
        elif pricing_mode == "item":
            key = sku
        else:
            key = str(int(idx) + 2)
        configured_value = pricing_values.get(key, default_value)
        # Cost (used for the auto-created QuickBooks item's PurchaseCost
        # and any future PO line) is *always* MSRP * multiplier (or, when
        # the user chose "use actual cost", the configured value itself).
        # Source column I is intentionally ignored for cost purposes.
        if use_actual_cost:
            cost_value = round(configured_value, 2)
        else:
            cost_value = round(msrp * configured_value, 2)
        # Sales-order rate is the price we billed the customer. Source
        # column J "Price" holds that pre-calculated value. When it's
        # missing or zero we fall back to the cost-side calculation so
        # the line still has *some* number.
        if source_price > 0:
            rate = round(source_price, 2)
        else:
            rate = cost_value

        product_name = str(row.get("productName", "")).strip()
        desc_parts = [brand, sku, product_name]
        combined_desc = " ".join(part for part in desc_parts if part).strip()

        # Per-line note from the source "Notes" column (blank when absent).
        raw_note = row.get("Notes", "")
        line_note = "" if (raw_note is None or (isinstance(raw_note, float) and math.isnan(raw_note))) else str(raw_note).strip()
        if line_note.lower() == "nan":
            line_note = ""

        rows.append(
            {
                "Sales Order No": settings.sales_order_no,
                "Customer": settings.customer_name,
                "Sales Order Date": settings.sales_order_date,
                "Due Date": settings.due_date,
                "Terms": settings.terms,
                "Ship Date": "",
                "Shipping Address Line 1": "",
                "Billing Address Line 1": "",
                "Shipping Method": settings.shipping_method,
                "Memo": settings.memo,
                "Product/Service": sku,
                "Product/Service Description": combined_desc,
                "Product/Service Quantity": qty,
                "Product/Service Rate": rate,
                "Unit of Measure": "$",
                "Product/Service Sales Tax Code": settings.sales_tax_code,
                "Sales Tax Item": "None",
                "Customer Sales Tax Code": "None",
                "Currency": settings.currency,
                ROOM_COLUMN: room_value,
                COST_COLUMN: cost_value,
                NOTES_COLUMN: line_note,
            }
        )

    output_df = pd.DataFrame(rows, columns=TEMPLATE_COLUMNS + [ROOM_COLUMN, COST_COLUMN, NOTES_COLUMN])
    if output_df.empty:
        errors.append("No valid rows were generated.")
    return output_df, errors


def is_document_type(value: object) -> bool:
    return value in DOCUMENT_TYPES


def document_noun(document_type: str) -> str:
    if document_type == DOCUMENT_ESTIMATE:
        return "estimate"
    return "sales order"


def document_noun_title(document_type: str) -> str:
    if document_type == DOCUMENT_ESTIMATE:
        return "Estimate"
    return "Sales Order"


def document_number_label(document_type: str) -> str:
    if document_type == DOCUMENT_ESTIMATE:
        return "Estimate No"
    return "Sales Order No"


def document_date_label(document_type: str) -> str:
    if document_type == DOCUMENT_ESTIMATE:
        return "Estimate Date"
    return "Sales Order Date"


def saasant_sheet_name(document_type: str) -> str:
    if document_type == DOCUMENT_ESTIMATE:
        return "Estimate"
    return "Sales Order"


def saasant_filename_prefix(document_type: str) -> str:
    if document_type == DOCUMENT_ESTIMATE:
        return "Estimate"
    return "SalesOrder"


def saasant_header_for(document_type: str, internal_column: str) -> str:
    if internal_column == "Sales Order No":
        return document_number_label(document_type)
    if internal_column == "Sales Order Date":
        return document_date_label(document_type)
    return internal_column


def saasant_export_frame(df: pd.DataFrame, document_type: str) -> pd.DataFrame:
    """Rename SaaSant header columns for Estimate vs Sales Order.

    Internal preview columns stay as Sales Order No/Date so QuickBooks upload
    and the rest of the app keep one schema. Only the Excel export remaps.
    """
    out = df.copy()
    rename = {}
    for col in list(out.columns):
        mapped = saasant_header_for(document_type, col)
        if mapped != col:
            rename[col] = mapped
    if rename:
        out = out.rename(columns=rename)
    return out


def normalize_vendor_name(name: str) -> str:
    text = str(name or "").lower().replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def compact_vendor_name(name: str) -> str:
    return normalize_vendor_name(name).replace(" ", "")


def suggest_vendor(source_name: str, known_names: list[str]) -> str | None:
    needle = compact_vendor_name(source_name)
    if not needle:
        return None
    exact = [name for name in known_names if compact_vendor_name(name) == needle]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        return exact[0]
    contained = [
        name
        for name in known_names
        if needle and (
            compact_vendor_name(name).find(needle) >= 0
            or needle.find(compact_vendor_name(name)) >= 0
        )
        and compact_vendor_name(name)
    ]
    if len(contained) == 1:
        return contained[0]
    return None
