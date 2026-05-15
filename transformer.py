from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Dict, List, Tuple

import pandas as pd


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
    source_df = pd.read_excel(source_path, sheet_name="Sheet1")
    optional_col = source_df.get("Optional")
    if optional_col is not None:
        source_df = source_df[optional_col.astype(str).str.lower().ne("yes")]
    source_df = source_df[source_df["SKU"].notna()].copy()
    source_df["Brand"] = source_df["Brand"].fillna("").astype(str).str.strip()
    return source_df


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
            }
        )

    output_df = pd.DataFrame(rows, columns=TEMPLATE_COLUMNS + [ROOM_COLUMN, COST_COLUMN])
    if output_df.empty:
        errors.append("No valid rows were generated.")
    return output_df, errors
