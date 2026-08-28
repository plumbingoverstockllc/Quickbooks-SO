"""Linux-safe tests for Estimate vs Sales Order export mapping."""
from __future__ import annotations

import ast
from pathlib import Path

import pandas as pd

from transformer import (
    DOCUMENT_ESTIMATE,
    DOCUMENT_SALES_ORDER,
    document_date_label,
    document_number_label,
    saasant_export_frame,
    saasant_filename_prefix,
    saasant_sheet_name,
    suggest_vendor,
)


def test_estimate_labels():
    assert document_number_label(DOCUMENT_ESTIMATE) == "Estimate No"
    assert document_date_label(DOCUMENT_ESTIMATE) == "Estimate Date"
    assert saasant_sheet_name(DOCUMENT_ESTIMATE) == "Estimate"
    assert saasant_filename_prefix(DOCUMENT_ESTIMATE) == "Estimate"


def test_sales_order_labels():
    assert document_number_label(DOCUMENT_SALES_ORDER) == "Sales Order No"
    assert saasant_sheet_name(DOCUMENT_SALES_ORDER) == "Sales Order"
    assert saasant_filename_prefix(DOCUMENT_SALES_ORDER) == "SalesOrder"


def test_saasant_export_renames_estimate_headers():
    df = pd.DataFrame(
        [
            {
                "Sales Order No": "1001",
                "Customer": "Acme",
                "Sales Order Date": "08/28/2026",
                "Product/Service": "SKU-1",
            }
        ]
    )
    out = saasant_export_frame(df, DOCUMENT_ESTIMATE)
    assert "Estimate No" in out.columns
    assert "Estimate Date" in out.columns
    assert "Sales Order No" not in out.columns
    assert out.iloc[0]["Estimate No"] == "1001"


def test_suggest_vendor_hansgrohe():
    known = ["Hansgrohe", "Kohler", "Moen"]
    assert suggest_vendor("Hans Grohe", known) == "Hansgrohe"


def test_no_dialog_hardcodes_sales_order_wording():
    """Dialogs kept saying "SO #1245" and "Sales Order Uploaded" while the
    user was building an estimate. Window titles and headings must come from
    the document_* helpers, not from a literal."""
    src = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
    offenders = []
    for node in ast.walk(ast.parse(src)):
        if not isinstance(node, ast.Constant) or not isinstance(node.value, str):
            continue
        text = node.value
        # In an f-string the literal part arrives on its own, so "SO #1245"
        # is seen here as "SO #" and "SO {number}" as "SO ".
        if text.startswith("SO #") or text == "SO ":
            offenders.append((node.lineno, text))
        elif text.strip() in ("Sales Order Uploaded", "Sales Order Upload"):
            offenders.append((node.lineno, text))
    assert not offenders, f"hardcoded sales-order wording in app.py: {offenders}"


if __name__ == "__main__":
    test_estimate_labels()
    test_sales_order_labels()
    test_saasant_export_renames_estimate_headers()
    test_suggest_vendor_hansgrohe()
    test_no_dialog_hardcodes_sales_order_wording()
    print("ok")
