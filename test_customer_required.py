"""DMQuotes must never create a QuickBooks customer on its own.

v1.069 added a CustomerAdd fallback so uploads would stop failing with an
opaque 3140 on CustomerRef. That traded a clear failure for a silent one: a
typo or a slightly different spelling produced a brand-new near-duplicate
customer that someone had to find and merge later. The upload now stops and
says so instead.
"""
from __future__ import annotations

import ast
from pathlib import Path

from quickbooks_client import CUSTOMER_NOT_FOUND_MESSAGE, CustomerNotFoundError

CLIENT_SRC = Path(__file__).with_name("quickbooks_client.py").read_text(encoding="utf-8")


def test_error_message_is_the_wording_the_user_sees():
    assert CUSTOMER_NOT_FOUND_MESSAGE == "Customer not in system, please add it first"
    assert str(CustomerNotFoundError("Mendel Kovalsky")) == CUSTOMER_NOT_FOUND_MESSAGE


def test_error_carries_the_name_so_the_popup_can_show_it():
    assert CustomerNotFoundError("Mendel Kovalsky").customer_name == "Mendel Kovalsky"


def test_it_is_catchable_as_a_plain_upload_failure():
    """The upload worker catches it specifically, but any caller that only
    knows about RuntimeError must still handle it."""
    assert issubclass(CustomerNotFoundError, RuntimeError)


def test_no_customer_add_request_anywhere_in_the_client():
    assert "CustomerAddRq" not in CLIENT_SRC
    assert "<CustomerAdd>" not in CLIENT_SRC


def test_upload_checks_the_customer_before_sending_anything():
    """The check is worthless if it runs after the transaction is submitted."""
    upload = next(
        node
        for node in ast.walk(ast.parse(CLIENT_SRC))
        if isinstance(node, ast.FunctionDef) and node.name == "upload_sales_order"
    )
    checks = [
        node.lineno
        for node in ast.walk(upload)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "_require_customer_exists"
    ]
    sends = [
        node.lineno
        for node in ast.walk(upload)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "_process"
    ]
    assert checks, "upload_sales_order no longer verifies the customer"
    assert sends, "expected upload_sales_order to submit a request"
    assert min(checks) < min(sends)


if __name__ == "__main__":
    test_error_message_is_the_wording_the_user_sees()
    test_error_carries_the_name_so_the_popup_can_show_it()
    test_it_is_catchable_as_a_plain_upload_failure()
    test_no_customer_add_request_anywhere_in_the_client()
    test_upload_checks_the_customer_before_sending_anything()
    print("ok")
