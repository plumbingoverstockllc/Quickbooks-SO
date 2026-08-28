"""Each QuickBooks vendor appears exactly once in the brand/item pickers.

The pickers used to offer both the vendor's list name and its Company Name,
so a vendor stored as "Deluxe Vanity (ACH)" with company name "Deluxe Vanity"
showed up as two separate choices that looked like unrelated records — one of
them easily mistaken for a customer of the same name.

app.py imports Windows-only modules, so the real method is lifted out of the
source with ast rather than mirrored here; a copy would not catch the method
being changed back.
"""
from __future__ import annotations

import ast
from pathlib import Path
from types import SimpleNamespace


def _load_method(name: str):
    src = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.FunctionDef) and node.name == name:
            module = ast.Module(body=[node], type_ignores=[])
            namespace: dict = {}
            exec(compile(module, "app.py", "exec"), namespace)  # noqa: S102
            return namespace[name]
    raise AssertionError(f"{name} not found in app.py")


qb_vendor_names = _load_method("_qb_vendor_names")


def _app(vendors: list) -> SimpleNamespace:
    return SimpleNamespace(quickbooks_vendors=vendors)


def test_company_name_is_not_offered_as_a_second_choice():
    names = qb_vendor_names(
        _app([{"name": "Deluxe Vanity (ACH)", "companyName": "Deluxe Vanity", "isActive": True}])
    )
    assert names == ["Deluxe Vanity (ACH)"]


def test_one_entry_per_vendor():
    vendors = [
        {"name": "Deluxe Vanity (ACH)", "companyName": "Deluxe Vanity", "isActive": True},
        {"name": "Phylrich", "companyName": "Phylrich International", "isActive": True},
    ]
    assert len(qb_vendor_names(_app(vendors))) == len(vendors)


def test_inactive_vendors_are_left_out():
    vendors = [
        {"name": "Active Co", "companyName": "", "isActive": True},
        {"name": "Closed Co", "companyName": "", "isActive": False},
    ]
    assert qb_vendor_names(_app(vendors)) == ["Active Co"]


def test_duplicate_and_junk_rows_do_not_break_the_list():
    vendors = [
        {"name": "Phylrich", "isActive": True},
        {"name": "Phylrich", "isActive": True},
        {"name": "   ", "isActive": True},
        {"name": "", "isActive": True},
        "not a dict",
    ]
    assert qb_vendor_names(_app(vendors)) == ["Phylrich"]


if __name__ == "__main__":
    test_company_name_is_not_offered_as_a_second_choice()
    test_one_entry_per_vendor()
    test_inactive_vendors_are_left_out()
    test_duplicate_and_junk_rows_do_not_break_the_list()
    print("ok")
