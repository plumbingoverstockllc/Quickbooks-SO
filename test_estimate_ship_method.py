"""Ensure EstimateAdd XML never includes ShipMethodRef (QB parse error)."""
from __future__ import annotations

import re


def _header_xml(
    *,
    is_estimate: bool,
    shipping_method: str = "Standard Ground",
    terms: str = "Prepaid",
) -> str:
    """Mirror upload_sales_order header assembly for unit testing."""
    parts: list[str] = [
        "<CustomerRef><FullName>Acme</FullName></CustomerRef>",
        "<TxnDate>2026-08-27</TxnDate>",
        "<RefNumber>1001</RefNumber>",
    ]
    if terms:
        parts.append(f"<TermsRef><FullName>{terms}</FullName></TermsRef>")
    parts.append("<DueDate>2026-08-27</DueDate>")
    if shipping_method and not is_estimate:
        parts.append(
            f"<ShipMethodRef><FullName>{shipping_method}</FullName></ShipMethodRef>"
        )
    parts.append("<ItemSalesTaxRef><FullName>CA Tax</FullName></ItemSalesTaxRef>")
    parts.append("<CustomerSalesTaxCodeRef><FullName>TAX</FullName></CustomerSalesTaxCodeRef>")
    return "".join(parts)


def test_estimate_omits_ship_method():
    xml = _header_xml(is_estimate=True)
    assert "ShipMethodRef" not in xml
    assert "ItemSalesTaxRef" in xml
    # DueDate must still precede ItemSalesTaxRef (EstimateAdd order).
    assert xml.index("DueDate") < xml.index("ItemSalesTaxRef")


def test_sales_order_includes_ship_method():
    xml = _header_xml(is_estimate=False)
    assert "ShipMethodRef" in xml
    assert re.search(
        r"DueDate.*ShipMethodRef.*ItemSalesTaxRef", xml, re.DOTALL
    )


if __name__ == "__main__":
    test_estimate_omits_ship_method()
    test_sales_order_includes_ship_method()
    print("ok")
