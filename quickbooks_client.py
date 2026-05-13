from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from typing import Iterable, List

import pythoncom
import win32com.client


class QuickBooksClient:
    def __init__(self, app_name: str = "SO Desktop App", company_file_path: str = "") -> None:
        self.app_name = app_name
        self.company_file_path = (company_file_path or "").strip()
        self._rp = None
        self._ticket = None

    def connect(self) -> None:
        pythoncom.CoInitialize()
        self._rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        self._rp.OpenConnection2("", self.app_name, 1)

        # Order matters: attach to whatever QuickBooks already has open first.
        # If we pass a path while QB has a different file open, QBXMLRP2 will
        # try to launch a second QB / open the file, which is what we want to
        # avoid. Only fall back to the saved path when no live session exists.
        path_candidates: list[str] = [""]
        if self.company_file_path:
            path_candidates.append(self.company_file_path)

        last_error = None
        for path in path_candidates:
            for open_mode in (2, 0, 1):  # 2=DontCare, 0=SingleUser, 1=MultiUser
                try:
                    self._ticket = self._rp.BeginSession(path, open_mode)
                    return
                except Exception as exc:
                    last_error = exc

        self.close()
        hint = ""
        if not self.company_file_path:
            hint = (
                "\n\nTip: set the QuickBooks Company File (.QBW) path in the app's Configuration "
                "card so it can open the file directly when QuickBooks isn't running."
            )
        raise RuntimeError(
            "Could not connect to QuickBooks Desktop. Open QuickBooks and your company file first, "
            "then run this app with the same Windows user/session as QuickBooks."
            f"{hint}\n\nDetails: {last_error}"
        )

    def close(self) -> None:
        if self._rp and self._ticket:
            try:
                self._rp.EndSession(self._ticket)
            except Exception:
                pass
        if self._rp:
            try:
                self._rp.CloseConnection()
            except Exception:
                pass
        self._rp = None
        self._ticket = None
        pythoncom.CoUninitialize()

    def _process(self, request_xml: str) -> str:
        if not self._rp or not self._ticket:
            raise RuntimeError("QuickBooks is not connected.")
        return self._rp.ProcessRequest(self._ticket, request_xml)

    def test_connection(self) -> str:
        self.connect()
        try:
            request_xml = """<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <CompanyQueryRq>
      <IncludeRetElement>CompanyName</IncludeRetElement>
    </CompanyQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""
            response = self._process(request_xml)
            root = ET.fromstring(response)
            company_name = root.findtext(".//CompanyRet/CompanyName", default="QuickBooks Company")
            return company_name
        finally:
            self.close()

    def _query_ref_numbers(self) -> List[str]:
        ref_numbers: List[str] = []
        iterator = "Start"
        iterator_id = ""

        while True:
            iterator_attr = f' iterator="{iterator}"'
            iterator_id_attr = f' iteratorID="{iterator_id}"' if iterator_id else ""
            query_xml = f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <SalesOrderQueryRq{iterator_attr}{iterator_id_attr}>
      <MaxReturned>200</MaxReturned>
      <IncludeRetElement>RefNumber</IncludeRetElement>
    </SalesOrderQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""
            response = self._process(query_xml)
            root = ET.fromstring(response)

            for ref in root.findall(".//SalesOrderRet/RefNumber"):
                if ref.text:
                    ref_numbers.append(ref.text.strip())

            query_rs = root.find(".//SalesOrderQueryRs")
            if query_rs is None:
                break
            remaining = int(query_rs.attrib.get("iteratorRemainingCount", "0"))
            if remaining <= 0:
                break
            iterator = "Continue"
            iterator_id = query_rs.attrib.get("iteratorID", "")
            if not iterator_id:
                break

        return ref_numbers

    def get_next_sales_order_number(self) -> str:
        self.connect()
        try:
            refs = self._query_ref_numbers()
            numeric_values = []
            for ref in refs:
                match = re.fullmatch(r"\d+", ref)
                if match:
                    numeric_values.append(int(ref))
            if not numeric_values:
                return "1"
            return str(max(numeric_values) + 1)
        finally:
            self.close()

    def upload_sales_order(
        self,
        customer_name: str,
        sales_order_no: str,
        txn_date: str,
        due_date: str,
        terms: str,
        shipping_method: str,
        memo: str,
        lines: Iterable[dict],
        tax_code: str,
    ) -> str:
        self.connect()
        try:
            line_xml = []
            for line in lines:
                sku = html.escape(str(line["Product/Service"]))
                desc = html.escape(str(line["Product/Service Description"]))
                qty = float(line["Product/Service Quantity"])
                rate = float(line["Product/Service Rate"])
                line_xml.append(
                    f"""
      <SalesOrderLineAdd>
        <ItemRef><FullName>{sku}</FullName></ItemRef>
        <Desc>{desc}</Desc>
        <Quantity>{qty}</Quantity>
        <Rate>{rate}</Rate>
        <SalesTaxCodeRef><FullName>{html.escape(tax_code)}</FullName></SalesTaxCodeRef>
      </SalesOrderLineAdd>"""
                )

            request_xml = f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <SalesOrderAddRq>
      <SalesOrderAdd>
        <CustomerRef><FullName>{html.escape(customer_name)}</FullName></CustomerRef>
        <RefNumber>{html.escape(sales_order_no)}</RefNumber>
        <TxnDate>{html.escape(txn_date)}</TxnDate>
        <DueDate>{html.escape(due_date)}</DueDate>
        <TermsRef><FullName>{html.escape(terms)}</FullName></TermsRef>
        <ShipMethodRef><FullName>{html.escape(shipping_method)}</FullName></ShipMethodRef>
        <Memo>{html.escape(memo)}</Memo>
        {''.join(line_xml)}
      </SalesOrderAdd>
    </SalesOrderAddRq>
  </QBXMLMsgsRq>
</QBXML>"""
            response = self._process(request_xml)
            root = ET.fromstring(response)
            add_rs = root.find(".//SalesOrderAddRs")
            if add_rs is None:
                raise RuntimeError("QuickBooks returned an unexpected response.")
            status_code = add_rs.attrib.get("statusCode", "")
            status_message = add_rs.attrib.get("statusMessage", "")
            if status_code != "0":
                raise RuntimeError(f"QuickBooks upload failed ({status_code}): {status_message}")

            ref = root.findtext(".//SalesOrderRet/RefNumber", default=sales_order_no)
            txn_id = root.findtext(".//SalesOrderRet/TxnID", default="")
            return f"Uploaded Sales Order #{ref} (TxnID: {txn_id})"
        finally:
            self.close()
