from __future__ import annotations

import ctypes
import html
import logging
import os
import re
import subprocess
import traceback
import xml.etree.ElementTree as ET
from typing import Iterable, List

import pythoncom
import win32com.client


log = logging.getLogger("qb_so_app.quickbooks")

_QB_PROCESS_NAMES = ("qbw32.exe", "qbw64.exe", "qbw.exe")


def _is_quickbooks_running() -> bool:
    """Return True if a QuickBooks Desktop process is visible to this Windows session.

    Note: an elevated QuickBooks process running in a different UAC context may
    still be hidden from a non-elevated tasklist call, but the common case (QB
    running at the same privilege level) is covered.
    """
    try:
        result = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=5,
            creationflags=0x08000000,  # CREATE_NO_WINDOW
        )
    except Exception:
        return False
    output = result.stdout.lower()
    return any(name in output for name in _QB_PROCESS_NAMES)


def _is_process_elevated(pid: int) -> str:
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    TOKEN_QUERY = 0x0008
    TokenElevation = 20
    process_handle = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not process_handle:
        return "unknown"
    try:
        token_handle = ctypes.c_void_p()
        if not ctypes.windll.advapi32.OpenProcessToken(process_handle, TOKEN_QUERY, ctypes.byref(token_handle)):
            return "unknown"
        try:
            elevation = ctypes.c_ulong()
            size = ctypes.c_ulong()
            ok = ctypes.windll.advapi32.GetTokenInformation(
                token_handle,
                TokenElevation,
                ctypes.byref(elevation),
                ctypes.sizeof(elevation),
                ctypes.byref(size),
            )
            if not ok:
                return "unknown"
            return "elevated" if elevation.value else "standard"
        finally:
            ctypes.windll.kernel32.CloseHandle(token_handle)
    finally:
        ctypes.windll.kernel32.CloseHandle(process_handle)


def _current_process_elevation() -> str:
    return _is_process_elevated(os.getpid())


def _quickbooks_process_details() -> list[dict]:
    details: list[dict] = []
    try:
        result = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=5,
            creationflags=0x08000000,  # CREATE_NO_WINDOW
        )
        for raw_line in (result.stdout or "").splitlines():
            line = raw_line.strip()
            if not line:
                continue
            parts = [p.strip().strip('"') for p in line.split('","')]
            if len(parts) < 2:
                continue
            exe_name = parts[0].lower().strip('"')
            if exe_name not in _QB_PROCESS_NAMES:
                continue
            try:
                pid = int(parts[1].strip('"'))
            except Exception:
                pid = -1
            details.append(
                {
                    "name": exe_name,
                    "pid": pid,
                    "elevation": _is_process_elevated(pid) if pid > 0 else "unknown",
                }
            )
    except Exception:
        return []
    return details


def _runtime_context_note() -> str:
    qb_details = _quickbooks_process_details()
    qb_summary = (
        ", ".join(
            f"{row['name']} pid={row['pid']} elevation={row['elevation']}"
            for row in qb_details
        )
        if qb_details
        else "none detected"
    )
    return (
        f"App PID={os.getpid()} elevation={_current_process_elevation()}; "
        f"QuickBooks processes={len(qb_details)} [{qb_summary}]"
    )


def _normalize_company_path(path: str) -> str:
    normalized = (path or "").strip().replace("/", "\\")
    return normalized


class QuickBooksClient:
    def __init__(self, app_name: str = "SO Desktop App", company_file_path: str = "") -> None:
        self.app_name = app_name
        self.company_file_path = _normalize_company_path(company_file_path)
        self._rp = None
        self._ticket = None

    def connect(self) -> None:
        log.info("connect(): app_name=%r, company_file_path=%r", self.app_name, self.company_file_path)
        pythoncom.CoInitialize()
        log.debug("connect(): CoInitialize OK")
        try:
            self._rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
            log.debug("connect(): Dispatched QBXMLRP2.RequestProcessor")
        except Exception as exc:
            log.exception("connect(): Dispatch failed")
            pythoncom.CoUninitialize()
            raise RuntimeError(
                "Could not load the QuickBooks SDK (QBXMLRP2.RequestProcessor). "
                "Make sure QuickBooks Desktop is installed on this machine."
                f"\n\nDetails: {exc}"
            )

        try:
            self._rp.OpenConnection2("", self.app_name, 1)
            log.debug("connect(): OpenConnection2 OK")
        except Exception as exc:
            log.exception("connect(): OpenConnection2 failed")
            self.close()
            raise RuntimeError(f"QuickBooks OpenConnection2 failed: {exc}")

        qb_already_running = _is_quickbooks_running()
        log.info("connect(): QB process detected = %s", qb_already_running)

        # First: try to attach to whatever QuickBooks already has open. Passing
        # an empty path tells QBXMLRP2 to use the current session if one exists.
        # Keep the FIRST failure (mode 2 = DontCare) — that's the most diagnostic
        # error to surface, since the later modes typically fail for the same
        # reason but with less specific text.
        attach_error = None
        for open_mode in (2, 0, 1):  # 2=DontCare, 0=SingleUser, 1=MultiUser
            try:
                log.debug("connect(): BeginSession(path='', mode=%d)", open_mode)
                self._ticket = self._rp.BeginSession("", open_mode)
                log.info("connect(): attached to running session (mode=%d)", open_mode)
                return
            except Exception as exc:
                log.warning("connect(): BeginSession(path='', mode=%d) failed — %s", open_mode, exc)
                if attach_error is None:
                    attach_error = exc

        if qb_already_running:
            log.error("connect(): QB is running but attach failed; path-based launch fallback is disabled")
            context_note = _runtime_context_note()
            log.error("connect(): context diagnostics: %s", context_note)
            self.close()
            raise RuntimeError(
                "QuickBooks Desktop is running, but this app could not attach to the open "
                "company file. Common causes: the app has not been granted access to this "
                "specific company file by a QuickBooks Admin, QuickBooks is at the "
                "\"No Company Open\" / login screen, or this app is running in a different "
                "Windows/UAC context than QuickBooks."
                f"\n\nAttach details: {attach_error}"
                f"\nContext: {context_note}"
            )

        # Never launch/open QuickBooks as a fallback from this app.
        log.error("connect(): QB not running or not attachable; launch fallback disabled")
        context_note = _runtime_context_note()
        log.error("connect(): context diagnostics: %s", context_note)
        self.close()
        raise RuntimeError(
            "Could not connect to QuickBooks Desktop. This app is configured to attach only "
            "to an already-open QuickBooks company session and will not launch/open QuickBooks "
            "as a fallback.\n\nOpen QuickBooks manually, open/login to the target company file, "
            "and ensure this app runs under the same Windows user and UAC elevation level.\n\n"
            f"Details: {attach_error}"
            f"\nContext: {context_note}"
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
