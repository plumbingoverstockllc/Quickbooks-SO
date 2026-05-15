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

# Sales order numbers below this are known-historical and never the "next".
# Used as a FromRefNumber filter so the SalesOrderQuery only paginates the
# tail of the table instead of the entire history.
MIN_SALES_ORDER_NUMBER = "168250"


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
    """Return one entry per *real* QuickBooks UI process.

    The QBXMLRP2 SDK runs an auxiliary qbw.exe whose top-level window title
    is "DDE Server Window". That helper is part of QuickBooks itself — it is
    expected and should not be treated as a second user-facing QB instance.
    Likewise tasklist sometimes reports system-side qbw helpers with no
    window. We filter both so the "multiple QB processes" safety gate only
    fires for genuine second sessions.
    """
    details: list[dict] = []
    try:
        result = subprocess.run(
            ["tasklist", "/V", "/FO", "CSV", "/NH"],
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
            window_title = parts[-1].strip() if len(parts) >= 9 else ""
            # Skip SDK helpers — they always carry the DDE server title.
            if window_title and window_title.lower() == "dde server window":
                continue
            details.append(
                {
                    "name": exe_name,
                    "pid": pid,
                    "elevation": _is_process_elevated(pid) if pid > 0 else "unknown",
                    "window_title": window_title,
                }
            )
    except Exception:
        return []
    return details


def _quickbooks_window_titles() -> list[str]:
    titles: list[str] = []
    try:
        result = subprocess.run(
            ["tasklist", "/V", "/FO", "CSV", "/NH"],
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
            if len(parts) < 9:
                continue
            exe_name = parts[0].lower().strip('"')
            if exe_name not in _QB_PROCESS_NAMES:
                continue
            window_title = parts[-1].strip()
            if window_title and window_title.upper() != "N/A":
                titles.append(window_title)
    except Exception:
        return []
    return titles


def _has_no_company_open_window() -> bool:
    for title in _quickbooks_window_titles():
        if "no company open" in title.lower():
            return True
    return False


def _has_secondary_window_title() -> bool:
    for title in _quickbooks_window_titles():
        if "(secondary)" in title.lower():
            return True
    return False


def _is_company_session_lost_error(exc: Exception | None) -> bool:
    if exc is None:
        return False
    text = str(exc).lower()
    return (
        "-2147220478" in text
        or "-2147220469" in text
        or "could not get the name of the current quickbooks company data file" in text
        or "unexpected error. check the \"qbsdklog.txt\"" in text
    )


def _runtime_context_note() -> str:
    qb_details = _quickbooks_process_details()
    qb_titles = _quickbooks_window_titles()
    qb_summary = (
        ", ".join(
            f"{row['name']} pid={row['pid']} elevation={row['elevation']}"
            for row in qb_details
        )
        if qb_details
        else "none detected"
    )
    title_summary = ", ".join(qb_titles) if qb_titles else "none"
    return (
        f"App PID={os.getpid()} elevation={_current_process_elevation()}; "
        f"QuickBooks processes={len(qb_details)} [{qb_summary}]; "
        f"QB window titles=[{title_summary}]"
    )


def _normalize_company_path(path: str) -> str:
    normalized = (path or "").strip().replace("/", "\\")
    return normalized


def _to_qb_date(date_text: str) -> str:
    """Convert a user-entered or display date to QBXML's YYYY-MM-DD format.

    QBXML's DATETYPE requires ISO-8601 dates. Common user formats (MM/DD/YYYY,
    MM-DD-YYYY, M/D/YY, etc.) need to be normalized before being inserted into
    the request XML, otherwise QuickBooks rejects the whole payload with
    -2147220480 "QuickBooks found an error when parsing the provided XML
    text stream."
    """
    text = (date_text or "").strip()
    if not text:
        return ""
    from datetime import datetime
    for fmt in (
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m-%d-%Y",
        "%m-%d-%y",
        "%d/%m/%Y",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text  # Fall through; QB will return a date-format error if invalid.


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

        qb_details = _quickbooks_process_details()
        qb_already_running = len(qb_details) > 0
        log.info("connect(): QB process detected = %s (count=%d)", qb_already_running, len(qb_details))
        app_elevation = _current_process_elevation()

        # Hard safety gate: do not call BeginSession in ambiguous process states.
        # This avoids any chance of side effects when QuickBooks has multiple
        # running processes/windows with mixed contexts.
        if not qb_already_running:
            context_note = _runtime_context_note()
            log.error("connect(): QuickBooks is not running; attach-only mode, no launch fallback")
            log.error("connect(): context diagnostics: %s", context_note)
            self.close()
            raise RuntimeError(
                "QuickBooks Desktop is not running. This app is configured to attach only and "
                "will never launch/open QuickBooks automatically.\n\n"
                f"Context: {context_note}"
            )

        # Real-instance check: refuse only when there are genuinely multiple
        # *user-facing* QuickBooks UI processes (DDE Server helpers were
        # already filtered out in _quickbooks_process_details). One main
        # company-file window + the SDK helper is the normal steady state.
        if len(qb_details) > 1:
            context_note = _runtime_context_note()
            log.error("connect(): multiple user-facing QuickBooks UI processes detected; refusing attach")
            log.error("connect(): context diagnostics: %s", context_note)
            self.close()
            raise RuntimeError(
                "Multiple QuickBooks windows are open. This app will not attempt to connect "
                "until only one QuickBooks instance remains.\n\n"
                "Close all QuickBooks windows, reopen one company file, then retry.\n\n"
                f"Context: {context_note}"
            )

        # Elevation is reported but no longer used to refuse the attach. The
        # working v0.9.708 attach succeeded in non-elevated mode; elevation is
        # left as informational diagnostic so the log still captures it when an
        # attach actually fails later.
        if app_elevation == "elevated":
            log.warning(
                "connect(): app is elevated; QuickBooks may surface a UAC permission prompt or "
                "spawn an additional window. Proceeding with attach attempt anyway."
            )
        mismatched = [
            row for row in qb_details if row.get("elevation") not in ("unknown", app_elevation)
        ]
        if mismatched:
            log.warning(
                "connect(): QuickBooks/app UAC elevation mismatch (%s vs %s). Attach may fail.",
                mismatched[0].get("elevation"),
                app_elevation,
            )

        if _has_no_company_open_window():
            context_note = _runtime_context_note()
            log.error("connect(): QuickBooks is at 'No Company Open'; refusing BeginSession")
            log.error("connect(): context diagnostics: %s", context_note)
            self.close()
            raise RuntimeError(
                "QuickBooks is currently at 'No Company Open'. This app will not attempt "
                "BeginSession in this state to avoid opening/spawning additional QuickBooks "
                "windows.\n\nOpen the target company file in QuickBooks first, then retry.\n\n"
                f"Context: {context_note}"
            )

        if _has_secondary_window_title():
            context_note = _runtime_context_note()
            log.error("connect(): QuickBooks has a Secondary window title; refusing BeginSession")
            log.error("connect(): context diagnostics: %s", context_note)
            self.close()
            raise RuntimeError(
                "QuickBooks shows a Secondary window/session. Connection is blocked to avoid "
                "triggering additional QuickBooks windows.\n\n"
                "Close all QuickBooks windows/processes, then open only one company window and retry.\n\n"
                f"Context: {context_note}"
            )

        # Single safe attach attempt only. Do not cascade retries after one SDK
        # failure, because retries can worsen unstable QB/network state.
        attach_error = None
        try:
            log.debug("connect(): BeginSession(path='', mode=2)")
            self._ticket = self._rp.BeginSession("", 2)
            log.info("connect(): attached to running session (mode=2)")
            return
        except Exception as exc:
            attach_error = exc
            log.warning("connect(): BeginSession(path='', mode=2) failed — %s", exc)

        context_note = _runtime_context_note()
        log.error("connect(): context diagnostics: %s", context_note)
        self.close()
        if _is_company_session_lost_error(attach_error):
            raise RuntimeError(
                "QuickBooks lost access to the company session while this connection attempt was running "
                "(SDK errors -2147220478 / -2147220469). This usually indicates a QuickBooks host/network "
                "company-file instability, not an app launch fallback.\n\n"
                "In QuickBooks, resolve the connection-lost state first, reopen the company file, "
                "then retry once stable.\n\n"
                f"Attach details: {attach_error}\nContext: {context_note}"
            )
        raise RuntimeError(
            "QuickBooks Desktop is running, but this app could not attach to the open "
            "company file. This app will not launch/open QuickBooks as fallback.\n\n"
            "Ensure QuickBooks is fully logged into the company file and running under the "
            "same Windows user and UAC elevation level as this app.\n\n"
            f"Attach details: {attach_error}"
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
        page = 0

        while True:
            iterator_attr = f' iterator="{iterator}"'
            iterator_id_attr = f' iteratorID="{iterator_id}"' if iterator_id else ""
            # MaxReturned=1000 reduces COM round-trips over network shares. The
            # SDK clamps to its internal max if a smaller value applies.
            # RefNumberRangeFilter with FromRefNumber=MIN_SALES_ORDER_NUMBER
            # skips the long historical tail. RefNumber comparisons in QBXML
            # are lexicographic, which is fine here because ref numbers in this
            # company file are fixed-width numeric strings.
            # The filter is only valid on the initial request — once iterating
            # with iterator="Continue", QB rejects re-specified filters.
            filter_xml = (
                f"<RefNumberRangeFilter><FromRefNumber>{MIN_SALES_ORDER_NUMBER}</FromRefNumber></RefNumberRangeFilter>"
                if iterator == "Start"
                else ""
            )
            query_xml = f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <SalesOrderQueryRq{iterator_attr}{iterator_id_attr}>
      <MaxReturned>1000</MaxReturned>
      {filter_xml}
      <IncludeRetElement>RefNumber</IncludeRetElement>
    </SalesOrderQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""
            page += 1
            log.debug("_query_ref_numbers: requesting page %d (so far %d refs)", page, len(ref_numbers))
            response = self._process(query_xml)
            root = ET.fromstring(response)

            for ref in root.findall(".//SalesOrderRet/RefNumber"):
                if ref.text:
                    ref_numbers.append(ref.text.strip())

            query_rs = root.find(".//SalesOrderQueryRs")
            if query_rs is None:
                log.debug("_query_ref_numbers: no QueryRs element on page %d; stopping", page)
                break
            remaining = int(query_rs.attrib.get("iteratorRemainingCount", "0"))
            log.debug(
                "_query_ref_numbers: page %d returned %d total refs so far, %d remaining",
                page,
                len(ref_numbers),
                remaining,
            )
            if remaining <= 0:
                break
            iterator = "Continue"
            iterator_id = query_rs.attrib.get("iteratorID", "")
            if not iterator_id:
                log.debug("_query_ref_numbers: missing iteratorID on page %d; stopping", page)
                break

        log.info("_query_ref_numbers: completed, %d refs across %d page(s)", len(ref_numbers), page)
        return ref_numbers

    def get_next_sales_order_number(self) -> str:
        self.connect()
        try:
            refs = self._query_ref_numbers()
            numeric_values = []
            for ref in refs:
                if re.fullmatch(r"\d+", ref):
                    numeric_values.append(int(ref))
            # _query_ref_numbers is filtered to RefNumber >= MIN_SALES_ORDER_NUMBER,
            # so when nothing matches we start the count there rather than at 1.
            min_floor = int(MIN_SALES_ORDER_NUMBER)
            if not numeric_values:
                return str(min_floor)
            highest = max(numeric_values)
            return str(max(highest, min_floor - 1) + 1)
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

            # QBXML schema requires SalesOrderAdd children in a strict order:
            # CustomerRef, ClassRef, ARAccountRef, TemplateRef, TxnDate,
            # RefNumber, BillAddress, ShipAddress, PONumber, TermsRef, DueDate,
            # SalesRep, FOB, ShipDate, ShipMethodRef, IsManuallyClosed, Memo,
            # ..., SalesOrderLineAdd. Out-of-order children cause QB to return
            # -2147220480 "QuickBooks found an error when parsing the provided
            # XML text stream." Build the optional ones conditionally so empty
            # fields don't ship as empty elements (also a schema risk).
            txn_date_iso = _to_qb_date(txn_date)
            due_date_iso = _to_qb_date(due_date)
            log.debug(
                "upload_sales_order: customer=%r so=%r txn=%s due=%s terms=%r ship=%r tax=%r lines=%d",
                customer_name,
                sales_order_no,
                txn_date_iso,
                due_date_iso,
                terms,
                shipping_method,
                tax_code,
                len(line_xml),
            )
            parts: list[str] = [
                f"<CustomerRef><FullName>{html.escape(customer_name)}</FullName></CustomerRef>",
            ]
            if txn_date_iso:
                parts.append(f"<TxnDate>{txn_date_iso}</TxnDate>")
            parts.append(f"<RefNumber>{html.escape(sales_order_no)}</RefNumber>")
            if terms:
                parts.append(f"<TermsRef><FullName>{html.escape(terms)}</FullName></TermsRef>")
            if due_date_iso:
                parts.append(f"<DueDate>{due_date_iso}</DueDate>")
            if shipping_method:
                parts.append(f"<ShipMethodRef><FullName>{html.escape(shipping_method)}</FullName></ShipMethodRef>")
            if memo:
                parts.append(f"<Memo>{html.escape(memo)}</Memo>")
            parts.append("".join(line_xml))

            request_xml = f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <SalesOrderAddRq>
      <SalesOrderAdd>
        {''.join(parts)}
      </SalesOrderAdd>
    </SalesOrderAddRq>
  </QBXMLMsgsRq>
</QBXML>"""
            try:
                response = self._process(request_xml)
            except Exception:
                # Surface the XML we attempted to send so the log shows exactly
                # what the SDK rejected. Bounded length to keep the log readable.
                preview = request_xml if len(request_xml) <= 8000 else request_xml[:8000] + "\n... [truncated]"
                log.error("upload_sales_order: ProcessRequest raised; XML sent (preview):\n%s", preview)
                raise
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
