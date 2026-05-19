from __future__ import annotations

import ctypes
import html
import logging
import os
import re
import subprocess
import traceback
import xml.etree.ElementTree as ET
from typing import Callable, Iterable, List, Optional

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


def _clean(text: str) -> str:
    """Make a string safe to embed as the text content of a QBXML element.

    Drops ASCII control chars (except tab/newline/CR) and any non-ASCII
    codepoints, normalizes line breaks to spaces, trims, then html-escapes
    so `<`, `>`, `&`, quotes are safe.

    See upload_sales_order for the long explanation of why non-ASCII is
    stripped (pywin32 marshalling vs. QB's parser).
    """
    s = str(text or "")
    s = "".join(
        ch for ch in s
        if ch in "\t\n\r" or 0x20 <= ord(ch) <= 0x7E
    )
    s = s.replace("\r", " ").replace("\n", " ").strip()
    return html.escape(s)


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
        # Populated by upload_sales_order with metadata about each item the
        # upload had to auto-create. The caller reads this after the call
        # to show a report.
        self.last_created_items: list[dict] = []

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

        # Process-detection via tasklist is unreliable: a non-elevated app
        # cannot see processes from other users or elevated processes, which
        # produced false "QB process detected = False" results. The SDK
        # itself is the authoritative source — BeginSession("") is *safe*
        # (it can only attach or fail, never spawn QB) so we just call it
        # directly and let the SDK tell us the truth. The tasklist scan is
        # kept only for the context-note diagnostics on errors.
        app_elevation = _current_process_elevation()
        log.info("connect(): app elevation=%s", app_elevation)

        # Step 1: try to attach to a running session. This call cannot launch
        # QuickBooks. If QB is open with a company file, this succeeds.
        attach_error = None
        try:
            log.debug("connect(): BeginSession(path='', mode=2)")
            self._ticket = self._rp.BeginSession("", 2)
            log.info("connect(): attached to running session (mode=2)")
            return
        except Exception as exc:
            attach_error = exc
            log.warning("connect(): BeginSession(path='', mode=2) failed — %s", exc)

        # Step 2: the attach failed. Decode the error and decide whether to
        # fall back to launching QB via the saved .QBW path.
        attach_text = str(attach_error).lower()
        no_file_open_error = (
            "-2147220457" in attach_text
            or "data file is not open" in attach_text
            or "must include the name of the data file" in attach_text
        )

        if no_file_open_error and self.company_file_path:
            # No live session — let the SDK launch QuickBooks with the saved
            # path. This is the legitimate "open QB for me" flow.
            log.info(
                "connect(): no current session; launching via saved path %r",
                self.company_file_path,
            )
            launch_error = None
            for open_mode in (2, 0, 1):
                try:
                    log.debug(
                        "connect(): BeginSession(path=%r, mode=%d)",
                        self.company_file_path,
                        open_mode,
                    )
                    self._ticket = self._rp.BeginSession(self.company_file_path, open_mode)
                    log.info("connect(): launched QB via saved path (mode=%d)", open_mode)
                    return
                except Exception as exc:
                    log.warning(
                        "connect(): BeginSession(path=%r, mode=%d) failed — %s",
                        self.company_file_path,
                        open_mode,
                        exc,
                    )
                    if launch_error is None:
                        launch_error = exc
            context_note = _runtime_context_note()
            log.error("connect(): all path-based launch modes failed: %s", launch_error)
            log.error("connect(): context diagnostics: %s", context_note)
            self.close()
            raise RuntimeError(
                "Could not launch QuickBooks Desktop using the saved .QBW path. "
                "Open QuickBooks manually with your company file, then click Connect again.\n\n"
                f"Attach details: {launch_error}\nContext: {context_note}"
            )

        # Step 3: the attach failed for a reason that isn't "no file open",
        # OR there's no saved path. Surface the SDK error to the user.
        context_note = _runtime_context_note()
        log.error("connect(): context diagnostics: %s", context_note)
        self.close()

        if no_file_open_error and not self.company_file_path:
            raise RuntimeError(
                "QuickBooks Desktop has no company file open, and no QuickBooks Company File "
                "(.QBW) path is set in this app's Configuration card.\n\n"
                "Either open QuickBooks with your company file first, or set the .QBW path in "
                "Configuration so this app can open the file for you.\n\n"
                f"Context: {context_note}"
            )

        if _is_company_session_lost_error(attach_error):
            raise RuntimeError(
                "QuickBooks lost access to the company session while this connection attempt was running "
                "(SDK errors -2147220478 / -2147220469). This usually indicates a QuickBooks host/network "
                "company-file instability.\n\n"
                "In QuickBooks, resolve the connection-lost state first, reopen the company file, "
                "then retry once stable.\n\n"
                f"Attach details: {attach_error}\nContext: {context_note}"
            )

        raise RuntimeError(
            "Could not attach to QuickBooks Desktop. Make sure QuickBooks is open with your "
            "company file fully loaded, and that this app is running under the same Windows "
            "user (and same UAC elevation level) as QuickBooks.\n\n"
            f"Attach details: {attach_error}\nContext: {context_note}"
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

    def _customer_has_default_tax_item(self, customer_name: str) -> tuple[bool, str]:
        """Check whether a customer already has a default ItemSalesTaxRef.

        Returns (has_default, edit_sequence). edit_sequence is needed if a
        CustomerMod is required afterwards. has_default is True when the
        customer record already has any tax item attached, so we don't
        clobber it.
        """
        name = _clean(customer_name)
        if not name:
            return False, ""
        query_xml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="continueOnError">
    <CustomerQueryRq>
      <FullName>{name}</FullName>
      <IncludeRetElement>ListID</IncludeRetElement>
      <IncludeRetElement>EditSequence</IncludeRetElement>
      <IncludeRetElement>ItemSalesTaxRef</IncludeRetElement>
    </CustomerQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""
        try:
            response = self._process(query_xml)
        except Exception:
            log.exception("_customer_has_default_tax_item: query failed for %r", customer_name)
            return False, ""
        try:
            root = ET.fromstring(response)
        except ET.ParseError:
            log.exception("_customer_has_default_tax_item: bad response XML")
            return False, ""
        cust = root.find(".//CustomerRet")
        if cust is None:
            return False, ""
        edit_seq = cust.findtext("EditSequence", default="")
        tax_full = cust.findtext("ItemSalesTaxRef/FullName", default="").strip()
        return bool(tax_full), edit_seq

    def _set_customer_default_tax_item(
        self,
        customer_name: str,
        edit_sequence: str,
        tax_item: str,
    ) -> tuple[bool, str]:
        """CustomerMod: set the customer's default ItemSalesTaxRef.

        Returns (ok, message). The caller is expected to have already
        confirmed that no default is set, via _customer_has_default_tax_item.
        """
        name = _clean(customer_name)
        seq = _clean(edit_sequence)
        item = _clean(tax_item)
        if not name or not seq or not item:
            return False, "missing customer name, edit sequence, or tax item"
        # CustomerModRq requires the EditSequence on the immediate child.
        # We use ListID-less FullName lookup; QB resolves the customer via
        # FullName + EditSequence.
        # Find the customer's ListID first; CustomerMod requires it.
        list_id_xml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="continueOnError">
    <CustomerQueryRq>
      <FullName>{name}</FullName>
      <IncludeRetElement>ListID</IncludeRetElement>
    </CustomerQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""
        try:
            response = self._process(list_id_xml)
            list_id = ET.fromstring(response).findtext(".//CustomerRet/ListID", default="").strip()
        except Exception:
            log.exception("_set_customer_default_tax_item: ListID lookup failed")
            return False, "could not look up customer ListID"
        if not list_id:
            return False, "customer not found"
        mod_xml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="continueOnError">
    <CustomerModRq>
      <CustomerMod>
        <ListID>{list_id}</ListID>
        <EditSequence>{seq}</EditSequence>
        <ItemSalesTaxRef><FullName>{item}</FullName></ItemSalesTaxRef>
      </CustomerMod>
    </CustomerModRq>
  </QBXMLMsgsRq>
</QBXML>"""
        try:
            response = self._process(mod_xml)
        except Exception as exc:
            log.exception("_set_customer_default_tax_item: ProcessRequest raised")
            return False, f"SDK error: {exc}"
        try:
            root = ET.fromstring(response)
        except ET.ParseError as exc:
            return False, f"bad response XML: {exc}"
        rs = root.find(".//CustomerModRs")
        if rs is None:
            return False, "no CustomerModRs in response"
        status_code = rs.attrib.get("statusCode", "")
        status_message = rs.attrib.get("statusMessage", "")
        if status_code != "0":
            return False, f"{status_code}: {status_message}"
        return True, "ok"

    def _ensure_customer_tax_item(self, customer_name: str, tax_item: str) -> None:
        """If `customer_name` has no default ItemSalesTaxRef, set it to
        `tax_item`. Best-effort: failures are logged but not raised, so the
        sales-order upload still gets attempted. The caller (upload_sales_order)
        will surface QB's 3180 error if it ultimately matters.
        """
        if not customer_name or not tax_item:
            return
        has_default, edit_seq = self._customer_has_default_tax_item(customer_name)
        if has_default:
            log.debug(
                "_ensure_customer_tax_item: %r already has a default tax item, leaving alone",
                customer_name,
            )
            return
        if not edit_seq:
            log.warning(
                "_ensure_customer_tax_item: could not get EditSequence for %r; skipping CustomerMod",
                customer_name,
            )
            return
        log.info(
            "_ensure_customer_tax_item: setting %r default ItemSalesTaxRef to %r",
            customer_name,
            tax_item,
        )
        ok, msg = self._set_customer_default_tax_item(customer_name, edit_seq, tax_item)
        if not ok:
            log.warning(
                "_ensure_customer_tax_item: CustomerMod failed for %r: %s",
                customer_name,
                msg,
            )

    def _query_existing_items(self, skus: Iterable[str]) -> set[str]:
        """Return the subset of `skus` that exist in the QuickBooks Item list.

        Batches FullName filters into one ItemQueryRq per chunk to keep the
        number of round-trips low. The SDK accepts multiple <FullName>
        elements in a single ItemQueryRq.

        Caller must already be inside an active BeginSession (connect()).
        """
        unique = sorted({s for s in (str(x or "").strip() for x in skus) if s})
        if not unique:
            return set()
        existing: set[str] = set()
        chunk_size = 50
        for i in range(0, len(unique), chunk_size):
            chunk = unique[i : i + chunk_size]
            fullnames_xml = "".join(f"<FullName>{html.escape(s)}</FullName>" for s in chunk)
            query_xml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="continueOnError">
    <ItemQueryRq>
      {fullnames_xml}
      <IncludeRetElement>Name</IncludeRetElement>
      <IncludeRetElement>FullName</IncludeRetElement>
    </ItemQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""
            try:
                response = self._process(query_xml)
            except Exception:
                log.exception("_query_existing_items: ItemQuery chunk failed; assuming all missing")
                continue
            try:
                root = ET.fromstring(response)
            except ET.ParseError:
                log.exception("_query_existing_items: could not parse ItemQuery response")
                continue
            for full in root.iter("FullName"):
                if full.text:
                    existing.add(full.text.strip())
            for name in root.iter("Name"):
                if name.text:
                    existing.add(name.text.strip())
            log.debug(
                "_query_existing_items: chunk %d (%d skus) — %d found so far",
                i // chunk_size + 1,
                len(chunk),
                len(existing),
            )
        log.info(
            "_query_existing_items: %d of %d requested SKUs exist in QuickBooks",
            len(existing & set(unique)),
            len(unique),
        )
        return existing

    def _add_non_inventory_item(
        self,
        name: str,
        desc: str,
        income_account: str,
        cost: float = 0.0,
        expense_account: str = "",
    ) -> tuple[bool, str]:
        """Create a Non-Inventory item in QuickBooks.

        Returns (success, status_message). Caller is already inside an active
        BeginSession.
        """
        def _ascii(text: str) -> str:
            s = str(text or "")
            s = "".join(
                ch for ch in s
                if ch in "\t\n\r" or 0x20 <= ord(ch) <= 0x7E
            )
            s = s.replace("\r", " ").replace("\n", " ").strip()
            return html.escape(s)

        name_clean = _ascii(name)
        desc_clean = _ascii(desc)
        income_clean = _ascii(income_account)
        expense_clean = _ascii(expense_account)
        if not name_clean or not income_clean:
            return False, "missing name or income account"
        # When an expense account is provided we use the SalesAndPurchase
        # block, which lets us stamp PurchaseCost + ExpenseAccountRef onto
        # the new item (so QuickBooks shows the cost in Purchase Information
        # and posts COGS correctly). Otherwise we fall back to the simpler
        # SalesOrPurchase block (sales-side only).
        if expense_clean:
            cost_xml = ""
            try:
                cost_value = float(cost or 0.0)
            except Exception:
                cost_value = 0.0
            if cost_value > 0:
                cost_xml = f"<PurchaseCost>{cost_value}</PurchaseCost>"
            body = f"""        <SalesAndPurchase>
          <SalesDesc>{desc_clean}</SalesDesc>
          <IncomeAccountRef><FullName>{income_clean}</FullName></IncomeAccountRef>
          <PurchaseDesc>{desc_clean}</PurchaseDesc>
          {cost_xml}
          <ExpenseAccountRef><FullName>{expense_clean}</FullName></ExpenseAccountRef>
        </SalesAndPurchase>"""
        else:
            body = f"""        <SalesOrPurchase>
          <Desc>{desc_clean}</Desc>
          <AccountRef><FullName>{income_clean}</FullName></AccountRef>
        </SalesOrPurchase>"""
        # QBXML schema: each Add must be wrapped in a corresponding *Rq.
        # ItemNonInventoryAdd requires Name + IsActive.
        req_xml = f"""<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="continueOnError">
    <ItemNonInventoryAddRq>
      <ItemNonInventoryAdd>
        <Name>{name_clean}</Name>
        <IsActive>true</IsActive>
{body}
      </ItemNonInventoryAdd>
    </ItemNonInventoryAddRq>
  </QBXMLMsgsRq>
</QBXML>"""
        try:
            response = self._process(req_xml)
        except Exception as exc:
            log.exception("_add_non_inventory_item: ProcessRequest raised for %r", name)
            return False, f"SDK error: {exc}"
        try:
            root = ET.fromstring(response)
        except ET.ParseError as exc:
            return False, f"could not parse response: {exc}"
        rs = root.find(".//ItemNonInventoryAddRs")
        if rs is None:
            return False, "no ItemNonInventoryAddRs in response"
        status_code = rs.attrib.get("statusCode", "")
        status_message = rs.attrib.get("statusMessage", "")
        if status_code != "0":
            return False, f"{status_code}: {status_message}"
        return True, "ok"

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
        fallback_item: str = "",
        income_account: str = "",
        group_by_room: bool = False,
        sales_tax_item: str = "",
        expense_account: str = "",
        progress_cb: Optional[Callable[[str], None]] = None,
    ) -> str:
        # v1.037: progress_cb fires at each meaningful step so the UI can
        # show line-by-line activity in a streaming log. Callback must be
        # cheap & non-blocking (caller marshals to the Tk main loop).
        def _progress(msg: str) -> None:
            if progress_cb is None:
                return
            try:
                progress_cb(msg)
            except Exception:
                log.exception("upload_sales_order: progress_cb raised; continuing")

        self.last_created_items = []
        _progress("Connecting to QuickBooks...")
        self.connect()
        try:
            # Make sure the customer has a default sales-tax item so the
            # SalesOrderAdd doesn't trip QB error 3180 ("Transaction Sales
            # Tax field cannot be left blank"). QBXML 13.0 doesn't allow an
            # ItemSalesTaxRef directly on SalesOrderAdd, so we set it on the
            # customer instead. Best-effort -- logged failures don't block
            # the upload attempt.
            if sales_tax_item:
                _progress(f"Verifying sales-tax setup for {customer_name}...")
                self._ensure_customer_tax_item(customer_name, sales_tax_item)

            # _clean is module-level (see top of this file) so the customer-
            # tax-item helpers can use it too. The local redefinition that
            # used to live here was removed.

            # Materialize lines once so we can pre-flight the SKU check and
            # still iterate them when building the XML.
            line_rows: list[dict] = []
            for idx, line in enumerate(lines, start=1):
                try:
                    line_cost = 0.0
                    if isinstance(line, dict):
                        try:
                            line_cost = float(line.get("Cost", 0) or 0)
                        except Exception:
                            line_cost = 0.0
                    line_rows.append({
                        "idx": idx,
                        "sku_raw": str(line["Product/Service"] or "").strip(),
                        "sku": _clean(line["Product/Service"]),
                        "desc": _clean(line["Product/Service Description"]),
                        "qty": float(line["Product/Service Quantity"]),
                        "rate": float(line["Product/Service Rate"]),
                        "cost": line_cost,
                        "room": str(line.get("Room", "") or "").strip() if isinstance(line, dict) else "",
                    })
                except Exception as exc:
                    raise RuntimeError(
                        f"Line {idx} has invalid data ({exc}). Check the SKU, "
                        f"description, quantity, and rate for that row before uploading."
                    )

            # Pre-flight: which SKUs actually exist in the QuickBooks item list?
            # Missing items would otherwise fail the whole SalesOrderAdd with
            # error 3140. Two recovery modes, preferred in this order:
            #   1. If a Default Income Account is configured, auto-create each
            #      missing SKU as a Non-Inventory item under that account
            #      using the description from the source data. This keeps each
            #      SKU unique in the QuickBooks item list.
            #   2. Otherwise, if a Fallback Item is set, substitute that
            #      ItemRef in place and stamp the original SKU into the line
            #      description.
            #   3. If neither is set, fail with a clear message listing the
            #      missing SKUs.
            raw_skus = [row["sku_raw"] for row in line_rows if row["sku_raw"]]
            _progress(f"Checking which of {len(set(raw_skus))} SKU(s) exist in QuickBooks...")
            existing_items = self._query_existing_items(raw_skus)
            missing_skus = sorted({s for s in raw_skus if s not in existing_items})
            fallback_clean = _clean(fallback_item)
            income_clean = (income_account or "").strip()
            if missing_skus:
                if income_clean:
                    log.info(
                        "upload_sales_order: auto-creating %d missing item(s) under income account %r",
                        len(missing_skus),
                        income_clean,
                    )
                    # Build SKU -> (description, cost) maps from the input
                    # lines so each newly-created item carries the right
                    # default description and PurchaseCost.
                    desc_by_sku: dict[str, str] = {}
                    cost_by_sku: dict[str, float] = {}
                    for row in line_rows:
                        s = row["sku_raw"]
                        if s and s not in desc_by_sku:
                            desc_by_sku[s] = row.get("desc", "")
                            cost_by_sku[s] = float(row.get("cost", 0) or 0)
                    created: list[str] = []
                    failed: list[tuple[str, str]] = []
                    _progress(f"Auto-creating {len(missing_skus)} missing item(s) in QuickBooks...")
                    for ci, sku in enumerate(missing_skus, start=1):
                        line_desc = desc_by_sku.get(sku, "")
                        line_cost = cost_by_sku.get(sku, 0.0)
                        _progress(f"  Creating item {ci}/{len(missing_skus)}: {sku}")
                        ok, msg = self._add_non_inventory_item(
                            name=sku,
                            desc=line_desc,
                            income_account=income_clean,
                            cost=line_cost,
                            expense_account=expense_account,
                        )
                        if ok:
                            created.append(sku)
                            existing_items.add(sku)
                            self.last_created_items.append({
                                "sku": sku,
                                "description": line_desc,
                                "income_account": income_clean,
                                "expense_account": expense_account,
                                "cost": line_cost,
                            })
                        else:
                            failed.append((sku, msg))
                            log.error(
                                "upload_sales_order: could not auto-create item %r: %s",
                                sku,
                                msg,
                            )
                    log.info(
                        "upload_sales_order: created %d / %d missing item(s); %d failed",
                        len(created),
                        len(missing_skus),
                        len(failed),
                    )
                    if failed:
                        first = "; ".join(f"{s} -> {m}" for s, m in failed[:5])
                        raise RuntimeError(
                            f"Could not auto-create {len(failed)} item(s) in QuickBooks. "
                            f"Verify that the Default Income Account '{income_clean}' "
                            f"exists in QuickBooks's Chart of Accounts and is an Income-type "
                            f"account.\n\nFirst failures: {first}"
                        )
                    # After auto-create the missing list is now empty for the
                    # purposes of substitution below.
                    missing_skus = []
                elif fallback_clean:
                    if fallback_item.strip() not in self._query_existing_items([fallback_item.strip()]):
                        raise RuntimeError(
                            f"Fallback Item {fallback_item!r} does not exist in QuickBooks. "
                            f"Set Fallback Item in the Configuration card to a real item name."
                        )
                    log.warning(
                        "upload_sales_order: substituting %d missing SKU(s) with fallback %r: %s",
                        len(missing_skus),
                        fallback_item.strip(),
                        ", ".join(missing_skus[:20]),
                    )
                else:
                    log.error(
                        "upload_sales_order: %d SKU(s) not in QuickBooks item list and no recovery configured: %s",
                        len(missing_skus),
                        ", ".join(missing_skus[:20]),
                    )
                    raise RuntimeError(
                        f"{len(missing_skus)} item(s) in this sales order do not exist in QuickBooks:\n\n  "
                        + ", ".join(missing_skus[:20])
                        + ("\n  ..." if len(missing_skus) > 20 else "")
                        + "\n\nFix this one of two ways in the Configuration card:\n"
                        "  - Set 'Default Income Account' to the name of an income account in "
                        "your QuickBooks Chart of Accounts (e.g. 'Sales' or 'Sales Income'). "
                        "The app will then auto-create each missing SKU as a Non-Inventory "
                        "item under that account, preserving each unique SKU.\n"
                        "  - OR set 'Fallback Item' to the name of an existing generic "
                        "QuickBooks item, and missing SKUs will be lumped under it with the "
                        "original SKU stamped into the line description."
                    )

            def _product_line_xml(row: dict) -> str:
                sku = row["sku"]
                sku_raw = row["sku_raw"]
                desc = row["desc"]
                if sku_raw and sku_raw not in existing_items and fallback_clean:
                    note = _clean(f"[{sku_raw}] ")
                    desc = (note + desc).strip()
                    sku = fallback_clean
                return f"""
      <SalesOrderLineAdd>
        <ItemRef><FullName>{sku}</FullName></ItemRef>
        <Desc>{desc}</Desc>
        <Quantity>{row["qty"]}</Quantity>
        <Rate>{row["rate"]}</Rate>
        <SalesTaxCodeRef><FullName>{_clean(tax_code)}</FullName></SalesTaxCodeRef>
      </SalesOrderLineAdd>"""

            def _header_line_xml(label: str) -> str:
                # Description-only line: no ItemRef, no Quantity, no Rate.
                # QuickBooks Desktop accepts this for section headers when
                # only the Desc element is provided.
                return f"""
      <SalesOrderLineAdd>
        <Desc>{_clean(label)}</Desc>
      </SalesOrderLineAdd>"""

            def _blank_line_xml() -> str:
                return """
      <SalesOrderLineAdd>
        <Desc></Desc>
      </SalesOrderLineAdd>"""

            # Stream each line into the progress log so the user sees what
            # they're shipping to QuickBooks. The actual SalesOrderAdd is a
            # single atomic request below — the per-line callbacks happen as
            # the XML is built, not as QB processes each row.
            _progress(f"Building sales-order XML for {len(line_rows)} line(s)...")
            for row in line_rows:
                try:
                    qty = row["qty"]
                    rate = row["rate"]
                    sku_label = row["sku_raw"] or "(blank SKU)"
                    _progress(f"  Line {row['idx']}/{len(line_rows)}: {sku_label}  qty={qty:g}  rate=${rate:,.2f}")
                except Exception:
                    pass

            line_xml: list[str] = []
            if group_by_room:
                # Walk line_rows in order; whenever the Room value changes
                # (including to/from blank), close the previous group with two
                # blank lines and start a new group with a **ROOM** header.
                # Blank-room groups are labeled "Unnamed Room" or
                # "Unnamed Room N" when there are multiple of them.
                raw_groups: list[tuple[str, list[dict]]] = []
                current_room: str | None = None
                current_lines: list[dict] = []
                for row in line_rows:
                    room_val = row.get("room", "")
                    if current_room is None:
                        current_room = room_val
                        current_lines = [row]
                    elif room_val == current_room:
                        current_lines.append(row)
                    else:
                        raw_groups.append((current_room, current_lines))
                        current_room = room_val
                        current_lines = [row]
                if current_lines:
                    raw_groups.append((current_room or "", current_lines))

                unnamed_count = sum(1 for r, _ in raw_groups if not r)
                unnamed_idx = 0
                log.info(
                    "upload_sales_order: room grouping enabled - %d group(s)",
                    len(raw_groups),
                )
                for gi, (room, group_lines) in enumerate(raw_groups):
                    if room:
                        label = f"**{room.upper()}**"
                    elif unnamed_count > 1:
                        unnamed_idx += 1
                        label = f"**UNNAMED ROOM {unnamed_idx}**"
                    else:
                        label = "**UNNAMED ROOM**"
                    if gi > 0:
                        # Two blank separator lines between groups.
                        line_xml.append(_blank_line_xml())
                        line_xml.append(_blank_line_xml())
                    line_xml.append(_header_line_xml(label))
                    for row in group_lines:
                        line_xml.append(_product_line_xml(row))
            else:
                for row in line_rows:
                    line_xml.append(_product_line_xml(row))

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
                f"<CustomerRef><FullName>{_clean(customer_name)}</FullName></CustomerRef>",
            ]
            if txn_date_iso:
                parts.append(f"<TxnDate>{txn_date_iso}</TxnDate>")
            parts.append(f"<RefNumber>{_clean(sales_order_no)}</RefNumber>")
            if terms:
                parts.append(f"<TermsRef><FullName>{_clean(terms)}</FullName></TermsRef>")
            if due_date_iso:
                parts.append(f"<DueDate>{due_date_iso}</DueDate>")
            if shipping_method:
                parts.append(f"<ShipMethodRef><FullName>{_clean(shipping_method)}</FullName></ShipMethodRef>")
            if memo:
                parts.append(f"<Memo>{_clean(memo)}</Memo>")
            # CustomerSalesTaxCodeRef sets the header-level "Tax" code on the
            # sales order so QuickBooks doesn't default it to None. The line-
            # level SalesTaxCodeRef on each SalesOrderLineAdd marks the line
            # as taxable; without the header code, QB still treats the order
            # as non-taxable and the user has to flip it manually.
            if tax_code:
                parts.append(
                    f"<CustomerSalesTaxCodeRef><FullName>{_clean(tax_code)}</FullName></CustomerSalesTaxCodeRef>"
                )
            # Note: ItemSalesTaxRef is NOT a valid child of SalesOrderAdd in
            # QBXML 13.0 -- emitting it here triggers a parse error
            # (-2147220480). The header tax rate is instead inherited from
            # the customer's default tax item, which we ensure is set via
            # _ensure_customer_tax_item() above before this XML is built.
            parts.append("".join(line_xml))

            # Intentionally NO encoding="..." on the XML declaration.
            # pywin32 marshals Python str -> COM BSTR -> UTF-16, but if the
            # declaration says "utf-8" QuickBooks's parser refuses the
            # mismatch with the generic -2147220480 "parsing error". Omitting
            # the encoding lets QBXMLRP2 use whatever encoding COM delivered.
            request_xml = f"""<?xml version="1.0"?>
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

            # Parse the XML locally before sending. If it doesn't parse here,
            # QuickBooks can't possibly parse it — and ET tells us exactly
            # which character/line is broken, which is far more useful than
            # QB's "found an error when parsing" generic error.
            try:
                ET.fromstring(request_xml)
            except ET.ParseError as parse_exc:
                log.error(
                    "upload_sales_order: locally-built XML is not well-formed (%s). "
                    "This means a SKU or description likely contains characters that "
                    "broke escaping.",
                    parse_exc,
                )
                log.error("upload_sales_order: full XML being rejected locally:\n%s", request_xml)
                raise RuntimeError(
                    "The sales order XML built from this preview is not valid XML "
                    "before it even reaches QuickBooks. A SKU or description "
                    f"contains characters that broke escaping. Local parse error: {parse_exc}.\n\n"
                    f"Full XML written to the log."
                )

            _progress(f"Submitting sales order to QuickBooks ({len(line_rows)} line(s))...")
            try:
                response = self._process(request_xml)
            except Exception:
                # On any upload failure, dump the FULL XML to the log so the
                # rejected row can be identified. Logs are rotated at 512 KB
                # so even multi-hundred-line sales orders fit.
                log.error("upload_sales_order: ProcessRequest raised; full XML sent:\n%s", request_xml)
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
            _progress(f"Sales order #{ref} created in QuickBooks.")
            return f"Uploaded Sales Order #{ref} (TxnID: {txn_id})"
        finally:
            self.close()
