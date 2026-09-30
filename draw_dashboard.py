"""
Construction Draw Tracker  (Streamlit)
======================================
Salesforce login (same Connected App as the AM slide app), then:

  • Turn-Time      business days from Full Draw Package Received -> Wire Date for every
                   construction draw wired in the chosen period, with the borrower-side
                   (request -> package) time, by year / quarter / month / coordinator,
                   draw detail, and a one-click Excel workbook.
  • Open Pipeline  draws not yet wired, by status, oldest first (for the Thursday call).
  • Draw Lookup    one property / advance: every milestone date and its turn-times.

Data: Advance__c, Record Type "Construction Advance". Field names come from the
Advance__c describe() we ran; optional extras (amount, loan #, product type, payoff date,
Land Gorilla loan ID) are only selected if the org has them.

Secrets (.streamlit/secrets.toml or Streamlit Cloud -> Settings -> Secrets):
  [salesforce]
  client_id     = "..."
  client_secret = "..."
  redirect_uri  = "https://construction-draw-tracker.streamlit.app/"
  auth_host     = "https://cvest.my.salesforce.com"

Run locally:  streamlit run draw_dashboard.py
"""
from __future__ import annotations

import base64
import hashlib
import io
import json
import secrets
from datetime import date, datetime, timedelta
from typing import Any
from urllib.error import HTTPError
from urllib.parse import parse_qs, urlencode, urlparse
from urllib.request import Request, urlopen

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st
from pandas.tseries.holiday import USFederalHolidayCalendar
from simple_salesforce import Salesforce

# ═══════════════════════════════ fields ═════════════════════════════════════
CONSTRUCTION_DEV_NAME = "Construction_Advance"
PKG  = "Date_Submitted_to_Capital_Partner__c"      # label: "Date Full Draw Package Received"
WIRE = "Wire_Date__c"
REQ  = "Date_Advance_Requested__c"
INSP = "Date_Of_Inspection__c"
RPT  = "Date_Inspection_Report_Received__c"
BUSINESS_TZ = "America/Los_Angeles"                 # datetime fields -> local date

# Salesforce field -> column heading (same headings as the notebook export)
TEXT_FIELDS = {
    "Name": "Advance #",
    "Deal__r.Name": "Property",
    "Lender__c": "Lender",
    "Status__c": "Status",
    "IC_Approval_Status__c": "IC Approval",
    "Advance_Coordinator__r.Name": "Advance Coordinator",
    "Advance_Analyst__r.Name": "Advance Analyst",
    "Underwriter__r.Name": "Underwriter",
    "Exception__c": "Exception",
    "Cancellation_Reason__c": "Cancellation Reason",
    "Inspection_Method__c": "Inspection Method",
}
DATE_FIELDS = {                                     # in milestone order
    REQ: "Advance Requested",
    "Date_Inspection_Ordered__c": "Inspection Ordered",
    INSP: "Inspection Date",
    RPT: "Inspection Report Received",
    "Date_Submitted_For_Approval__c": "Submitted For Review",
    "Date_Internal_Review_Complete__c": "Internal Review Complete",
    PKG: "Full Draw Package Received",
    "Manager_Approval_Date__c": "Manager Approval",
    WIRE: "Wire Date",
}
# Only selected when describe() says the org has them.
OPTIONAL_ADVANCE = {"Net_Funded_Amount__c": "Net Funded ($)",
                    "Target_Advance_Date__c": "Requested Funding Date"}
OPTIONAL_DEAL = {"Deal_Loan_Number__c": "Loan #",           # Opportunity, per the column glossary
                 "LOC_Loan_Type__c": "Product Type",
                 "Payoff_Date__c": "Loan Payoff Date"}
LG_FIELD = "ConstructionManagementLoanId__c"                  # Property__c "Land Gorilla Loan ID"

# Computed columns
T_PKG   = "Turn-Time bd (Package→Wire)"     # OFFICIAL
T_INSP  = "Turn-Time bd (InspRpt→Wire)"     # cross-check, broader coverage
T_PRE   = "Pre-Package bd (Req→Package)"    # borrower / inspection / title
T_TOTAL = "Total Elapsed bd (Req→Wire)"
STAGES = [("Requested → Inspection", "Advance Requested", "Inspection Date"),
          ("Inspection → Report received", "Inspection Date", "Inspection Report Received"),
          ("Report received → Full package", "Inspection Report Received", "Full Draw Package Received")]

OPEN_EXCLUDE = ["Completed", "Cancelled", "Rescinded", "Rejected by Capital Partner"]
HOLD_STATUSES = ["Hold", "Pending Borrower Response", "Pending Inspection Report Revision"]
PERIODS = ["Last month", "Last quarter", "Last year", "Month to date", "Quarter to date",
           "Year to date", "Last 90 days", "All time", "Custom range"]

_HOLS = USFederalHolidayCalendar().holidays("2018-01-01", "2032-12-31").values.astype("datetime64[D]")


# ═══════════════════════════════ Salesforce OAuth ═══════════════════════════
def install_truststore() -> None:
    try:
        import truststore
        truststore.inject_into_ssl()
    except Exception:
        pass


def load_salesforce_oauth_config() -> dict[str, str]:
    section = dict(st.secrets.get("salesforce", {}))
    missing = [k for k in ["client_id", "client_secret", "redirect_uri", "auth_host"] if not section.get(k)]
    if missing:
        raise RuntimeError("Missing Salesforce secrets: " + ", ".join(missing)
                           + ". Add them under [salesforce] in Streamlit secrets.")
    section.setdefault("scope", "api refresh_token")
    section.setdefault("prompt", "login")
    return section


@st.cache_resource
def _pkce_store() -> dict:
    return {}


def build_salesforce_login_url(cfg: dict[str, str]) -> str:
    state = secrets.token_urlsafe(24)
    verifier = base64.urlsafe_b64encode(secrets.token_bytes(64)).rstrip(b"=").decode()
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    store = _pkce_store()
    store[state] = verifier
    for old in list(store)[:-50]:
        store.pop(old, None)
    query = urlencode({
        "response_type": "code", "client_id": cfg["client_id"], "redirect_uri": cfg["redirect_uri"],
        "scope": cfg["scope"], "prompt": cfg["prompt"], "state": state,
        "code_challenge": challenge, "code_challenge_method": "S256",
    })
    return f"{cfg['auth_host'].rstrip('/')}/services/oauth2/authorize?{query}"


def exchange_code_for_token(cfg: dict[str, str], code: str, verifier: str | None) -> dict[str, Any]:
    install_truststore()
    fields = {"grant_type": "authorization_code", "client_id": cfg["client_id"],
              "client_secret": cfg["client_secret"], "redirect_uri": cfg["redirect_uri"], "code": code}
    if verifier:
        fields["code_verifier"] = verifier
    req = Request(f"{cfg['auth_host'].rstrip('/')}/services/oauth2/token",
                  data=urlencode(fields).encode(),
                  headers={"Content-Type": "application/x-www-form-urlencoded"}, method="POST")
    try:
        with urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="ignore")
        try:
            body = json.loads(body).get("error_description", body)
        except Exception:
            pass
        raise RuntimeError(f"Salesforce login failed: {body}") from exc


def finish_login(cfg: dict[str, str], code: str, state: str | None) -> None:
    verifier = _pkce_store().pop(state, None) if state else None
    if state and not verifier:
        raise RuntimeError("That login link has expired (the app restarted between steps). "
                           "Click 'Log in to Salesforce' and try again.")
    payload = exchange_code_for_token(cfg, code, verifier)
    if not payload.get("access_token") or not payload.get("instance_url"):
        raise RuntimeError("Login succeeded but Salesforce returned no access token.")
    st.session_state["salesforce_auth"] = {"access_token": payload["access_token"],
                                           "instance_url": payload["instance_url"]}
    st.session_state["_last_sf_code"] = code


def _qp(name: str) -> str | None:
    v = st.query_params.get(name)
    return v[0] if isinstance(v, list) else v


def maybe_finish_oauth(cfg: dict[str, str]) -> None:
    """Complete the login when Salesforce redirects back here with ?code=...&state=..."""
    if _qp("error"):
        desc = _qp("error_description") or _qp("error")
        st.query_params.clear()
        raise RuntimeError(f"Salesforce login was not completed: {desc}")
    code = _qp("code")
    if not code:
        return
    if st.session_state.get("_last_sf_code") == code and st.session_state.get("salesforce_auth"):
        st.query_params.clear(); return
    state = _qp("state")
    st.query_params.clear()
    finish_login(cfg, code, state)
    st.rerun()


def clear_salesforce_session() -> None:
    for k in ["salesforce_auth", "_last_sf_code"]:
        st.session_state.pop(k, None)


# ═══════════════════════════════ Salesforce data ════════════════════════════
def soql_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("'", "\\'")


def flatten(rec: dict, prefix: str = "") -> dict:
    """{'Deal__r': {'Name': 'x'}} -> {'Deal__r.Name': 'x'}; drops 'attributes'."""
    out: dict[str, Any] = {}
    for k, v in rec.items():
        if k == "attributes":
            continue
        if isinstance(v, dict):
            out.update(flatten(v, f"{prefix}{k}."))
        else:
            out[f"{prefix}{k}"] = v
    return out


def _sf(inst: str, tok: str) -> Salesforce:
    install_truststore()
    return Salesforce(instance_url=inst, session_id=tok)


@st.cache_data(ttl=900, show_spinner=False)
def field_names(inst: str, tok: str, sobject: str) -> set[str]:
    try:
        return {f["name"] for f in getattr(_sf(inst, tok), sobject).describe()["fields"]}
    except Exception as exc:
        if "INVALID_SESSION_ID" in str(exc):
            raise
        return set()


@st.cache_data(ttl=900, show_spinner=False)
def construction_rt_id(inst: str, tok: str) -> str | None:
    recs = _sf(inst, tok).query("SELECT Id, DeveloperName FROM RecordType "
                                "WHERE SobjectType = 'Advance__c'")["records"]
    return next((r["Id"] for r in recs if r["DeveloperName"] == CONSTRUCTION_DEV_NAME), None)


@st.cache_data(ttl=300, show_spinner=False)
def run_soql(inst: str, tok: str, soql: str) -> pd.DataFrame:
    return pd.DataFrame([flatten(r) for r in _sf(inst, tok).query_all(soql)["records"]])


def select_map(inst: str, tok: str) -> dict[str, str]:
    """Salesforce path -> column heading, for every field this org actually has."""
    adv = field_names(inst, tok, "Advance__c")

    def has(path: str) -> bool:
        base = path.split(".")[0]
        base = base[:-3] + "__c" if base.endswith("__r") else base
        return not adv or base in adv          # describe failed -> trust the list

    m = {"Id": "Id", "Deal__c": "Deal Id"}
    m.update({k: v for k, v in {**TEXT_FIELDS, **DATE_FIELDS}.items() if has(k)})
    m.update({k: v for k, v in OPTIONAL_ADVANCE.items() if k in adv})
    deal = field_names(inst, tok, "Opportunity")
    m.update({f"Deal__r.{k}": v for k, v in OPTIONAL_DEAL.items() if k in deal})
    return m


def query_advances(inst: str, tok: str, rt: str, where: str, order: str = WIRE) -> pd.DataFrame:
    fields = select_map(inst, tok)
    raw = run_soql(inst, tok, f"SELECT {', '.join(fields)} FROM Advance__c "
                              f"WHERE RecordTypeId = '{rt}' AND {where} ORDER BY {order} DESC NULLS LAST")
    df = raw.rename(columns=fields).reindex(columns=list(fields.values()))
    return attach_land_gorilla_ids(inst, tok, df)


def attach_land_gorilla_ids(inst: str, tok: str, df: pd.DataFrame) -> pd.DataFrame:
    """Land Gorilla loan ID lives on Property__c; join it per deal (a deal can have several)."""
    if df.empty or LG_FIELD not in field_names(inst, tok, "Property__c"):
        return df
    ids = sorted(df["Deal Id"].dropna().astype(str).unique())
    parts = []
    for i in range(0, len(ids), 200):
        in_list = ",".join(f"'{soql_escape(x)}'" for x in ids[i:i + 200])
        parts.append(run_soql(inst, tok, f"SELECT Deal__c, {LG_FIELD} FROM Property__c "
                                         f"WHERE Deal__c IN ({in_list}) AND {LG_FIELD} != null"))
    props = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
    if props.empty:
        return df
    lg = props.groupby("Deal__c")[LG_FIELD].agg(lambda s: "; ".join(dict.fromkeys(s.astype(str))))
    out = df.copy()
    out.insert(out.columns.get_loc("Property") + 1, "Land Gorilla Loan ID", out["Deal Id"].map(lg))
    return out


# ═══════════════════════════════ turn-time math ═════════════════════════════
def to_date(s: pd.Series) -> pd.Series:
    """Salesforce Date ('2025-09-03') or DateTime ('...T17:02:00.000+0000') -> naive date."""
    s = pd.Series(s, dtype="object").astype("string")
    is_dt = (s.str.len() > 10).fillna(False).astype(bool)
    out = pd.to_datetime(s.str[:10].where(~is_dt), errors="coerce", format="%Y-%m-%d")
    if is_dt.any():
        dt = (pd.to_datetime(s.where(is_dt), errors="coerce", utc=True)
              .dt.tz_convert(BUSINESS_TZ).dt.tz_localize(None).dt.normalize())
        out = out.where(~is_dt, dt)
    return out


def bdays(a: pd.Series, b: pd.Series, same_day_as: int = 0) -> pd.Series:
    """Business days a -> b (weekends + US federal holidays excluded). NaN if either is blank."""
    m = a.notna() & b.notna()
    out = pd.Series(np.nan, index=a.index)
    if m.any():
        out[m] = np.busday_count(a[m].values.astype("datetime64[D]"),
                                 b[m].values.astype("datetime64[D]"), holidays=_HOLS) + same_day_as
    return out


def add_turn_times(df: pd.DataFrame, same_day_as: int) -> pd.DataFrame:
    """Dates -> real dates, then the four intervals. Reversed dates are blanked and flagged."""
    df = df.copy()
    for col in [*DATE_FIELDS.values(), "Requested Funding Date", "Loan Payoff Date"]:
        if col in df:
            df[col] = to_date(df[col])
    if "Net Funded ($)" in df:
        df["Net Funded ($)"] = pd.to_numeric(df["Net Funded ($)"], errors="coerce")
    df["Reversed Dates"] = False
    for name, a, b in [(T_PKG, "Full Draw Package Received", "Wire Date"),
                       (T_INSP, "Inspection Report Received", "Wire Date"),
                       (T_PRE, "Advance Requested", "Full Draw Package Received"),
                       (T_TOTAL, "Advance Requested", "Wire Date")]:
        v = bdays(df[a], df[b], same_day_as)
        df["Reversed Dates"] |= (v < 0).fillna(False)
        df[name] = v.where(v >= 0)
    for name, a, b in STAGES:                 # plain counts so the stages add up
        v = bdays(df[a], df[b])
        df[name] = v.where(v >= 0)
    w = df["Wire Date"]
    df["Wire Year"] = w.dt.year.astype("Int64").astype("string")
    df["Wire Quarter"] = w.dt.to_period("Q").astype("string")
    df["Wire Month"] = w.dt.to_period("M").astype("string")
    return df


def period_bounds(choice: str, today: date | None = None) -> tuple[date, date]:
    today = today or date.today()
    month_start = today.replace(day=1)
    q_start = date(today.year, 3 * ((today.month - 1) // 3) + 1, 1)
    if choice == "Month to date":
        return month_start, today
    if choice == "Quarter to date":
        return q_start, today
    if choice == "Year to date":
        return date(today.year, 1, 1), today
    if choice == "Last 90 days":
        return today - timedelta(days=90), today
    if choice == "Last month":
        end = month_start - timedelta(days=1)
        return end.replace(day=1), end
    if choice == "Last quarter":
        end = q_start - timedelta(days=1)
        return date(end.year, 3 * ((end.month - 1) // 3) + 1, 1), end
    if choice == "Last year":
        return date(today.year - 1, 1, 1), date(today.year - 1, 12, 31)
    if choice == "All time":
        return date(2018, 1, 1), today
    raise ValueError(f"Unknown period: {choice}")


# ═══════════════════════════════ stats ══════════════════════════════════════
def _med(s: pd.Series) -> float | None:
    s = s.dropna()
    return float(s.median()) if len(s) else None


def _pct(s: pd.Series, limit: int) -> float | None:
    s = s.dropna()
    return float((s <= limit).mean()) if len(s) else None


def _q90(s: pd.Series) -> float | None:
    s = s.dropna()
    return float(s.quantile(0.9)) if len(s) else None


def fmt_bd(x: float | None) -> str:
    return "—" if x is None or pd.isna(x) else f"{x:.0f}"


def fmt_pct(x: float | None) -> str:
    return "—" if x is None else f"{x:.0%}"


def money(x) -> str:
    return "—" if x is None or pd.isna(x) else f"${x:,.0f}"


def rollup(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """Per-group stats. Official turn-time is over draws that have a package date."""
    key = df[by].fillna("(none)")
    t = df[T_PKG].groupby(key)
    out = pd.DataFrame({"Draws Wired": df.groupby(key).size(), "Measured": t.count()})
    if "Net Funded ($)" in df:
        out["Funded ($)"] = df["Net Funded ($)"].groupby(key).sum()
    out["Median bd"] = t.median()
    out["Mean bd"] = t.mean().round(1)
    out["% ≤3 bd"] = t.apply(lambda s: _pct(s, 3))
    out["Median Pre-Package bd"] = df[T_PRE].groupby(key).median()
    out.index.name = by
    return out.reset_index()


def headline_rows(df: pd.DataFrame) -> list[tuple[str, str]]:
    t, pre, ins = df[T_PKG], df[T_PRE], df[T_INSP]
    n, no = len(df), int(t.notna().sum())
    rows = [("Construction advances (wired)", f"{n:,}")]
    if "Net Funded ($)" in df:
        rows.append(("Total funded", money(df["Net Funded ($)"].sum())))
    return rows + [
        ("…with a Full Draw Package Received date", f"{no:,}  ({no / n:.0%} coverage)" if n else "0"),
        ("Median turn-time (business days)", fmt_bd(_med(t))),
        ("Mean turn-time (business days)", f"{t.mean():.1f}" if no else "—"),
        ("% funded within 1 business day", fmt_pct(_pct(t, 1))),
        ("% funded within 3 business days", fmt_pct(_pct(t, 3))),
        ("% funded within 5 business days", fmt_pct(_pct(t, 5))),
        ("Median PRE-package (borrower/inspection/title)", f"{fmt_bd(_med(pre))} bd  (90th pct {fmt_bd(_q90(pre))})"),
        ("Median total elapsed (request → wire)", f"{fmt_bd(_med(df[T_TOTAL]))} bd"),
        ("Secondary view — Inspection Report→Wire",
         f"median {fmt_bd(_med(ins))} bd, {fmt_pct(_pct(ins, 3))} ≤3  (n={int(ins.notna().sum()):,})"),
    ]


def readme_rows(*, period: str, start: date, end: date, same_day: int, coverage: float,
                reversed_rows: int, filters: list[str]) -> list[tuple[str, str]]:
    return [
        ("Generated", datetime.now().strftime("%m/%d/%Y %H:%M")),
        ("Report period", f"{period}  ·  wire date {start:%m/%d/%Y} – {end:%m/%d/%Y}"),
        ("Filters", "; ".join(filters) if filters else "None"),
        ("Source", "Salesforce Advance__c · Record Type: Construction Advance · wired advances only. "
                   "No filter on loan status, so draws on loans paid off after funding are included."),
        ("Official metric", "Business days from FULL DRAW PACKAGE RECEIVED to WIRE DATE"),
        ("  field used", f"{PKG}  (label: 'Date Full Draw Package Received')"),
        ("Business days", "Excludes weekends and US federal bank holidays"),
        ("Same-day convention", f"package-in / wire same-day = {same_day} business day(s)"),
        ("Coverage caveat", f"Only {coverage:.0%} of these draws carry a package-received date; the metric "
                            "covers that subset. Fuller coverage needs the Land Gorilla feed / better SF data "
                            "entry (ref ticket IHD-109768)."),
        ("Delay story", "For completed draws Status shows 'Completed' only (hold history not retained). Use "
                        "the Pre-Package interval (Req→Package) as the borrower/inspection/title delay measure."),
        ("Secondary view", "Inspection Report Received→Wire is included for broader coverage as a cross-check."),
        ("Reversed-date rows", f"{reversed_rows} rows had an end date before the start date; those intervals "
                               "were blanked and excluded from metrics."),
        ("How to filter by period", "Use 'Draw Detail' → filter on Wire Year / Wire Quarter / Wire Month, "
                                    "or build a PivotTable off that sheet."),
    ]


# ═══════════════════════════════ Excel export ═══════════════════════════════
# Same look as the notebook export: Arial 10, navy header row, thin grey borders.
def _xl_styles():
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    thin = Side(style="thin", color="D9D9D9")
    return {
        "body": Font(name="Arial", size=10), "bold": Font(name="Arial", size=10, bold=True),
        "head": Font(name="Arial", size=10, bold=True, color="FFFFFF"),
        "fill": PatternFill("solid", fgColor="1F3864"),
        "title": Font(name="Arial", size=14, bold=True, color="1F3864"),
        "sub": Font(name="Arial", size=9, italic=True, color="808080"),
        "border": Border(left=thin, right=thin, top=thin, bottom=thin),
        "center": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "wrap": Alignment(wrap_text=True, vertical="top"),
    }


def _cell_value(v):
    if v is None or (np.ndim(v) == 0 and pd.isna(v)):
        return None
    if isinstance(v, np.generic):
        return v.item()
    return v


def write_df(ws, df: pd.DataFrame, r0: int, c0: int = 1, money_cols=(), pct_cols=()) -> int:
    """Write a header + rows at (r0, c0); returns the next free row."""
    s = _xl_styles()
    for j, col in enumerate(df.columns):
        cell = ws.cell(r0, c0 + j, col)
        cell.font, cell.fill, cell.border, cell.alignment = s["head"], s["fill"], s["border"], s["center"]
    for i, row in enumerate(df.itertuples(index=False), start=r0 + 1):
        for j, (col, v) in enumerate(zip(df.columns, row)):
            cell = ws.cell(i, c0 + j, _cell_value(v))
            cell.font, cell.border = s["body"], s["border"]
            if isinstance(cell.value, (datetime, date)):
                cell.number_format = "mm/dd/yyyy"
            elif col in money_cols:
                cell.number_format = "$#,##0"
            elif col in pct_cols:
                cell.number_format = "0%"
    return r0 + len(df) + 1


def _size_columns(ws, df: pd.DataFrame, c0: int = 1, wide: tuple[str, ...] = ()):
    from openpyxl.utils import get_column_letter
    for j, col in enumerate(df.columns):
        vals = [len(str(v)) for v in df[col].head(500) if pd.notna(v)]
        is_date = pd.api.types.is_datetime64_any_dtype(df[col])
        width = 14 if is_date else min(max([len(str(col)) * 0.9, 10] + vals) + 2, 60)
        ws.column_dimensions[get_column_letter(c0 + j)].width = 34 if col in wide else width


def _table_sheet(wb, name: str, df: pd.DataFrame, money_cols=(), pct_cols=(), wrap_cols=()):
    from openpyxl.utils import get_column_letter
    ws = wb.create_sheet(name[:31])
    write_df(ws, df, 1, money_cols=money_cols, pct_cols=pct_cols)
    ws.freeze_panes = "A2"
    if len(df):
        ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
    _size_columns(ws, df)
    s = _xl_styles()
    for j, col in enumerate(df.columns, start=1):
        if col in wrap_cols:
            ws.column_dimensions[get_column_letter(j)].width = 50
            for r in range(2, len(df) + 2):
                ws.cell(r, j).alignment = s["wrap"]
    return ws


def build_workbook(sheets: dict[str, pd.DataFrame], money_cols: tuple[str, ...] = ()) -> bytes:
    """Plain multi-sheet export (used for the open pipeline)."""
    import openpyxl
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for name, df in sheets.items():
        _table_sheet(wb, name, df, money_cols=money_cols)
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def build_turn_time_workbook(*, readme: list[tuple[str, str]], headline: list[tuple[str, str]],
                             subtitle: str, rollups: dict[str, pd.DataFrame], detail: pd.DataFrame,
                             money_cols=(), wrap_cols=()) -> bytes:
    """Read Me · Turn-Time Summary (headline, By Year, By Quarter, chart) · Draw Detail · extra rollups."""
    import openpyxl
    from openpyxl.chart import BarChart, Reference
    s = _xl_styles()
    wb = openpyxl.Workbook(); wb.remove(wb.active)

    rm = wb.create_sheet("Read Me")
    rm["A1"] = "Construction Draw Turn-Time Report"; rm["A1"].font = s["title"]
    for r, (k, v) in enumerate(readme, start=3):
        rm.cell(r, 1, k).font = s["bold"]
        c = rm.cell(r, 2, v); c.font = s["body"]; c.alignment = s["wrap"]
    rm.column_dimensions["A"].width = 26; rm.column_dimensions["B"].width = 95

    sm = wb.create_sheet("Turn-Time Summary")
    sm["A1"] = "Turn-Time Summary — Full Draw Package Received → Wire"; sm["A1"].font = s["title"]
    sm["A2"] = subtitle; sm["A2"].font = s["sub"]
    r = 4
    for k, v in headline:
        a, b = sm.cell(r, 1, k), sm.cell(r, 2, v)
        a.font, b.font, a.border, b.border = s["bold"], s["body"], s["border"], s["border"]
        r += 1
    pct = ("% ≤3 bd",)
    for title in ["By Year", "By Quarter"]:
        tbl = rollups.get(title)
        if tbl is None or tbl.empty:
            continue
        r += 1; sm.cell(r, 1, title).font = s["title"]; r += 1
        top = r
        r = write_df(sm, tbl, r, money_cols=money_cols, pct_cols=pct)
        if title == "By Quarter" and "% ≤3 bd" in tbl:
            col = list(tbl.columns).index("% ≤3 bd") + 1
            ch = BarChart(); ch.title = "% funded within 3 business days, by quarter"
            ch.y_axis.title = "% ≤3 bd"; ch.y_axis.numFmt = "0%"; ch.height, ch.width = 7, 18
            ch.add_data(Reference(sm, min_col=col, min_row=top, max_row=top + len(tbl)), titles_from_data=True)
            ch.set_categories(Reference(sm, min_col=1, min_row=top + 1, max_row=top + len(tbl)))
            ch.legend = None
            sm.add_chart(ch, "H4")
    sm.column_dimensions["A"].width = 46
    for col in "BCDEF":
        sm.column_dimensions[col].width = 16

    _table_sheet(wb, "Draw Detail", detail, money_cols=money_cols, wrap_cols=wrap_cols)
    for name, tbl in rollups.items():
        if name not in ("By Year", "By Quarter") and tbl is not None and not tbl.empty:
            _table_sheet(wb, name, tbl, money_cols=money_cols, pct_cols=pct)

    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


# ═══════════════════════════════ charts ═════════════════════════════════════
# Deck-style: navy bars, orange highlight, grey = muted / caveat. Every bar carries its
# value label (orange and grey are low-contrast on white). White background so the
# chart menu's "Save as PNG" drops straight into a deck.
NAVY, ORANGE, GREY, INK, MUTED = "#1F3864", "#EF7D22", "#B8B8B8", "#333333", "#777777"


def _style(chart: alt.Chart, title: str, subtitle: str | None = None, height: int = 300) -> alt.Chart:
    return (chart.properties(
                title=alt.TitleParams(title, subtitle=subtitle or "", anchor="start", fontSize=16,
                                      color=NAVY, subtitleColor=ORANGE, subtitleFontSize=12,
                                      subtitleFontWeight="bold", offset=12),
                height=height, background="white",
                padding={"left": 12, "right": 12, "top": 12, "bottom": 12})
            .configure_view(stroke=None)
            .configure_axis(labelColor=INK, titleColor=MUTED, gridColor="#EEEEEE",
                            domainColor="#CCCCCC", tickColor="#CCCCCC", labelFontSize=11, titleFontSize=11))


def chart_cumulative(t: pd.Series) -> alt.Chart:
    """HEADLINE: cumulative % of complete packages wired within N business days."""
    t = t.dropna()
    data = pd.DataFrame({"Days": [str(d) for d in range(8)],
                         "Pct": [(t <= d).mean() for d in range(8)]})
    data["Highlight"] = data["Days"] == "3"
    base = alt.Chart(data).encode(
        x=alt.X("Days:N", sort=None, title="Business days from complete draw package to wire",
                axis=alt.Axis(labelAngle=0)),
        y=alt.Y("Pct:Q", title="Draws funded (cumulative)", scale=alt.Scale(domain=[0, 1.08]),
                axis=alt.Axis(format="%", tickCount=6)),
        tooltip=[alt.Tooltip("Days:N", title="Within N business days"),
                 alt.Tooltip("Pct:Q", title="Draws funded", format=".0%")])
    bars = base.mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, size=46).encode(
        color=alt.condition("datum.Highlight", alt.value(ORANGE), alt.value(NAVY)))
    labels = base.mark_text(dy=-8, fontSize=12).encode(
        text=alt.Text("Pct:Q", format=".0%"),
        color=alt.condition("datum.Highlight", alt.value(ORANGE), alt.value(INK)))
    med = t.median()
    return _style(bars + labels, "Once the package is complete, we fund fast",
                  f"~{(t <= 3).mean():.0%} of complete packages wired within 3 business days · "
                  f"median {med:.0f} day(s) · n={len(t):,}", height=340)


def chart_distribution(t: pd.Series) -> alt.Chart:
    order = ["0", "1", "2", "3", "4–5", "6–10", "10+"]
    b = pd.cut(t.dropna(), [-np.inf, 0, 1, 2, 3, 5, 10, np.inf], labels=order)
    data = b.value_counts().reindex(order, fill_value=0).rename_axis("Bucket").reset_index(name="Draws")
    base = alt.Chart(data).encode(
        x=alt.X("Bucket:N", sort=order, title="Business days", axis=alt.Axis(labelAngle=0)),
        y=alt.Y("Draws:Q", title="Number of draws"),
        tooltip=[alt.Tooltip("Bucket:N", title="Business days"), alt.Tooltip("Draws:Q", format=",")])
    return _style(base.mark_bar(color=NAVY, cornerRadiusTopLeft=4, cornerRadiusTopRight=4, size=36)
                  + base.mark_text(dy=-8, fontSize=11, color=INK).encode(text=alt.Text("Draws:Q", format=",")),
                  "How long funding takes (package → wire)")


def chart_where_time_goes(df: pd.DataFrame) -> alt.Chart:
    pre, fund = df[T_PRE].dropna(), df[T_PKG].dropna()
    rows = []
    for stage, s in [("Borrower / inspection / title", pre), ("Our funding", fund)]:
        if len(s):
            rows += [{"Stage": stage, "Measure": "Median", "Days": float(s.median())},
                     {"Stage": stage, "Measure": "90th percentile", "Days": float(s.quantile(.9))}]
    data = pd.DataFrame(rows)
    color = alt.Color("Measure:N", scale=alt.Scale(domain=["Median", "90th percentile"], range=[GREY, ORANGE]),
                      legend=alt.Legend(title=None, orient="top-right"))
    base = alt.Chart(data).encode(
        y=alt.Y("Stage:N", title=None, sort=None, axis=alt.Axis(labelLimit=260, labelFontSize=12),
                scale=alt.Scale(paddingInner=0.25)),
        yOffset=alt.YOffset("Measure:N", sort=["Median", "90th percentile"]),
        x=alt.X("Days:Q", title="Business days", axis=alt.Axis(format="d", tickMinStep=1)),
        tooltip=["Stage:N", "Measure:N", alt.Tooltip("Days:Q", format=".0f")])
    return _style(base.mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4, size=26).encode(color=color)
                  + base.mark_text(dx=6, align="left", fontSize=11, color=INK).encode(text=alt.Text("Days:Q", format=".0f")),
                  "The delay is upstream — not funding",
                  "Request → complete package vs. complete package → wire")


def chart_by_year(df: pd.DataFrame, caveat_years: list[str]) -> alt.Chart:
    off = df[df[T_PKG].notna()]
    data = (off.groupby("Wire Year")[T_PKG].agg(["median", "size"]).reset_index()
            .rename(columns={"median": "Median", "size": "Draws"}))
    data["Package dates"] = np.where(data["Wire Year"].isin(caveat_years), "Unreliable (entered = wire date)", "Reliable")
    color = alt.Color("Package dates:N", scale=alt.Scale(domain=["Reliable", "Unreliable (entered = wire date)"],
                                                         range=[NAVY, GREY]),
                      legend=alt.Legend(title=None, orient="top-right"))
    base = alt.Chart(data).encode(
        x=alt.X("Wire Year:N", title="Wire year", axis=alt.Axis(labelAngle=0)),
        y=alt.Y("Median:Q", title="Median business days", axis=alt.Axis(format="d", tickMinStep=1)),
        tooltip=["Wire Year:N", alt.Tooltip("Median:Q", format=".0f"), alt.Tooltip("Draws:Q", format=","),
                 "Package dates:N"])
    return _style(base.mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, size=40).encode(color=color)
                  + base.mark_text(dy=-8, fontSize=11, color=INK).encode(text=alt.Text("Median:Q", format=".0f")),
                  "Median funding time by year")


def chart_coverage(total: int, covered: int) -> alt.Chart:
    data = pd.DataFrame([{"Group": "Measurable (has package date)", "Draws": covered, "Order": 0},
                         {"Group": "No package date", "Draws": total - covered, "Order": 1}])
    bar = alt.Chart(data).mark_bar(size=34).encode(
        x=alt.X("Draws:Q", stack=True, title=None, axis=alt.Axis(format=",.0f", tickCount=5)),
        color=alt.Color("Group:N", scale=alt.Scale(domain=list(data["Group"]), range=[ORANGE, "#E5E5E5"]),
                        legend=alt.Legend(title=None, orient="bottom")),
        order="Order:Q",
        tooltip=["Group:N", alt.Tooltip("Draws:Q", format=",")])
    return _style(bar, "Coverage today — why Land Gorilla matters",
                  f"{covered:,} of {total:,} wired draws measurable ({covered / total:.0%})" if total else "",
                  height=90)


# ═══════════════════════════════ tables ═════════════════════════════════════
DATE_COLS = set(DATE_FIELDS.values()) | {"Requested Funding Date", "Loan Payoff Date"}


def col_config(df: pd.DataFrame) -> dict:
    cfg: dict[str, Any] = {}
    for c in df.columns:
        if c in DATE_COLS:
            cfg[c] = st.column_config.DateColumn(c, format="MM/DD/YYYY")
        elif c.startswith("%"):
            cfg[c] = st.column_config.NumberColumn(c, format="percent")
        elif "($)" in c:
            cfg[c] = st.column_config.NumberColumn(c, format="dollar")
        elif c.startswith("Mean"):
            cfg[c] = st.column_config.NumberColumn(c, format="%.1f")
        elif " bd" in c or c in ("Days Open", "Draws Wired", "Measured"):
            cfg[c] = st.column_config.NumberColumn(c, format="%d")
        elif c == "Salesforce":
            cfg[c] = st.column_config.LinkColumn(c, display_text="Open ↗")
    return cfg


def show_table(df: pd.DataFrame, height: int | None = None) -> None:
    kw = {"height": height} if height else {}
    st.dataframe(df, hide_index=True, width="stretch", column_config=col_config(df), **kw)


DETAIL_COLS = ["Advance #", "Property", "Land Gorilla Loan ID", "Loan #", "Product Type", "Lender", "Status",
               "Advance Coordinator", "Advance Analyst", "Inspection Method", "Net Funded ($)",
               *DATE_FIELDS.values(), T_PKG, T_INSP, T_PRE, T_TOTAL, "Loan Payoff Date",
               "Wire Year", "Wire Quarter", "Wire Month", "Salesforce"]


def detail_view(df: pd.DataFrame, inst: str) -> pd.DataFrame:
    out = df.copy()
    out["Salesforce"] = inst.rstrip("/") + "/" + out["Id"].astype(str)
    return out[[c for c in DETAIL_COLS if c in out]]


# ═══════════════════════════════ pages ══════════════════════════════════════
def _filter_row(df: pd.DataFrame, cols, specs: list[tuple[str, str]], key: str) -> tuple[pd.DataFrame, list[str]]:
    applied = []
    for slot, (col, label) in zip(cols, specs):
        if col in df and df[col].notna().any():
            opts = sorted(df[col].dropna().astype(str).unique())
            pick = slot.multiselect(label, opts, placeholder="All", key=f"{key}_{col}")
            if pick:
                df = df[df[col].astype(str).isin(pick)]
                applied.append(f"{label}: {', '.join(pick)}")
    return df, applied


def page_turn_time(inst: str, tok: str, rt: str, same_day: int, caveat_years: list[str]) -> None:
    c = st.columns([1.2, 1, 1, 1, 1])
    period = c[0].selectbox("Wire date", PERIODS, index=PERIODS.index("All time"), key="tt_period")
    if period == "Custom range":
        d = st.columns([1.2, 1.2, 3])
        start = d[0].date_input("From", date.today().replace(day=1), format="MM/DD/YYYY", key="tt_from")
        end = d[1].date_input("To", date.today(), format="MM/DD/YYYY", key="tt_to")
        label = f"{start:%m/%d/%Y} – {end:%m/%d/%Y}"
    else:
        start, end = period_bounds(period)
        label = period
    if start > end:
        st.error("'From' is after 'To'."); return

    with st.spinner("Pulling construction draws from Salesforce…"):
        df = query_advances(inst, tok, rt, f"{WIRE} >= {start:%Y-%m-%d} AND {WIRE} <= {end:%Y-%m-%d}")
    if df.empty:
        st.info(f"No construction draws were wired {start:%m/%d/%Y} – {end:%m/%d/%Y}."); return
    df = add_turn_times(df, same_day)
    df, applied = _filter_row(df, c[1:], [("Lender", "Lender"), ("Product Type", "Product type"),
                                          ("Inspection Method", "Inspection method"),
                                          ("Advance Coordinator", "Coordinator")], "tt")
    if df.empty:
        st.info("No draws match these filters."); return

    t = df[T_PKG]
    n, no = len(df), int(t.notna().sum())
    reversed_rows = int(df["Reversed Dates"].sum())
    st.caption(f"Wire dates {start:%m/%d/%Y} – {end:%m/%d/%Y} · {n:,} construction draws"
               + (f" · filters: {'; '.join(applied)}" if applied else ""))

    # ── KPI row ──
    k = st.columns(5)
    k[0].metric("Funded within 3 business days", fmt_pct(_pct(t, 3)), border=True,
                help="Share of draws with a complete package that were wired ≤3 business days later.")
    k[1].metric("Median funding time", f"{fmt_bd(_med(t))} bd", border=True,
                help="Full Draw Package Received → Wire Date, business days.")
    k[2].metric("Median borrower-side time", f"{fmt_bd(_med(df[T_PRE]))} bd", border=True,
                help="Advance Requested → Full Draw Package Received: borrower, inspection, title.")
    k[3].metric("Draws wired", f"{n:,}",
                money(df["Net Funded ($)"].sum()) + " funded" if "Net Funded ($)" in df else None,
                delta_color="off", border=True)
    k[4].metric("Measurable", f"{no / n:.0%}", f"{no:,} have a package date", delta_color="off", border=True)

    if no == 0:
        st.warning("None of these draws has a Full Draw Package Received date, so funding time can't be "
                   "measured for this period."); return
    if no / n < 0.5:
        st.caption(f"⚠ Only {no / n:.0%} of these draws have a Full Draw Package Received date in Salesforce — "
                   "turn-time is measured on that subset. Land Gorilla (IHD-109768) extends it to the full book.")
    if reversed_rows:
        st.caption(f"⚠ {reversed_rows} draw(s) have an end date before the start date; those intervals are "
                   "excluded from the metrics.")

    # ── charts ──
    st.altair_chart(chart_cumulative(t), width="stretch")
    a, b = st.columns(2)
    a.altair_chart(chart_distribution(t), width="stretch")
    b.altair_chart(chart_where_time_goes(df), width="stretch")
    a, b = st.columns(2)
    a.altair_chart(chart_by_year(df, caveat_years), width="stretch")
    if any(y in caveat_years for y in df["Wire Year"].dropna().unique()):
        a.caption("Grey years: package date is mostly entered as the wire date in Salesforce, so funding "
                  "time reads artificially fast. Quote the navy years. (Set in the sidebar.)")
    b.altair_chart(chart_coverage(n, no), width="stretch")
    b.caption("Only draws with a recorded package-complete date can be measured. Land Gorilla captures it "
              "on every RB0 draw.")

    # ── tables ──
    rollups = {"By Year": rollup(df, "Wire Year"), "By Quarter": rollup(df, "Wire Quarter"),
               "By Month": rollup(df, "Wire Month"), "By Coordinator": rollup(df, "Advance Coordinator")}
    same = (df["Full Draw Package Received"] == df["Wire Date"]).groupby(df["Wire Year"]).mean()
    rollups["By Year"]["% Package = Wire Date"] = rollups["By Year"]["Wire Year"].map(same)
    detail = detail_view(df, inst)
    slow = detail.dropna(subset=[T_PKG]).sort_values(T_PKG, ascending=False).head(15)

    st.subheader("Numbers")
    names = [*rollups, "Slowest 15", "All draws"]
    for tab, name in zip(st.tabs(names), names):
        with tab:
            if name == "Slowest 15":
                show_table(slow)
            elif name == "All draws":
                show_table(detail, height=460)
            else:
                show_table(rollups[name])

    xlsx = build_turn_time_workbook(
        readme=readme_rows(period=label, start=start, end=end, same_day=same_day, coverage=no / n,
                           reversed_rows=reversed_rows, filters=applied),
        headline=headline_rows(df),
        subtitle=f"{label} · median {fmt_bd(_med(t))} business day(s) · {fmt_pct(_pct(t, 3))} funded "
                 f"within 3 · n={no:,}",
        rollups=rollups, detail=detail.drop(columns=["Salesforce"]),
        money_cols=("Funded ($)", "Net Funded ($)"))
    st.download_button("⬇  Download Excel report", xlsx, type="primary",
                       file_name=f"Construction_Draw_TurnTime_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.xlsx")
    st.caption("Charts: use the ⋯ menu on any chart → Save as PNG for decks.")


def page_pipeline(inst: str, tok: str, rt: str, same_day: int) -> None:
    excluded = ", ".join(f"'{s}'" for s in OPEN_EXCLUDE)
    with st.spinner("Pulling open draws…"):
        df = query_advances(inst, tok, rt, f"{WIRE} = null AND Status__c NOT IN ({excluded})", order=REQ)
    if df.empty:
        st.success("No open construction draws."); return
    df = add_turn_times(df, same_day)
    df["Days Open"] = (pd.Timestamp(date.today()) - df["Advance Requested"]).dt.days

    c = st.columns(4)
    df, _ = _filter_row(df, c, [("Status", "Status"), ("Lender", "Lender"),
                                ("Advance Coordinator", "Coordinator"), ("Product Type", "Product type")], "op")
    held = df["Status"].isin(HOLD_STATUSES)
    k = st.columns(4)
    k[0].metric("Open draws", f"{len(df):,}", border=True)
    k[1].metric("On hold / waiting on borrower", f"{int(held.sum()):,}", border=True)
    k[2].metric("Pending inspection", f"{int(df['Status'].str.contains('Inspection', na=False).sum()):,}", border=True)
    k[3].metric("Oldest open", f"{fmt_bd(df['Days Open'].max())} days", border=True)

    by = df["Status"].fillna("(blank)").value_counts().rename_axis("Status").reset_index(name="Draws")
    base = alt.Chart(by).encode(y=alt.Y("Status:N", sort="-x", title=None, axis=alt.Axis(labelLimit=240)),
                                x=alt.X("Draws:Q", title="Open draws"),
                                tooltip=["Status:N", alt.Tooltip("Draws:Q", format=",")])
    st.altair_chart(_style(base.mark_bar(color=NAVY, cornerRadiusTopRight=4, cornerRadiusBottomRight=4, size=20)
                           + base.mark_text(dx=6, align="left", fontSize=11, color=INK).encode(text="Draws:Q"),
                           "Open draws by status", height=max(160, 34 * len(by))), width="stretch")

    view = detail_view(df, inst)
    view.insert(view.columns.get_loc("Status") + 1, "Days Open", df["Days Open"])
    keep = [c for c in view.columns if c not in (T_PKG, T_INSP, T_TOTAL, "Wire Date", "Wire Year",
                                                  "Wire Quarter", "Wire Month")]
    view = view[keep].sort_values("Days Open", ascending=False)
    st.subheader("Oldest first")
    show_table(view, height=460)
    st.download_button("⬇  Download open pipeline (Excel)",
                       build_workbook({"Open draws": view.drop(columns=["Salesforce"])},
                                      money_cols=("Net Funded ($)",)),
                       file_name=f"Open_Draw_Pipeline_{date.today():%Y-%m-%d}.xlsx", type="primary")


def page_lookup(inst: str, tok: str, rt: str, same_day: int) -> None:
    fields = select_map(inst, tok)
    modes = {"Property": "Deal__r.Name", "Advance #": "Name"}
    if "Deal__r.Deal_Loan_Number__c" in fields:
        modes["Loan #"] = "Deal__r.Deal_Loan_Number__c"
    c = st.columns([3, 1.4])
    text = c[0].text_input("Search", placeholder="e.g. 745 South 9th Street", label_visibility="collapsed")
    mode = c[1].segmented_control("Match on", list(modes), default="Property", label_visibility="collapsed")
    if not text:
        st.info("Search a property, advance # or loan # to see every milestone date for its draws."); return
    with st.spinner("Searching…"):
        df = query_advances(inst, tok, rt, f"{modes[mode or 'Property']} LIKE '%{soql_escape(text)}%'", order=REQ)
    if df.empty:
        st.warning("No matching construction advances."); return
    df = add_turn_times(df, same_day)

    st.caption(f"{len(df)} matching draw(s), newest first.")
    view = detail_view(df, inst)
    show_table(view[[c for c in ["Advance #", "Property", "Loan #", "Status", "Advance Requested",
                                 "Full Draw Package Received", "Wire Date", T_PKG, T_PRE, "Salesforce"]
                     if c in view]])

    pick = st.selectbox("Show milestones for", df["Advance #"] + " · " + df["Property"].fillna(""))
    row = df.iloc[list(df["Advance #"] + " · " + df["Property"].fillna("")).index(pick)]
    a, b = st.columns([1.3, 1])
    with a:
        steps = pd.DataFrame({"Milestone": list(DATE_FIELDS.values()),
                              "Date": [row[c] for c in DATE_FIELDS.values()]})
        steps.insert(0, "", np.where(steps["Date"].notna(), "✅", "⬜"))
        show_table(steps)
    with b:
        b.metric("Funding time (package → wire)", f"{fmt_bd(row[T_PKG])} bd", border=True)
        b.metric("Borrower-side (request → package)", f"{fmt_bd(row[T_PRE])} bd", border=True)
        b.metric("Total (request → wire)", f"{fmt_bd(row[T_TOTAL])} bd", border=True)
        facts = [f"**{k}:** {row[k]}" for k in ["Status", "Lender", "Land Gorilla Loan ID", "Loan #",
                                                 "Advance Coordinator", "Inspection Method"]
                 if k in row and pd.notna(row[k])]
        st.markdown("  \n".join(facts))


# ═══════════════════════════════ main ═══════════════════════════════════════
def login_screen(cfg: dict[str, str]) -> None:
    _, mid, _ = st.columns([1, 2, 1])
    with mid, st.container(border=True):
        st.markdown("### Log in to Salesforce")
        st.write("Construction draw turn-times, open pipeline and draw lookup — live from Salesforce.")
        st.link_button("Log in to Salesforce", build_salesforce_login_url(cfg), type="primary")
        with st.expander("Landed on a different app after logging in?"):
            st.caption("If Salesforce sent you to another app (its callback URL is set there), copy that "
                       "page's full address and paste it here.")
            pasted = st.text_input("Redirected URL", placeholder="https://…/?code=…&state=…",
                                   label_visibility="collapsed")
            if pasted:
                q = parse_qs(urlparse(pasted.strip()).query)
                if "code" not in q:
                    st.error("That address has no ?code= in it.")
                else:
                    try:
                        finish_login(cfg, q["code"][0], q.get("state", [None])[0])
                        st.rerun()
                    except RuntimeError as exc:
                        st.error(str(exc))
        st.caption(f"Callback URL: {cfg['redirect_uri']}")


def main() -> None:
    st.set_page_config(page_title="Construction Draw Tracker", page_icon="🏗️", layout="wide")
    st.markdown("""<style>
        .block-container {padding-top: 2rem; max-width: 1400px;}
        [data-testid="stMetricValue"] {font-size: 1.9rem;}
        h1 {color: #1F3864;}
    </style>""", unsafe_allow_html=True)
    st.title("Construction Draw Tracker")

    try:
        cfg = load_salesforce_oauth_config()
        maybe_finish_oauth(cfg)
    except RuntimeError as exc:
        st.error(str(exc)); st.stop()

    auth = st.session_state.get("salesforce_auth")
    if not auth:
        login_screen(cfg); st.stop()
    inst, tok = auth["instance_url"], auth["access_token"]

    with st.sidebar:
        st.success("Connected to Salesforce")
        st.caption(inst)
        if st.button("Log out", width="stretch"):
            clear_salesforce_session(); st.rerun()
        st.divider()
        st.subheader("Settings")
        same_day = 1 if st.radio("Package in and wired same day counts as",
                                 ["0 business days", "1 business day"]).startswith("1") else 0
        this_year = date.today().year
        caveat_years = st.multiselect(
            "Years with unreliable package dates", [str(y) for y in range(2019, this_year + 1)],
            default=[str(y) for y in (2025, 2026) if y <= this_year],
            help="Years where the package date was mostly entered as the wire date. Shown grey; check "
                 "'% Package = Wire Date' in the By Year table.")
        if st.button("Refresh data", width="stretch"):
            st.cache_data.clear(); st.rerun()

    try:
        rt = construction_rt_id(inst, tok)
        if not rt:
            st.error("Could not find the 'Construction Advance' record type on Advance__c."); st.stop()
        tt, op, lk = st.tabs(["Turn-Time", "Open Pipeline", "Draw Lookup"])
        with tt:
            page_turn_time(inst, tok, rt, same_day, caveat_years)
        with op:
            page_pipeline(inst, tok, rt, same_day)
        with lk:
            page_lookup(inst, tok, rt, same_day)
    except Exception as exc:
        if "INVALID_SESSION_ID" in str(exc) or "Session expired" in str(exc):
            clear_salesforce_session()
            st.warning("Your Salesforce session expired. Log in again."); st.stop()
        raise


if __name__ == "__main__":
    main()
