"""
Construction Draw Tracker  (Streamlit)
======================================
ONE fused tracker — not separate Salesforce vs Land Gorilla views.

Salesforce is the spine: the Construction Advance record type IS the small-balance
RTL / fix-and-flip (RB0) book. Queried from the Advance__c OBJECT (not the pipeline
report), so paid-off draws are retained — closing the "paid off today drops off the
report" gap Melanie flagged. It carries the rb number (Loan_Number__c), the milestone
dates, Notes__c, and the real dollars.

Land Gorilla is folded in for one thing only: each advance's per-draw timeline
(submitted -> approved -> funded + amount), reached directly via the advance's
DrawContainerId__c. No noisy separate tab, no loan-list scraping.

Two lenses on the same data:
  • Pipeline (macro) — the on-demand monthly/quarter/year report: turn-time
    (complete package -> wire) beside the pre-package (borrower/inspection/title)
    interval, status, notes, both dollar figures, KPIs, by-month rollup, CSV.
  • Loan detail (micro) — search a borrower / property / loan# and see each draw's
    full cycle: the Salesforce milestone timeline + notes, with the Land Gorilla
    draw detail stacked beneath.

Secrets (.streamlit/secrets.toml): [salesforce] (OAuth, same as the AM app) and
[landgorilla] (user, password, verify).  Run:  streamlit run draw_tracker.py
"""
from __future__ import annotations

import base64
import hashlib
import io
import json
import re
import secrets
import time
from datetime import date, datetime, timezone
from typing import Any
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import numpy as np
import pandas as pd
import streamlit as st
try:
    import altair as alt
except Exception:
    alt = None
from pandas.tseries.holiday import USFederalHolidayCalendar
from simple_salesforce import Salesforce

# ───────────────────────────── domain constants ─────────────────────────────
CONSTRUCTION_DEV_NAME = "Construction_Advance"
PKG_FIELD  = "Date_Submitted_to_Capital_Partner__c"   # "Date Full Draw Package Received"
WIRE_FIELD = "Wire_Date__c"
REQ_FIELD  = "Date_Advance_Requested__c"
NOTES_FIELD = "Notes__c"
CONTAINER_FIELD = "DrawContainerId__c"
NET_FIELD, GROSS_FIELD = "Net_Funding_Total__c", "Aggregate_Funding__c"
DAYS_FIELD = "Days_to_Fund__c"        # native SF calc = request -> wire, calendar days (the headline metric)
LG_TEMPLATE_ID_DEFAULT = "437"   # a pipeline template that returns project%/funded/last-draw/risk; override in secrets

# Milestone chain — only the dates that actually populate (dead 0%-filled ones like Manager approval pruned).
MILESTONES = [
    ("Requested",                  REQ_FIELD),
    ("Inspection ordered",         "Date_Inspection_Ordered__c"),
    ("Inspection",                 "Date_Of_Inspection__c"),
    ("Inspection report received", "Date_Inspection_Report_Received__c"),
    ("Submitted for review",       "Date_Submitted_For_Approval__c"),
    ("Internal review complete",   "Date_Internal_Review_Complete__c"),
    ("Full draw package received", PKG_FIELD),
    ("Wired",                      WIRE_FIELD),
]
# Fully-populated construction dollars worth surfacing (api, friendly label).
CONSTRUCTION_FIELDS = [
    ("Renovation_Budget_Total__c",          "Renovation budget"),
    ("Approved_Renovation_Amount_Total__c", "Approved renovation"),
    ("Interest_Reserve_Total__c",           "Interest reserve"),
    ("Remaining_Interest_Reserve__c",       "Interest reserve left"),
    ("Current_UPB__c",                      "Current UPB"),
    ("Outstanding_Facility_Amount__c",      "Remaining commitment"),
    ("Total_Fees__c",                       "Total fees"),
]
WISH = (["Id", "Name", "Loan_Number__c", "Loan_Advance_Number__c", "Deal__r.Name",
         "Borrower_Name_Text__c", "Lender__c", "Status__c", "Inspection_Method__c",
         "Advance_Coordinator__r.Name", "Advance_Analyst__r.Name", NOTES_FIELD, CONTAINER_FIELD,
         NET_FIELD, GROSS_FIELD, DAYS_FIELD] + [f for f, _ in CONSTRUCTION_FIELDS] + [m[1] for m in MILESTONES])
TERMINAL = ["Completed", "Cancelled", "Rescinded", "Rejected by Capital Partner"]

_HOLS = USFederalHolidayCalendar().holidays("2018-01-01", "2032-12-31").values.astype("datetime64[D]")


# ───────────────────────────── Salesforce OAuth ─────────────────────────────
def install_truststore() -> None:
    try:
        import truststore
        truststore.inject_into_ssl()
    except Exception:
        pass


def load_sf_oauth() -> dict[str, str]:
    sec = dict(st.secrets.get("salesforce", {}))
    missing = [k for k in ("client_id", "client_secret", "redirect_uri", "auth_host") if not sec.get(k)]
    if missing:
        raise RuntimeError("Missing Salesforce OAuth secrets: " + ", ".join(missing)
                           + ". Add them under [salesforce] in Streamlit secrets.")
    sec.setdefault("scope", "api refresh_token")
    sec.setdefault("prompt", "login")
    return sec


@st.cache_resource
def _pkce_store() -> dict:
    return {}


def _pkce_pair() -> tuple[str, str]:
    v = base64.urlsafe_b64encode(secrets.token_bytes(64)).rstrip(b"=").decode()
    c = base64.urlsafe_b64encode(hashlib.sha256(v.encode()).digest()).rstrip(b"=").decode()
    return v, c


def login_url(cfg: dict[str, str]) -> str:
    state = secrets.token_urlsafe(24)
    verifier, challenge = _pkce_pair()
    store = _pkce_store(); store[state] = verifier
    if len(store) > 50:
        for old in list(store.keys())[:-50]:
            store.pop(old, None)
    q = urlencode({"response_type": "code", "client_id": cfg["client_id"],
                   "redirect_uri": cfg["redirect_uri"], "scope": cfg.get("scope", "api refresh_token"),
                   "prompt": cfg.get("prompt", "login"), "state": state,
                   "code_challenge": challenge, "code_challenge_method": "S256"})
    return f"{str(cfg['auth_host']).rstrip('/')}/services/oauth2/authorize?{q}"


def exchange_code(cfg: dict[str, str], code: str, verifier: str | None) -> dict[str, Any]:
    install_truststore()
    url = f"{str(cfg['auth_host']).rstrip('/')}/services/oauth2/token"
    fields = {"grant_type": "authorization_code", "client_id": cfg["client_id"],
              "client_secret": cfg["client_secret"], "redirect_uri": cfg["redirect_uri"], "code": code}
    if verifier:
        fields["code_verifier"] = verifier
    req = Request(url, data=urlencode(fields).encode(),
                  headers={"Content-Type": "application/x-www-form-urlencoded"}, method="POST")
    try:
        with urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())
    except HTTPError as exc:
        body = exc.read().decode("utf-8", "ignore")
        try:
            body = json.loads(body).get("error_description", body)
        except Exception:
            pass
        raise RuntimeError(f"Salesforce login failed: {body}") from exc


def _qp(name: str) -> str | None:
    v = st.query_params.get(name)
    return v[0] if isinstance(v, list) else v


def clear_sf_session() -> None:
    for k in ("salesforce_auth", "_last_sf_code"):
        st.session_state.pop(k, None)


def finish_oauth(cfg: dict[str, str]) -> None:
    if _qp("error"):
        d = _qp("error_description") or _qp("error"); st.query_params.clear()
        raise RuntimeError(f"Salesforce login was not completed: {d}")
    code = _qp("code")
    if not code:
        return
    if st.session_state.get("_last_sf_code") == code and st.session_state.get("salesforce_auth"):
        st.query_params.clear(); return
    state = _qp("state")
    verifier = _pkce_store().pop(state, None) if state else None
    if state and not verifier:
        st.query_params.clear()
        raise RuntimeError("Login could not be completed (PKCE verifier missing — the app likely "
                           "restarted). Click 'Log in to Salesforce' and try again.")
    payload = exchange_code(cfg, code, verifier)
    at, iu = payload.get("access_token"), payload.get("instance_url")
    if not at or not iu:
        raise RuntimeError("Login succeeded but no access token / instance URL returned.")
    st.session_state["salesforce_auth"] = {"access_token": at, "instance_url": iu}
    st.session_state["_last_sf_code"] = code
    st.query_params.clear(); st.rerun()


def sf_from_session() -> Salesforce | None:
    install_truststore()
    a = st.session_state.get("salesforce_auth", {})
    if not a.get("instance_url") or not a.get("access_token"):
        return None
    return Salesforce(instance_url=a["instance_url"], session_id=a["access_token"])


@st.cache_resource(show_spinner=False)
def sf_login_credentials(username: str, password: str, token: str, domain: str) -> Salesforce:
    """Username/password SOAP login — needs NO connected-app callback URL. Cached so it logs in once."""
    install_truststore()
    return Salesforce(username=username, password=password,
                      security_token=token or "", domain=domain or "login")


def sf_from_credentials() -> Salesforce | None:
    """Use [salesforce] username/password from secrets if present (the no-callback path)."""
    sec = dict(st.secrets.get("salesforce", {}))
    if sec.get("username") and sec.get("password"):
        return sf_login_credentials(sec["username"], sec["password"],
                                    sec.get("security_token", ""), sec.get("domain", "login"))
    return None


# ───────────────────────────── SF query helpers ─────────────────────────────
def soql_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("'", "\\'")


def flatten(rec: dict) -> dict:
    out: dict[str, Any] = {}
    for k, v in rec.items():
        if k == "attributes":
            continue
        if isinstance(v, dict):
            for k2, v2 in v.items():
                if k2 != "attributes":
                    out[f"{k}.{k2}"] = v2.get("Name") if isinstance(v2, dict) else v2
        else:
            out[k] = v
    return out


@st.cache_data(ttl=900, show_spinner=False)
def describe_fields(inst: str, tok: str) -> list[str]:
    sf = Salesforce(instance_url=inst, session_id=tok)
    return [f["name"] for f in sf.Advance__c.describe()["fields"]]


@st.cache_data(ttl=900, show_spinner=False)
def construction_rt(inst: str, tok: str) -> str | None:
    sf = Salesforce(instance_url=inst, session_id=tok)
    for r in sf.query("SELECT Id,DeveloperName FROM RecordType WHERE SobjectType='Advance__c'")["records"]:
        if r["DeveloperName"] == CONSTRUCTION_DEV_NAME:
            return r["Id"]
    return None


@st.cache_data(ttl=300, show_spinner=False)
def run_soql(inst: str, tok: str, soql: str) -> pd.DataFrame:
    sf = Salesforce(instance_url=inst, session_id=tok)
    return pd.DataFrame([flatten(r) for r in sf.query_all(soql)["records"]])


def select_fields(inst: str, tok: str) -> str:
    present = set(describe_fields(inst, tok))
    cols = [f for f in WISH if f.split(".")[0] in present]
    return ",".join(dict.fromkeys(cols))


# ───────────────────────────── turn-time math ───────────────────────────────
def bdays(a: pd.Series, b: pd.Series, same_day: int) -> pd.Series:
    a, b = pd.to_datetime(a, errors="coerce"), pd.to_datetime(b, errors="coerce")
    m = a.notna() & b.notna()
    out = pd.Series(np.nan, index=a.index)
    if m.any():
        out[m] = np.busday_count(a[m].values.astype("datetime64[D]"),
                                 b[m].values.astype("datetime64[D]"), holidays=_HOLS) + same_day
    out[out < 0] = np.nan
    return out


def add_intervals(df: pd.DataFrame, same_day: int) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    if PKG_FIELD in df and WIRE_FIELD in df:
        df["turn_bd"] = bdays(df[PKG_FIELD], df[WIRE_FIELD], same_day)
    if REQ_FIELD in df and PKG_FIELD in df:
        df["prepkg_bd"] = bdays(df[REQ_FIELD], df[PKG_FIELD], same_day)
    return df


def period_bounds(choice: str) -> tuple[date, date]:
    t = date.today()
    if choice == "This month":
        return t.replace(day=1), t
    if choice == "This quarter":
        q = (t.month - 1) // 3
        return date(t.year, q * 3 + 1, 1), t
    if choice == "This year":
        return date(t.year, 1, 1), t
    if choice == "Last 90 days":
        return t - pd.Timedelta(days=90), t
    if choice == "Last month":
        first = t.replace(day=1); end = first - pd.Timedelta(days=1)
        return end.replace(day=1), end.date() if hasattr(end, "date") else end
    return date(t.year, 1, 1), t


def money(x) -> str:
    try:
        v = float(x)
    except Exception:
        return "—"
    return f"${v:,.0f}"


# ───────────────────────────── Land Gorilla (draw detail only) ──────────────
def lg_config() -> dict | None:
    sec = dict(st.secrets.get("landgorilla", {}))
    return sec if sec.get("user") and sec.get("password") else None


class LGClient:
    BASE = "https://clmapi.landgorilla.com"

    def __init__(self, user, password, *, verify=True, session=None):
        self._u, self._p, self.verify = user, password, verify
        self._tok, self._exp = None, 0.0
        if session is not None:
            self._s = session
        else:
            import requests
            self._s = requests.Session()

    def token(self):
        if self._tok and time.time() < self._exp - 120:
            return self._tok
        r = self._s.get(f"{self.BASE}/api/token",
                        headers={"USER": self._u, "PASSWORD": self._p},
                        data={"api_name": "clm"}, timeout=20, verify=self.verify)
        r.raise_for_status(); j = r.json()
        self._tok = j["token"]
        try:
            self._exp = datetime.strptime(j["expired"], "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc).timestamp()
        except Exception:
            self._exp = time.time() + 1800
        return self._tok

    def get(self, path, **params):
        url = path if path.startswith("http") else f"{self.BASE}{path}"
        return self._s.get(url, headers={"Authorization": f"Bearer {self.token()}", "Accept": "application/json"},
                           params=params or None, timeout=60, verify=self.verify)


def _lg_date(v):
    if not v:
        return None
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(v)[:19], fmt).date()
        except Exception:
            continue
    return None


def parse_draw_detail(payload: dict) -> dict:
    """Pull the per-draw timeline, amount, and payee breakdown out of GET /api/clm/draw/{id}."""
    d = payload.get("data", payload) if isinstance(payload, dict) else {}
    li = d.get("lineItems") or {}
    total = (li.get("total")) or {}
    payees = []
    for name, blk in (li.get("payee") or {}).items():
        amt = ((blk or {}).get("payeeDetail") or {}).get("loan")
        payees.append({"payee": name, "amount": amt})
    rb = d.get("requestedBy") or ""
    return {
        "name": d.get("name"),
        "type": d.get("type"),
        "container_number": d.get("containerNumber"),
        "requested_by": str(rb).split(" - ")[0].strip() if rb else None,
        "status": d.get("status"),
        "created": _lg_date(d.get("createdDate")),
        "submitted": _lg_date(d.get("submittedDate")),
        "approved": _lg_date(d.get("approvedDate")),
        "funded": _lg_date(d.get("effectiveDate")),
        "amount": total.get("totalLessRetainage"),
        "requested_total": total.get("requested"),
        "payees": payees,
    }


@st.cache_resource(show_spinner=False)
def get_lg_client(user: str, password: str, verify: bool):
    return LGClient(user, password, verify=verify)


@st.cache_data(ttl=600, show_spinner=False)
def lg_draw_detail(user: str, password: str, verify: bool, draw_id: str) -> dict | None:
    """Fetch + parse one draw's detail; returns an _error dict if gated/unavailable (degrades gracefully)."""
    try:
        cli = get_lg_client(user, password, verify)
        r = cli.get(f"/api/clm/draw/{draw_id}")
        if not r.ok:
            return {"_error": f"HTTP {r.status_code}"}
        return parse_draw_detail(r.json())
    except Exception as exc:
        return {"_error": str(exc)[:80]}


def items_of(payload: Any) -> list:
    d = payload.get("data", payload) if isinstance(payload, dict) else payload
    if isinstance(d, dict):
        return d.get("items", [])
    return d if isinstance(d, list) else []


def parse_template_loan(payload: dict) -> dict:
    """Pull the loan-level construction-progress fields out of a pipeline-template loan payload."""
    f = {}
    if isinstance(payload, dict):
        f = ((payload.get("data") or {}).get("fields")) or payload.get("fields") or {}
    b = f.get("borrower") or {}
    return {
        "project_pct": f.get("projectCompletedPercentage"),
        "duration_pct": f.get("projectDurationPercentage"),
        "funded_date": _lg_date(f.get("loanFundedDate")),
        "last_draw": _lg_date(f.get("lastDrawDate")),
        "due_date": _lg_date(f.get("currentLoanDueDate")),
        "original_due": _lg_date(f.get("originalLoanDueDate")),
        "program": f.get("loanProgram"),
        "risk": [x for x in (f.get("riskLabel") or []) if x],
        "property": f.get("propertyAddress"),
        "city": f.get("city"),
        "state": f.get("state"),
        "borrower": (b.get("borrower") if isinstance(b, dict) else b),
        "last_note_at": _lg_date(f.get("dateTimeLastNote")),
        "status": ((f.get("loanStatus") or [{}])[0].get("status") if f.get("loanStatus") else None),
        "users": [u.strip() for u in (f.get("systemUsers") or []) if str(u).strip()],
    }


@st.cache_data(ttl=600, show_spinner=False)
def lg_loan_overview(user: str, password: str, verify: bool, template_id: str, loan_no: str) -> dict:
    """Resolve the LG loan by file number (rb0+loan#), then pull template progress + summary balances."""
    try:
        cli = get_lg_client(user, password, verify)
        fn = f"rb0{loan_no}"
        r = cli.get("/api/clm/loan", fileNumber=fn)
        items = items_of(r.json()) if r.ok else []
        if not items:
            return {"_error": f"no Land Gorilla loan for {fn}"}
        lid = items[0]["id"]
        out: dict[str, Any] = {"lg_loan_id": lid, "file_number": fn}
        t = cli.get(f"/api/clm/pipelineReportTemplates/{template_id}/loans/{lid}")
        if t.ok:
            out.update(parse_template_loan(t.json()))
        s = cli.get(f"/api/clm/loan/{lid}")
        if s.ok:
            sd = (s.json() or {}).get("data", {})
            out["balance"] = sd.get("loanBalance")
            out["to_finish"] = sd.get("balanceToFinish")
            out["last_approved_draw"] = _lg_date(sd.get("lastApprovedDrawEffectiveDate"))
        return out
    except Exception as exc:
        return {"_error": str(exc)[:80]}


# ───────────────────────────── Excel export ─────────────────────────────────
def build_excel(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Formatted multi-sheet .xlsx (Arial, header band, freeze panes, autofilter) as bytes for download."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    HFONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    HFILL = PatternFill("solid", fgColor="1F3864")
    AR = Font(name="Arial", size=10)
    wb = Workbook(); wb.remove(wb.active)
    for name, df in sheets.items():
        ws = wb.create_sheet(str(name)[:31])
        for j, col in enumerate(df.columns, 1):
            c = ws.cell(1, j, str(col))
            c.font = HFONT; c.fill = HFILL
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, (_, row) in enumerate(df.iterrows(), 2):
            for j, v in enumerate(row, 1):
                try:
                    isna = pd.isna(v)
                except Exception:
                    isna = False
                val = None if isna else (v.isoformat() if hasattr(v, "isoformat") else v)
                ws.cell(i, j, val).font = AR
        ws.freeze_panes = "A2"
        if len(df.columns):
            ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df)+1}"
        for j, col in enumerate(df.columns, 1):
            sample = [str(col)] + [str(x) for x in df.iloc[:, j - 1].head(40).tolist()]
            ws.column_dimensions[get_column_letter(j)].width = max(10, min(42, max(len(s) for s in sample) + 2))
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


# ───────────────────────────── Pipeline (macro) ─────────────────────────────
_PRETTY = {
    "Loan_Number__c": "Loan #", "Deal__r.Name": "Property", "Borrower_Name_Text__c": "Borrower",
    "Status__c": "Status", NOTES_FIELD: "Notes", REQ_FIELD: "Requested",
    PKG_FIELD: "Package complete", WIRE_FIELD: "Wired", "turn_bd": "Once-complete (bus. days)",
    "prepkg_bd": "Before-package (days)", DAYS_FIELD: "Days to fund", NET_FIELD: "Funded ($)",
    GROSS_FIELD: "Total funding ($)", "Advance_Coordinator__r.Name": "Coordinator",
    "Loan_Advance_Number__c": "Draw #",
}


def _period_key(w: pd.Series, gran: str) -> pd.Series:
    if gran == "Monthly":
        return w.dt.to_period("M").astype(str)
    if gran == "Quarterly":
        return w.dt.to_period("Q").astype(str)
    return w.dt.year.astype(str)


def _span_start(gran: str, n: int) -> date:
    today = date.today()
    if gran == "Monthly":
        return (pd.Timestamp(today).to_period("M") - (n - 1)).start_time.date()
    if gran == "Quarterly":
        return (pd.Timestamp(today).to_period("Q") - (n - 1)).start_time.date()
    return date(today.year - (n - 1), 1, 1)


def render_pipeline(inst, tok, rt, sel, same_day):
    st.subheader("Overview — all draws")
    GMAP = {"Month": "Monthly", "Quarter": "Quarterly", "Year": "Yearly"}
    c = st.columns([1, 1, 2])
    gran = c[0].selectbox("View by", ["Month", "Quarter", "Year"], index=1)
    gi = GMAP[gran]
    opts = {"Month": [6, 12, 24], "Quarter": [4, 8, 12], "Year": [3, 5]}[gran]
    nper = c[1].selectbox("How many to show", opts, index=1 if gran != "Year" else 0)
    start = _span_start(gi, nper)
    c[2].caption(f"Draws wired since {start:%b %-d, %Y}, grouped by {gran.lower()}." if hasattr(start, "day")
                 else f"Draws grouped by {gran.lower()}.")

    done = add_intervals(run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND {WIRE_FIELD}>={start:%Y-%m-%d}"), same_day)
    flight = run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND {WIRE_FIELD}=null "
        f"AND Status__c NOT IN ({','.join(chr(39)+s+chr(39) for s in TERMINAL)})")
    if done.empty:
        st.info("No completed draws in this range."); return

    done["Period"] = _period_key(pd.to_datetime(done[WIRE_FIELD]), gi)
    AMT = NET_FIELD if NET_FIELD in done else None
    done["_days"] = pd.to_numeric(done.get(DAYS_FIELD), errors="coerce") if DAYS_FIELD in done else np.nan
    periods = sorted(done["Period"].unique())
    cur_key = _period_key(pd.Series([pd.Timestamp(date.today())]), gi).iloc[0]

    # rollup (row-by-row so every figure reconciles to the same pull)
    rows = []
    for p in periods:
        g = done[done["Period"] == p]
        d2f = g["_days"].dropna()
        tt = g["turn_bd"].dropna() if "turn_bd" in g else pd.Series(dtype=float)
        funded = float(pd.to_numeric(g.get(AMT), errors="coerce").sum()) if AMT else None
        n = len(g)
        rows.append({
            "_raw": p,
            "Period": p + ("  (in progress)" if p == cur_key else ""),
            "Draws funded": n,
            "Loans": (g["Loan_Number__c"].nunique() if "Loan_Number__c" in g else None),
            "Funded ($)": funded,
            "Avg draw ($)": (funded / n if (funded is not None and n) else None),
            "Avg days to fund": (round(d2f.mean(), 1) if len(d2f) else None),
            "Once-complete (days)": (round(tt.mean(), 1) if len(tt) else None),
            "% measured": (round(g["_days"].notna().mean() * 100) if len(g) else 0),
        })
    roll = pd.DataFrame(rows)

    # feature the most recent COMPLETE period (not the 6-day-old current one)
    complete = [p for p in periods if p != cur_key]
    default_feat = complete[-1] if complete else periods[-1]
    rev = list(reversed(periods))
    feat = st.selectbox("Summary for", rev, index=rev.index(default_feat),
                        format_func=lambda p: p + ("  (in progress)" if p == cur_key else ""))
    fg = done[done["Period"] == feat]
    fd2f = fg["_days"].dropna()

    k = st.columns(4)
    k[0].metric(f"Draws funded · {feat}", f"{len(fg):,}")
    k[1].metric("Amount funded", money(pd.to_numeric(fg.get(AMT), errors="coerce").sum()) if AMT else "—")
    k[2].metric("Avg days to fund", f"{fd2f.mean():.0f} days" if len(fd2f) else "—",
                f"{fg['_days'].notna().mean()*100:.0f}% measured" if len(fg) else None,
                help="Calendar days from the borrower's draw request to the wire (Salesforce Days_to_Fund).")
    holds = 0
    if not flight.empty and "Status__c" in flight:
        holds = int(flight["Status__c"].astype(str).str.contains("Hold|Pending Borrower|Revision", case=False, na=False).sum())
    k[3].metric("Open draws now", f"{len(flight):,}", f"{holds} need attention" if holds else None)

    # support stat: once the package is complete, we fund fast
    tt_all = done["turn_bd"].dropna() if "turn_bd" in done else pd.Series(dtype=float)
    if len(tt_all):
        st.caption(f"**Average {done['_days'].dropna().mean():.0f} days** from request to wire across this range · "
                   f"but once the draw **package is complete**, we fund in about **{tt_all.mean():.0f} business day(s)** — "
                   "the wait is upstream (borrower, inspection, title), not in our funding.")

    st.caption(f"Across this range: **{int(roll['Draws funded'].sum()):,}** draws funded over {len(roll)} {gran.lower()}s. "
               "The newest period is marked *in progress* because it isn't finished yet — that's why its count is small.")

    # ---- rollup table ----
    st.markdown(f"**By {gran.lower()}**")
    disp = roll.drop(columns="_raw").copy()
    for c in ["Funded ($)", "Avg draw ($)"]:
        disp[c] = disp[c].map(lambda x: money(x) if pd.notna(x) else "—")
    for c in ["Loans", "Avg days to fund", "Once-complete (days)", "% measured"]:
        disp[c] = disp[c].map(lambda x: x if pd.notna(x) else "—")
    st.dataframe(disp.iloc[::-1], hide_index=True, use_container_width=True)

    # ---- charts ----
    _render_charts(done, roll, gran)

    # ---- featured-period detail + Excel ----
    st.markdown(f"**Draws funded in {feat}**")
    cols = [c for c in ["Loan_Number__c", "Loan_Advance_Number__c", "Deal__r.Name", "Borrower_Name_Text__c",
                        "Status__c", REQ_FIELD, WIRE_FIELD, DAYS_FIELD, NET_FIELD, NOTES_FIELD] if c in fg]
    detail = fg[cols].rename(columns=_PRETTY).sort_values("Wired", ascending=False)
    st.dataframe(detail, use_container_width=True, height=340)
    xlsx = build_excel({f"By {gran.lower()}": roll.drop(columns="_raw"), f"Draws {feat}": detail})
    st.download_button("⬇️ Download Excel", xlsx, file_name=f"construction_draws_{feat}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.caption('"Avg days to fund" is calendar days from the borrower\'s draw request to the wire (Salesforce\'s own '
               'Days_to_Fund). "% measured" shows how many draws in the period carry that value. Paid-off draws are '
               "included (pulled from the Advance object, not the pipeline report).")


def _render_charts(done: pd.DataFrame, roll: pd.DataFrame, gran: str):
    """Four overview charts built on average days-to-fund + funded dollars."""
    d2f = done["_days"].dropna()
    if alt is None:
        # graceful fallback without altair
        st.bar_chart(roll.set_index("_raw")["Draws funded"])
        if "Avg days to fund" in roll:
            st.line_chart(roll.set_index("_raw")["Avg days to fund"])
        return

    r = roll.rename(columns={"_raw": "Period"})
    c1, c2 = st.columns(2)

    # 1) days-to-fund distribution (shows the tail), with a mean line
    with c1:
        st.markdown("**How long draws take (request → wire)**")
        if len(d2f):
            capped = d2f.clip(upper=45)
            hist = alt.Chart(pd.DataFrame({"days": capped})).mark_bar(color="#4C78A8").encode(
                x=alt.X("days:Q", bin=alt.Bin(maxbins=30), title="Calendar days (capped at 45)"),
                y=alt.Y("count()", title="Draws"))
            rule = alt.Chart(pd.DataFrame({"m": [d2f.mean()]})).mark_rule(color="#E45756", size=2).encode(x="m:Q")
            st.altair_chart(hist + rule, use_container_width=True)
            st.caption(f"Red line = average ({d2f.mean():.0f} days). The long right tail is borrower-driven delay.")
        else:
            st.info("No days-to-fund data in range.")

    # 2) average days-to-fund trend over time
    with c2:
        st.markdown("**Average days to fund over time**")
        tr = r.dropna(subset=["Avg days to fund"])
        if not tr.empty:
            line = alt.Chart(tr).mark_line(point=True, color="#E45756").encode(
                x=alt.X("Period:N", sort=list(tr["Period"]), title=""),
                y=alt.Y("Avg days to fund:Q", title="Avg days"))
            st.altair_chart(line, use_container_width=True)
        else:
            st.info("No trend data.")

    c3, c4 = st.columns(2)

    # 3) funded dollars per period
    with c3:
        st.markdown("**Amount funded per period**")
        bars = alt.Chart(r).mark_bar(color="#54A24B").encode(
            x=alt.X("Period:N", sort=list(r["Period"]), title=""),
            y=alt.Y("Funded ($):Q", title="Funded ($)"),
            tooltip=["Period", alt.Tooltip("Funded ($):Q", format="$,.0f")])
        st.altair_chart(bars, use_container_width=True)

    # 4) draws + avg days combo (dual axis)
    with c4:
        st.markdown("**Draws funded & average days**")
        base = alt.Chart(r).encode(x=alt.X("Period:N", sort=list(r["Period"]), title=""))
        bars = base.mark_bar(color="#B9D7A8").encode(y=alt.Y("Draws funded:Q", title="Draws"))
        line = base.mark_line(color="#E45756", point=True).encode(
            y=alt.Y("Avg days to fund:Q", axis=alt.Axis(title="Avg days", orient="right")))
        st.altair_chart(alt.layer(bars, line).resolve_scale(y="independent"), use_container_width=True)


# ───────────────────────────── Loan detail (micro) ──────────────────────────
def lg_filenumber(loan_no) -> str | None:
    """LG file number is 'rb0'+the SF loan number, but only for the clean numeric RB0 book.
       Hyphenated/alpha loan numbers aren't RB0 loans, so skip LG for them (keeps it quiet)."""
    s = str(loan_no).strip()
    return f"rb0{s}" if s.isdigit() else None


def render_loan_detail(inst, tok, rt, sel, same_day):
    st.markdown("## 🔎  Find a draw")
    st.caption("Search by **loan number**, **property / deal name**, **borrower**, or **draw number** — "
               "type any part of it.")
    text = st.text_input("search", label_visibility="collapsed",
                         placeholder="e.g.   64806      116 South Street      Panache Properties")
    if not text:
        st.info("Start typing above to pull up a draw and its full cycle.")
        return

    q = soql_escape(text)
    searchable = [f for f in ["Loan_Number__c", "Deal__r.Name", "Borrower_Name_Text__c",
                              "Loan_Advance_Number__c"] if f in sel]
    ors = " OR ".join(f"{f} LIKE '%{q}%'" for f in searchable) or "Id != null"
    df = add_intervals(run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND ({ors}) "
        f"ORDER BY Loan_Number__c, {REQ_FIELD}"), same_day)
    if df.empty:
        st.warning("No matching draws."); return
    nloans = df["Loan_Number__c"].nunique() if "Loan_Number__c" in df else len(df)
    st.caption(f"Found **{len(df)}** draw(s) across **{nloans}** loan(s).")

    cfg = lg_config()
    tmpl = (cfg or {}).get("template_id", LG_TEMPLATE_ID_DEFAULT)
    for loan_no, g in df.groupby("Loan_Number__c", dropna=False):
        prop = g["Deal__r.Name"].dropna().iloc[0] if "Deal__r.Name" in g and g["Deal__r.Name"].notna().any() else ""
        borrower = g["Borrower_Name_Text__c"].dropna().iloc[0] if "Borrower_Name_Text__c" in g and g["Borrower_Name_Text__c"].notna().any() else ""
        st.markdown(f"### Loan {loan_no} — {prop}")
        sub = []
        if borrower: sub.append(f"Borrower: {borrower}")
        sub.append(f"{len(g)} draw(s)")
        st.caption("  ·  ".join(sub))

        # --- Salesforce leads: each draw's cycle ---
        for _, row in g.iterrows():
            _render_draw_sf(row)

        # --- Land Gorilla: secondary, collapsed so it doesn't clutter ---
        if cfg:
            with st.expander("🦍  Land Gorilla details", expanded=False):
                _render_lg_for_loan(cfg, tmpl, loan_no, g)
        st.divider()


def _render_draw_sf(row: pd.Series):
    draw_no = row.get("Loan_Advance_Number__c") or row.get("Name") or "draw"
    top = st.columns([2, 1, 1])
    top[0].markdown(f"**Draw {draw_no}** — {row.get('Status__c','')}")
    if NET_FIELD in row and pd.notna(row.get(NET_FIELD)):
        top[1].metric("Funded", money(row[NET_FIELD]))
    if DAYS_FIELD in row and pd.notna(row.get(DAYS_FIELD)):
        top[2].metric("Days to fund", f"{float(row[DAYS_FIELD]):.0f}")
    elif "turn_bd" in row and pd.notna(row.get("turn_bd")):
        top[2].metric("Once-complete", f"{row['turn_bd']:.0f} bd")

    steps = [{"Milestone": lbl, "Date": (pd.to_datetime(row.get(f)).date() if pd.notna(row.get(f)) else None),
              "": "✅" if pd.notna(row.get(f)) else "⬜"} for lbl, f in MILESTONES]
    tdf = pd.DataFrame(steps)
    recorded = tdf["Date"].notna().sum()
    st.progress(recorded / len(MILESTONES), text=f"Milestones recorded: {recorded}/{len(MILESTONES)}")
    st.dataframe(tdf, hide_index=True, use_container_width=True,
                 column_config={"": st.column_config.TextColumn(width="small")})

    # construction dollars (only show the ones that are populated / non-zero)
    ctx = []
    for f, lbl in CONSTRUCTION_FIELDS:
        v = row.get(f)
        try:
            if pd.notna(v) and float(v) != 0:
                ctx.append(f"{lbl}: {money(v)}")
        except Exception:
            pass
    if ctx:
        st.caption("  ·  ".join(ctx))
    note = row.get(NOTES_FIELD)
    if pd.notna(note) and str(note).strip():
        st.caption(f"📝 {note}")


def _render_lg_for_loan(cfg: dict, tmpl: str, loan_no, g: pd.DataFrame):
    fn = lg_filenumber(loan_no)
    if fn:
        ov = lg_loan_overview(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), str(tmpl), str(loan_no))
        if ov and not ov.get("_error"):
            m = st.columns(4)
            if ov.get("project_pct") is not None:
                pct = float(ov["project_pct"]); m[0].metric("Project complete", f"{pct:.0f}%")
                m[0].progress(min(max(pct / 100, 0), 1.0))
                if ov.get("duration_pct") is not None:
                    m[0].caption(f"{float(ov['duration_pct']):.0f}% of term elapsed")
            if ov.get("to_finish") is not None: m[1].metric("Remaining to fund", money(ov["to_finish"]))
            if ov.get("balance") is not None: m[2].metric("Loan balance", money(ov["balance"]))
            if ov.get("last_draw"): m[3].metric("Last draw", f"{ov['last_draw']}")
            bits = []
            for lbl, key in [("Status", "status"), ("Program", "program"),
                             ("Funded", "funded_date"), ("Due", "due_date"), ("Last note", "last_note_at")]:
                if ov.get(key): bits.append(f"**{lbl}:** {ov[key]}")
            loc = " ".join(x for x in [ov.get("city"), ov.get("state")] if x)
            if loc: bits.append(f"**Location:** {loc}")
            if bits: st.caption("  ·  ".join(bits))
            if ov.get("risk"): st.caption("**Risk:** " + " · ".join(ov["risk"]))
            if ov.get("users"): st.caption("**LG team:** " + ", ".join(ov["users"]))
        else:
            st.caption(f"No matching loan in Land Gorilla ({fn}).")
    else:
        st.caption("This loan number isn't an RB0 Land Gorilla file, so there's no Land Gorilla match.")

    # --- Land Gorilla draws for this loan, as ONE table with proper columns ---
    # Columns mirror the GET /api/clm/draw/{id} response exactly:
    #   name, type, status, createdDate, submittedDate, approvedDate, effectiveDate(funded), total amount.
    draw_rows, all_payees = [], []
    for _, row in g.iterrows():
        container = row.get(CONTAINER_FIELD)
        if pd.isna(container) or not container:
            continue
        det = lg_draw_detail(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), str(container))
        sf_draw = row.get("Loan_Advance_Number__c") or row.get("Name", "")
        if not det or det.get("_error"):
            draw_rows.append({"Draw": sf_draw, "LG draw": "—", "Type": "—",
                              "Status": (det or {}).get("_error", "unavailable"),
                              "Created": None, "Submitted": None, "Approved": None, "Funded": None, "Amount": None})
            continue
        draw_rows.append({
            "Draw": sf_draw,
            "LG draw": det.get("name") or "—",
            "Type": det.get("type") or "—",
            "Status": det.get("status") or "—",
            "Created": det.get("created"),
            "Submitted": det.get("submitted"),
            "Approved": det.get("approved"),
            "Funded": det.get("funded"),
            "Amount": det.get("amount"),
        })
        for p in (det.get("payees") or []):
            all_payees.append({"Draw": det.get("name") or sf_draw, "Payee": p["payee"], "Amount": p.get("amount")})

    if draw_rows:
        st.markdown("**Land Gorilla draws**")
        ddf = pd.DataFrame(draw_rows)
        ddf["Amount"] = ddf["Amount"].map(lambda x: money(x) if pd.notna(x) else "—")
        st.dataframe(
            ddf, hide_index=True, use_container_width=True,
            column_config={
                "Created": st.column_config.DateColumn("Created", format="MM/DD/YYYY"),
                "Submitted": st.column_config.DateColumn("Submitted", format="MM/DD/YYYY"),
                "Approved": st.column_config.DateColumn("Approved", format="MM/DD/YYYY"),
                "Funded": st.column_config.DateColumn("Funded", format="MM/DD/YYYY"),
            })
        if all_payees:
            with st.expander(f"Payees across these draws ({len(all_payees)})"):
                pdf = pd.DataFrame(all_payees)
                pdf["Amount"] = pdf["Amount"].map(lambda x: money(x) if pd.notna(x) else "—")
                st.dataframe(pdf, hide_index=True, use_container_width=True)
    else:
        st.caption("No Land Gorilla draw containers on these advances yet.")


# ───────────────────────────────── main ─────────────────────────────────────
def main():
    st.set_page_config(page_title="Construction Draw Tracker", page_icon="🏗️", layout="wide")
    st.title("🏗️ Construction Draw Tracker")

    # --- Auth: prefer username/password (NO callback URL); fall back to OAuth redirect ---
    cfg = None; err = None; sf = None; mode = None
    try:
        sf = sf_from_credentials()
        if sf is not None:
            mode = "username/password"
            st.session_state["salesforce_auth"] = {
                "instance_url": f"https://{sf.sf_instance}", "access_token": sf.session_id}
    except Exception as exc:
        err = f"Salesforce username/password login failed: {exc}"
    if sf is None and err is None:                      # no creds -> OAuth redirect fallback
        try:
            cfg = load_sf_oauth(); finish_oauth(cfg)
            sf = sf_from_session(); mode = "oauth"
        except Exception as exc:
            err = str(exc)

    with st.sidebar:
        st.header("Salesforce")
        if err:
            st.error(err)
        elif sf is None:
            st.info("Not connected")
        else:
            st.success(f"Connected · {mode}")
            st.caption(st.session_state.get("salesforce_auth", {}).get("instance_url", ""))
            if mode == "oauth" and st.button("Log out", use_container_width=True):
                clear_sf_session(); st.rerun()
        st.divider()
        same_day = 1 if st.radio("Same-day convention", ["0 business days", "1 business day"],
                                 help="Package in & wired same day counts as this. Confirm with Melanie — "
                                      "it moves the median.").startswith("1") else 0
        st.divider()
        st.caption("Land Gorilla: " + ("configured ✅" if lg_config() else "not set"))

    if err:
        st.error(err)
        st.caption("For the no-callback path, put username / password / security_token under "
                   "[salesforce] in secrets. For OAuth, provide client_id / client_secret / redirect_uri / auth_host.")
        st.stop()
    if sf is None:                                       # OAuth path, not yet logged in
        st.subheader("Step 1 — Log in to Salesforce")
        st.info("Log in to load the construction draw book.")
        st.link_button("Log in to Salesforce", login_url(cfg))
        st.caption(f"Callback URL: {cfg['redirect_uri']}")
        st.stop()

    inst = st.session_state["salesforce_auth"]["instance_url"]
    tok = st.session_state["salesforce_auth"]["access_token"]
    rt = construction_rt(inst, tok)
    if not rt:
        st.error("Could not find the 'Construction Advance' record type."); st.stop()
    sel = select_fields(inst, tok)

    page = st.sidebar.radio("View", ["Pipeline (macro)", "Loan detail (micro)"])
    try:
        if page.startswith("Pipeline"):
            render_pipeline(inst, tok, rt, sel, same_day)
        else:
            render_loan_detail(inst, tok, rt, sel, same_day)
    except Exception as exc:
        msg = str(exc)
        if "INVALID_SESSION_ID" in msg or "Session expired" in msg:
            try:
                sf_login_credentials.clear()          # force a fresh login on the credentials path
            except Exception:
                pass
            clear_sf_session()
            st.warning("Your Salesforce session expired — reloading."); st.stop()
        raise


if __name__ == "__main__":
    main()
