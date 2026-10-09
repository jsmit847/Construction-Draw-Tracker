"""
Construction Draw Tracker  (Streamlit)
======================================
ONE fused tracker — not separate Salesforce vs Land Gorilla views.

Salesforce is the spine: the Construction Advance record type IS the small-balance
RTL / fix-and-flip (RB0) book. Queried from the Advance__c OBJECT (not the pipeline
report), so paid-off draws are retained — closing the "paid off today drops off the
report" gap Melanie flagged. It carries the rb number (Loan_Number__c), the milestone
dates, Notes__c, and the real dollars.

Land Gorilla is the source of truth for TIMING: Salesforce collapses the request /
package dates onto the wire date, so each draw's real request (createdDate) and
complete-package (approvedDate) dates come from Land Gorilla via DrawContainerId__c,
and the wire date from Salesforce. All spans are BUSINESS days.

Two lenses on the same data, plus an export:
  • Pipeline (macro) — monthly/quarter/year report: borrower wait (request -> package),
    our funding (package -> wire), KPIs, rollup, charts, turn-time Excel.
  • Loan detail (micro) — search a borrower / property / loan# and see each draw's
    full cycle: timing, the Salesforce milestones + notes, Land Gorilla detail beneath.
  • Funding-request file — Melanie's weekly "Construction/Renovation Funding Request"
    workbook, filled from Salesforce into the committed template (Advance_template.xlsx).

Secrets (.streamlit/secrets.toml): [salesforce] (username/password/security_token, or
OAuth client_id/client_secret/redirect_uri/auth_host) and [landgorilla] (user, password,
verify).  Run:  streamlit run draw_dashboard.py
"""
from __future__ import annotations

import base64
import hashlib
import io
import os
import json
import re
import secrets
import time
from datetime import date, datetime, timedelta, timezone
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
WISH = (["Id", "Name", "Loan_Number__c", "Loan_Advance_Number__c",
         "Deal__r.Name", "Deal__r.Account.Name", "Borrower_Name__c", "Borrower_Name_Text__c",
         "Lender__c", "Status__c", "Inspection_Method__c",
         "Advance_Coordinator__r.Name", "Advance_Analyst__r.Name", NOTES_FIELD, CONTAINER_FIELD,
         NET_FIELD, GROSS_FIELD] + [f for f, _ in CONSTRUCTION_FIELDS] + [m[1] for m in MILESTONES])
TERMINAL = ["Completed", "Cancelled", "Rescinded", "Rejected by Capital Partner"]
OPEN_WHERE = f"{WIRE_FIELD}=null AND Status__c NOT IN ({','.join(repr(s) for s in TERMINAL)})"
# Status__c stages a weekly funding request covers (approval through release) — the export's default filter.
FUNDING_STAGES = ("Pending Approval", "Approved", "Pending Capital Partner Release", "Pending Release")
SEARCH_LIMIT = 200    # loan-detail search: most draws rendered (and dated in Land Gorilla) per query
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

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
        d = re.sub(r"[^\w .,:;'/-]", "", str(d))[:200]      # plain text only — the URL is attacker-controllable
        raise RuntimeError(f"Salesforce login was not completed: {d}")
    code = _qp("code")
    if not code:
        return
    if st.session_state.get("_last_sf_code") == code and st.session_state.get("salesforce_auth"):
        st.query_params.clear(); return
    state = _qp("state")
    verifier = _pkce_store().pop(state, None) if state else None
    if not verifier:                                         # state is required: it binds the code to our login
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


@st.cache_resource(show_spinner=False)
def _login_failures() -> dict:
    return {}


def sf_from_credentials() -> Salesforce | None:
    """Use [salesforce] username/password from secrets if present (the no-callback path). A failed login is
       remembered for 15 minutes so every rerun doesn't retry it and lock the account; new secrets retry at once."""
    sec = dict(st.secrets.get("salesforce", {}))
    if sec.get("username") and sec.get("password"):
        key = hashlib.sha256("\0".join(str(sec.get(k, "")) for k in
                                       ("username", "password", "security_token", "domain")).encode()).hexdigest()
        failed = _login_failures().get(key)
        if failed and time.time() - failed[0] < 900:
            raise RuntimeError(failed[1])
        try:
            return sf_login_credentials(sec["username"], sec["password"],
                                        sec.get("security_token", ""), sec.get("domain", "login"))
        except Exception as exc:
            _login_failures()[key] = (time.time(), str(exc)[:300])
            raise
    return None


# ───────────────────────────── SF query helpers ─────────────────────────────
def soql_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("'", "\\'")


def flatten(rec: dict, prefix: str = "") -> dict:
    """Recursively flatten SF relationship objects to dotted leaf keys (Deal__r.Account.Name, etc.)."""
    out: dict[str, Any] = {}
    for k, v in rec.items():
        if k == "attributes":
            continue
        key = f"{prefix}{k}"
        if isinstance(v, dict):
            out.update(flatten(v, prefix=key + "."))
        else:
            out[key] = v
    return out


@st.cache_data(ttl=900, show_spinner=False)
def describe_object(inst: str, tok: str, obj: str) -> dict:
    """{field: {"type", "ref" (objects it points to), "rel" (relationship name)}} for one sObject."""
    sf = Salesforce(instance_url=inst, session_id=tok)
    return {f["name"]: {"type": f.get("type"), "ref": list(f.get("referenceTo") or []),
                        "rel": f.get("relationshipName")}
            for f in getattr(sf, obj).describe()["fields"]}


def describe_fields(inst: str, tok: str) -> list[str]:
    return list(describe_object(inst, tok, "Advance__c"))


def resolve_path(inst: str, tok: str, obj: str, path: str) -> dict | None:
    """Describe metadata of the field a dotted SOQL path (e.g. Deal__r.Account.Name) ends on, or None if it
       doesn't resolve from obj — lets optional columns drop out one by one instead of breaking a whole query."""
    head, _, rest = path.partition(".")
    try:
        fields = describe_object(inst, tok, obj)
    except Exception:
        return None
    if not rest:
        return fields.get(head)
    ref = next((m["ref"] for m in fields.values() if m["rel"] == head and m["ref"]), None)
    return resolve_path(inst, tok, ref[0], rest) if ref else None


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

    def ok(f: str) -> bool:
        seg = f.split(".")[0]                     # first path segment
        if seg in present:
            return True
        if seg.endswith("__r") and (seg[:-3] + "__c") in present:   # custom relationship -> its __c field
            return True
        if seg in ("Deal",) and "Deal__c" in present:               # safety alias
            return True
        return False

    cols = [f for f in WISH if ok(f)]
    return ",".join(dict.fromkeys(cols))


# ───────────────────────────── turn-time math ───────────────────────────────
def bdays(a: pd.Series, b: pd.Series, same_day: int) -> pd.Series:
    a, b = pd.to_datetime(a, errors="coerce"), pd.to_datetime(b, errors="coerce")
    m = a.notna() & b.notna()
    out = pd.Series(np.nan, index=a.index)
    if m.any():
        n = np.busday_count(a[m].values.astype("datetime64[D]"),
                            b[m].values.astype("datetime64[D]"), holidays=_HOLS)
        out[m] = np.where(n >= 0, n + same_day, np.nan)     # b before a -> blank (checked before the offset)
    return out


def add_intervals(df: pd.DataFrame) -> pd.DataFrame:
    """Blank timing columns, filled from Land Gorilla by apply_lg_correction. Salesforce's own request/package
       dates (and its Days_to_Fund__c) are collapsed onto the wire date, so they never feed timing."""
    if df.empty:
        return df
    df = df.copy()
    for col in ("_days", "_wait_bd", "_fund_bd"):
        df[col] = np.nan
    for col in ("_lg_req", "_lg_pkg"):
        df[col] = None
    df["_days_src"] = ""
    return df


LG_DATES_TTL = 86400   # a wired draw's LG created/approved dates don't change — keep them a day


@st.cache_resource(show_spinner=False)
def _lg_dates_store() -> dict:
    """Process-wide {container: (fetched_at, created_iso, approved_iso)}. Only successful fetches are stored,
       so a draw that failed (rate limit / timeout) is fetched again next time instead of staying blank."""
    return {}


def lg_draw_dates(user: str, password: str, verify: bool, containers) -> dict:
    """{container: (created_iso, approved_iso | None)} from GET /api/clm/draw/{id}, fetched CONCURRENTLY —
       ~635 serial calls took ~80s; parallel takes a few seconds. Repeats come from the store (no refetch)."""
    from concurrent.futures import ThreadPoolExecutor
    import requests, threading

    store, now = _lg_dates_store(), time.time()
    want = [str(c) for c in dict.fromkeys(containers) if c]
    need = [c for c in want if not (c in store and now - store[c][0] < LG_DATES_TTL)]
    if need:
        client = LGClient(user, password, verify=verify)   # dedicated client for this batch
        token = client.token()                              # pre-warm the token so threads don't race refreshing it
        tl = threading.local()                              # one requests.Session per worker thread (thread-safe)

        def one(container):
            s = getattr(tl, "s", None)
            if s is None:
                s = tl.s = requests.Session()
            try:
                r = s.get(f"{client.BASE}/api/clm/draw/{container}",
                          headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
                          timeout=30, verify=verify)
                det = parse_draw_detail(r.json()) if r.ok else {}
            except Exception:
                det = {}
            if not det.get("created"):
                return container, None
            approved = det.get("approved")                  # the complete-package point
            return container, (time.time(), det["created"].isoformat(), approved.isoformat() if approved else None)

        with ThreadPoolExecutor(max_workers=16) as ex:
            for container, val in ex.map(one, need):
                if val is not None:
                    store[container] = val
    return {c: store[c][1:] for c in want if c in store}


def apply_lg_correction(df: pd.DataFrame, cfg: dict | None, same_day: int) -> pd.DataFrame:
    """Fill each wired draw's timing from Land Gorilla, in business days: request = createdDate,
       complete package = approvedDate, wire = Salesforce Wire_Date__c. Draws LG can't date stay blank."""
    if df.empty or not cfg or CONTAINER_FIELD not in df or WIRE_FIELD not in df:
        return df
    df = df.copy()
    sub = df[df[CONTAINER_FIELD].notna() & df[WIRE_FIELD].notna()]
    try:
        lg = lg_draw_dates(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), sub[CONTAINER_FIELD].tolist())
    except Exception as exc:          # no exception text: a malformed secret would be echoed to every viewer
        st.warning(f"Land Gorilla is unavailable right now ({type(exc).__name__}), so timing is blank.")
        return df
    got = sub[CONTAINER_FIELD].map(lambda c: lg.get(str(c)) or (None, None))
    created, approved, wire = got.map(lambda t: t[0]), got.map(lambda t: t[1]), sub[WIRE_FIELD]
    total = bdays(created, wire, same_day)            # blank when LG has no date or the wire precedes it
    ok = total.index[total.notna()]
    df.loc[ok, "_days"] = total[ok]
    df.loc[ok, "_days_src"] = "Land Gorilla"
    df.loc[ok, "_lg_req"] = created[ok]
    df.loc[ok, "_lg_pkg"] = approved[ok]
    # Split only when request <= package <= wire (an approval stamped after the wire can't be apportioned).
    # The same-day convention applies to package -> wire, so wait + funding = total.
    c, a, w = (pd.to_datetime(s, errors="coerce").dt.normalize() for s in (created, approved, wire))
    split = ok.intersection(total.index[(c <= a) & (a <= w)])
    df.loc[split, "_wait_bd"] = bdays(created, approved, 0)[split]         # borrower / inspection / title side
    df.loc[split, "_fund_bd"] = bdays(approved, wire, same_day)[split]     # CoreVest's own funding speed
    return df


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
    d = pd.to_datetime(str(v), errors="coerce")        # any other layout (e.g. "2026-09-01 10:00:00")
    return None if pd.isna(d) else d.date()


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
        return {"_error": f"unavailable ({type(exc).__name__})"}


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
        return {"_error": type(exc).__name__}


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
                if isinstance(val, str):
                    val = _xl_text(val)
                _put(ws, i, j, val)
                ws.cell(i, j).font = AR
        ws.freeze_panes = "A2"
        if len(df.columns):
            ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df)+1}"
        for j, col in enumerate(df.columns, 1):
            sample = [str(col)] + [str(x) for x in df.iloc[:, j - 1].head(40).tolist()]
            ws.column_dimensions[get_column_letter(j)].width = max(10, min(42, max(len(s) for s in sample) + 2))
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


# ───────────────── Melanie's funding-request format (template fill) ──────────
# The weekly "CoreVest Construction/Renovation Funding Request" — the file Salesforce's Generate File button
# builds for the logged-in user. Its grain is ONE ROW PER ADVANCE (a property's draw): a portfolio loan gets a
# row per property. Each template column maps to a SOQL path from Advance__c, or to candidates tried in order
# (from the org's field glossary). Paths that don't resolve in the org, or that resolve to a raw lookup Id, drop
# out one by one (resolve_path), so a wrong guess blanks one column, not the file.
MELANIE_MAP = {
    "B": "Loan_Number__c",                                          # CV Loan #
    "C": "Deal__r.Name",                                            # Deal Name
    "D": "Borrower_Name_Text__c",                                   # Borrower Name
    "E": ("Deal__r.Contact__r.Name", "Deal__r.Sponsor_Entity__c"),  # Sponsor(s) — a person in her file (Primary Contact)
    "J": "Deal__r.Project_Strategy__c",                             # Project Type (Ground Up / Fix and Flip / …)
    "M": ("Purchase_Funded_Date__c", "Deal__r.CloseDate"),          # Funded Date
    "N": ("Deal__r.Updated_Loan_Maturity_Date__c",                  # Maturity Date (current, after extensions)
          "Deal__r.Current_Line_Maturity_Date__c"),
    "O": ("Deal__r.LOC_Commitment__c", "LOC_Commitment__c"),        # Total Loan Commitment (a Deal field)
    "U": NET_FIELD,                                                 # Current Draw Amount — the net wire; feeds E6
    "AC": "Deal__r.Next_Payment_Date__c",                           # Next Due Date
    "AD": "Remaining_Interest_Reserve__c",                          # IR Balance
    "AF": NOTES_FIELD,                                              # Comments
    "AH": ("Deal__r.Warehouse_Line__c", "Warehouse_Line__c"),       # Warehouse line (a Deal field)
    "AJ": "Deal__r.Owner.Name",                                     # RM/LO — the Deal owner ("CAF Originator" here)
}
# From the advance's Property__c: Advance__c's own Property lookup if it has one; else the Deal's property when the
# Deal has just one, or the single property whose Advance__c points at this draw. Otherwise blank — never a guess.
MELANIE_PROPERTY_MAP = {
    "G": "City__c",                                                       # Property City
    "H": "State__c",                                                      # ST
    "I": "Property_Type__c",                                              # Property Type
    "K": ("Origination_Date_Value__c", "Appraised_Value_Amount__c"),       # As-Is Value at Origination
    "L": ("Origination_After_Repair_Value__c", "After_Repair_Value__c"),  # As Complete Value at Origination
    "Q": "Initial_Disbursement__c",                                       # Initial Loan Funding
    "AB": "Current_UPB__c",                                               # Current UPB
}
STREET_PARTS = ("Street_Number__c", "Street_Name__c")   # F Property Street = number + name
# Computed columns written as live Excel formulas — only where every input cell is filled.
MELANIE_FORMULAS = {"R": "=Q{r}/K{r}",           # LTV% As-is = initial funding / as-is value
                    "S": "=O{r}/L{r}",           # LTV% As Complete = commitment / as-complete value
                    "P": "=O{r}/T{r}-O{r}",      # Borrower Equity (her formula)
                    "X": "=V{r}/(V{r}+W{r})"}    # Const. Loan Disbursement % (her formula)
TEMPLATE_CANDIDATES = ("Advance_template.xlsx", "melanie_template.xlsx",
                       "Construction-Renovation_Funding_Request_template.xlsx")
TEMPLATE_SHEET, TEMPLATE_HEADER_ROW = "Combined Funding Request", 7
PERSON_FIELDS = ("Advance_Coordinator__r.Name", "Advance_Analyst__r.Name")


def _find_template() -> str | None:
    here = os.path.dirname(os.path.abspath(__file__))
    return next((p for p in (os.path.join(here, n) for n in TEMPLATE_CANDIDATES) if os.path.exists(p)), None)


def _soql_in(inst: str, tok: str, select: str, obj: str, key: str, values) -> pd.DataFrame:
    """SELECT … WHERE key IN (…), 200 values per query to stay under Salesforce's URL-length limit."""
    vals = [str(v) for v in dict.fromkeys(values) if pd.notna(v) and str(v)]
    parts = [run_soql(inst, tok, f"SELECT {select} FROM {obj} WHERE {key} IN ('"
                      + "','".join(soql_escape(v) for v in vals[i:i + 200]) + "')")
             for i in range(0, len(vals), 200)]
    parts = [p for p in parts if not p.empty]
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def _cands(v) -> tuple:
    return (v,) if isinstance(v, str) else tuple(v)


def _usable(meta: dict | None) -> bool:
    return bool(meta) and meta["type"] not in ("reference", "id")      # never write a raw record Id


@st.cache_data(ttl=300, show_spinner=False)
def funding_request_rows(inst: str, tok: str, rt: str, where: str) -> tuple[pd.DataFrame, dict, list]:
    """Advances in scope, one row each, keyed by template column letter (plus the filter fields).
       Returns (rows, {column letter: Salesforce type}, [template columns this org can't fill])."""
    adv_meta = describe_object(inst, tok, "Advance__c")
    cols, ftypes = {}, {}
    for col, cands in MELANIE_MAP.items():
        for path in _cands(cands):
            meta = resolve_path(inst, tok, "Advance__c", path)
            if _usable(meta):
                cols[col], ftypes[col] = path, meta["type"]
                break
    unavailable = [col for col in MELANIE_MAP if col not in cols]
    prop_ref = next((f for f, m in adv_meta.items() if m["ref"] == ["Property__c"]), None)  # Advance -> Property
    extra = [p for p in ("Id", "Deal__c", "Status__c", "Loan_Advance_Number__c") + PERSON_FIELDS
             if resolve_path(inst, tok, "Advance__c", p)] + ([prop_ref] if prop_ref else [])
    adv = run_soql(inst, tok, f"SELECT {','.join(dict.fromkeys(extra + list(cols.values())))} FROM Advance__c "
                              f"WHERE RecordTypeId='{rt}' AND {where} ORDER BY Loan_Number__c")
    if adv.empty:
        return adv, ftypes, unavailable
    out = adv[[c for c in extra if c in adv]].copy()
    for col, path in cols.items():
        out[col] = adv[path] if path in adv else None
    out["_multi_property"] = False

    # Property columns — validated field by field, so one missing field can't blank the rest.
    try:
        pmeta = describe_object(inst, tok, "Property__c")
    except Exception:
        pmeta = {}
    pcols = {col: f for col, cands in MELANIE_PROPERTY_MAP.items()
             for f in [next((f for f in _cands(cands) if _usable(pmeta.get(f))), None)] if f}
    street = [f for f in STREET_PARTS if f in pmeta]
    unavailable += [col for col in MELANIE_PROPERTY_MAP if col not in pcols] + ([] if street else ["F"])
    pfields = list(dict.fromkeys(list(pcols.values()) + street))
    link = "Advance__c" if (pmeta.get("Advance__c") or {}).get("ref") == ["Advance__c"] else None  # Property -> draw
    match = [None] * len(adv)                                   # each advance's property record, when known
    if pfields and prop_ref and prop_ref in adv:                # the advance names its own property
        by_id = {r["Id"]: r for r in _soql_in(inst, tok, ",".join(["Id"] + pfields), "Property__c", "Id",
                                              adv[prop_ref]).to_dict("records")}
        match = [by_id.get(p) for p in adv[prop_ref]]
    elif pfields and "Deal__c" in pmeta and "Deal__c" in adv:  # else via the Deal
        sel = ",".join(dict.fromkeys(["Id", "Deal__c"] + ([link] if link else []) + pfields))
        per_deal, per_draw = {}, {}
        for r in _soql_in(inst, tok, sel, "Property__c", "Deal__c", adv["Deal__c"]).to_dict("records"):
            per_deal.setdefault(r["Deal__c"], []).append(r)
            if link and pd.notna(r.get(link)):
                per_draw.setdefault(r[link], []).append(r)
        ids = adv["Id"] if "Id" in adv else pd.Series([None] * len(adv))
        for i, (aid, deal) in enumerate(zip(ids, adv["Deal__c"])):
            props, hits = per_deal.get(deal, []), per_draw.get(aid, [])
            match[i] = props[0] if len(props) == 1 else (hits[0] if len(hits) == 1 else None)
        out["_multi_property"] = [len(per_deal.get(d, [])) > 1 and m is None for d, m in zip(adv["Deal__c"], match)]
    for col, f in pcols.items():
        out[col] = [m.get(f) if m else None for m in match]
        ftypes[col] = pmeta[f]["type"]
    if street:                                                  # "123" + "Main St" -> "123 Main St"; blanks skipped
        out["F"] = [" ".join(str(m[f]).strip() for f in street if pd.notna(m.get(f)) and str(m[f]).strip()) or None
                    if m else None for m in match]
    return out.reset_index(drop=True), ftypes, sorted(set(unavailable), key=lambda c: (len(c), c))


def _mdy(d) -> str:
    return f"{d.month}/{d.day}/{d.year}"                 # her file's text dates: 9/14/2026


def _xl(v, sftype: str | None):
    """A Salesforce value the way Melanie's file holds it: dates as m/d/yyyy text, percents as fractions."""
    try:
        if v is None or pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        v = v.strip()
        if not v:
            return None
    if sftype in ("date", "datetime"):
        d = pd.to_datetime(v, errors="coerce")
        return _mdy(d) if pd.notna(d) else v
    if sftype == "percent":
        return float(v) / 100
    if sftype == "boolean":
        return "Y" if v in (True, "true", "True") else "N"
    if isinstance(v, np.generic):
        return v.item()
    return _xl_text(v) if isinstance(v, str) else v


def _xl_text(s: str) -> str:
    """Strip control characters openpyxl rejects (pasted soft line breaks) and cap at Excel's cell limit."""
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    return ILLEGAL_CHARACTERS_RE.sub("", s)[:32000]


def _put(ws, r: int, c: int, v) -> None:
    """Write a value; Salesforce text that starts with '=' stays text instead of becoming a live formula."""
    cell = ws.cell(r, c)
    cell.value = v
    if isinstance(v, str) and v.startswith("="):
        cell.data_type = "s"


def fill_melanie_template(template: str, rows: pd.DataFrame, ftypes: dict, submitted_by: str,
                          funding_date: date) -> bytes:
    """Open the committed template, fill the title block, and write one row per draw in its own column order."""
    import openpyxl
    from openpyxl.utils import column_index_from_string as ci
    wb = openpyxl.load_workbook(template)
    ws = wb[TEMPLATE_SHEET]
    _put(ws, 3, 5, _xl_text(submitted_by))                                    # E3
    ws["E4"] = _mdy(pd.Timestamp.now(tz="America/New_York"))                 # submitted today (ET, not server UTC)
    ws["E5"] = _mdy(funding_date)
    letters = list(MELANIE_MAP) + list(MELANIE_PROPERTY_MAP) + ["F"]
    for r, (_, row) in enumerate(rows.iterrows(), TEMPLATE_HEADER_ROW + 1):
        for col in letters:
            v = _xl(row.get(col), ftypes.get(col))
            if v is not None:
                _put(ws, r, ci(col), v)
        if bool(row.get("_multi_property", False)):
            ws.cell(r, ci("A")).value = "Multi-property deal: property columns left blank, see Salesforce"
        for col, pattern in MELANIE_FORMULAS.items():     # every input filled and non-zero (no #DIV/0!)
            if all(ws.cell(r, ci(c)).value not in (None, "", 0) for c in re.findall(r"([A-Z]+)\{r\}", pattern)):
                ws.cell(r, ci(col)).value = pattern.format(r=r)
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


# ───────────────────────────── Pipeline (macro) ─────────────────────────────
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
    c[2].caption(f"Draws wired since {start:%b} {start.day}, {start:%Y}, grouped by {gran.lower()}.")

    done = add_intervals(run_soql(inst, tok,               # <= today: a scheduled / typo'd future wire isn't funded
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND {WIRE_FIELD}>={start:%Y-%m-%d} "
        f"AND {WIRE_FIELD}<={date.today():%Y-%m-%d}"))
    flight = run_soql(inst, tok, f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND {OPEN_WHERE}")
    if done.empty:
        st.info("No completed draws in this range."); return

    done["Period"] = _period_key(pd.to_datetime(done[WIRE_FIELD]), gi)
    AMT = NET_FIELD if NET_FIELD in done else None
    periods = sorted(done["Period"].unique())
    cur_key = _period_key(pd.Series([pd.Timestamp(date.today())]), gi).iloc[0]

    # pick the featured period FIRST (most recent complete one), so we only correct what we show
    complete = [p for p in periods if p < cur_key]
    default_feat = complete[-1] if complete else periods[-1]
    rev = list(reversed(periods))
    feat = st.selectbox("Summary for", rev, index=rev.index(default_feat),
                        format_func=lambda p: p + ("  (in progress)" if p == cur_key else ""))

    # ── Days-to-fund comes from Land Gorilla, corrected for the SELECTED period only ──
    # Salesforce collapses the request date to the wire date, so its Days_to_Fund reads ~0. LG's draw
    # createdDate is the real request. We correct only the chosen period's draws (~hundreds, not the whole
    # multi-year span of thousands) so it loads in seconds, fetched in parallel and cached 24h.
    cfg = lg_config()
    fg = done[done["Period"] == feat].copy()
    if cfg:
        with st.spinner(f"Loading true request dates from Land Gorilla for {feat}…"):
            fg = apply_lg_correction(fg, cfg, same_day)
    fd2f = fg["_days"].dropna()

    # rollup — volume & dollars across all periods (cheap, SF only). Days-to-fund is NOT shown per-period
    # here because it would require correcting every period; the true number lives in the KPI + chart below.
    rows = []
    for p in periods:
        g = done[done["Period"] == p]
        funded = float(pd.to_numeric(g.get(AMT), errors="coerce").sum()) if AMT else None
        n = len(g)
        rows.append({
            "_raw": p,
            "Period": p + ("  (in progress)" if p == cur_key else ""),
            "Draws funded": n,
            "Loans": (g["Loan_Number__c"].nunique() if "Loan_Number__c" in g else None),
            "Funded ($)": funded,
            "Avg draw ($)": (funded / n if (funded is not None and n) else None),
        })
    roll = pd.DataFrame(rows)

    wait = fg["_wait_bd"].dropna()   # request -> package (borrower)
    fund = fg["_fund_bd"].dropna()   # package -> wire (us)
    same = (fund == same_day).mean() * 100 if len(fund) else 0.0   # same-day reads as 0 or 1 per the sidebar

    k = st.columns(4)
    k[0].metric(f"Draws funded · {feat}", f"{len(fg):,}")
    k[1].metric("Amount funded", money(pd.to_numeric(fg.get(AMT), errors="coerce").sum()) if AMT else "—")
    k[2].metric("Borrower wait", f"{wait.mean():.0f} bd" if len(wait) else "—",
                help="Business days from the borrower's draw request to a complete package "
                     "(Land Gorilla createdDate → approvedDate). The borrower / inspection / title side.")
    k[3].metric("Our funding", f"{fund.mean():.1f} bd" if len(fund) else "—",
                f"{same:.0f}% same-day" if len(fund) else None,
                help="Business days from a complete package to the wire (Land Gorilla approvedDate → wire). "
                     "This is CoreVest's own funding speed once the package is in.")

    holds = 0
    if not flight.empty and "Status__c" in flight:
        holds = int(flight["Status__c"].astype(str).str.contains("Hold|Pending Borrower|Revision", case=False, na=False).sum())
    k2 = st.columns(2)
    k2[0].metric("Open draws now", f"{len(flight):,}", f"{holds} need attention" if holds else None)
    if len(fd2f):
        k2[1].metric("Total: request → wire", f"{fd2f.mean():.0f} bd",
                     f"{fg['_days'].notna().mean()*100:.0f}% measured",
                     help="Business days from the Land Gorilla request date to the wire. '% measured' = share of "
                          "this period's draws Land Gorilla could date; the rest are left out, not guessed.")

    if len(wait) or len(fund):
        parts = []
        if len(wait):
            parts.append(f"borrowers take an average of **{wait.mean():.0f} business days** to get a complete package in")
        if len(fund):
            parts.append(f"CoreVest then funds in **{fund.mean():.1f} business day(s)**"
                         + (f" (**{same:.0f}% same-day**)" if same >= 50 else ""))
        blame = len(wait) and len(fund) and wait.mean() > fund.mean()
        st.caption(f"In **{feat}**, " + ", and ".join(parts) +
                   (" — the wait is the borrower/inspection/title side, not our funding." if blame else "."))
    if not cfg:
        st.error("Land Gorilla isn't configured — the request and package dates can't be corrected (Salesforce "
                 "collapses them to the wire date). Add [landgorilla] secrets.")

    st.caption(f"Across this range: **{int(roll['Draws funded'].sum()):,}** draws funded over {len(roll)} {gran.lower()}s. "
               "The newest period is marked *in progress* because it isn't finished yet — that's why its count is small.")

    # ---- rollup table (volume & dollars) ----
    st.markdown(f"**By {gran.lower()}** — volume & dollars")
    disp = roll.drop(columns="_raw").copy()
    for c in ["Funded ($)", "Avg draw ($)"]:
        disp[c] = disp[c].map(lambda x: money(x) if pd.notna(x) else "")
    disp["Loans"] = pd.to_numeric(disp["Loans"], errors="coerce")
    st.dataframe(disp.iloc[::-1], hide_index=True, width="stretch")

    # ---- charts ----
    _render_charts(done, roll, gran, fg, feat)

    # ---- featured-period detail + Excel ----
    st.markdown(f"**Draws funded in {feat}**")
    d = fg.copy()
    # All timing dates from Land Gorilla (Salesforce's are collapsed onto the wire); wire from Salesforce.
    d["Requested"], d["Package complete"] = d["_lg_req"], d["_lg_pkg"]
    d["Borrower wait (bd)"], d["Our funding (bd)"], d["Total (bd)"] = d["_wait_bd"], d["_fund_bd"], d["_days"]
    d["Source"] = np.where(d["_days"].notna(), "Land Gorilla", "no LG dates")
    ren = {"Loan_Number__c": "Loan #", "Loan_Advance_Number__c": "Draw #", "Deal__r.Name": "Deal",
           "Deal__r.Account.Name": "Account", "Borrower_Name_Text__c": "Borrower", "Status__c": "Status",
           WIRE_FIELD: "Wired", NET_FIELD: "Funded ($)", NOTES_FIELD: "Notes"}
    cols = [c for c in ["Loan_Number__c", "Loan_Advance_Number__c", "Deal__r.Name", "Deal__r.Account.Name",
                        "Borrower_Name_Text__c", "Status__c", "Requested", "Package complete", WIRE_FIELD,
                        "Borrower wait (bd)", "Our funding (bd)", "Total (bd)", "Source",
                        NET_FIELD, NOTES_FIELD] if c in d.columns]
    detail = d[cols].rename(columns=ren).sort_values("Wired", ascending=False)
    st.dataframe(detail, width="stretch", height=340)

    # Build the Excel only when the user asks for it (keeps every page render fast).
    if st.button("⬇️ Build turn-time Excel", key=f"xls_{feat}"):
        with st.spinner("Building workbook…"):
            xlsx = build_excel({f"By {gran.lower()}": roll.drop(columns="_raw"), f"Draws {feat}": detail})
        st.download_button("Download turn-time Excel", xlsx, file_name=f"construction_draws_{feat}.xlsx",
                           mime=XLSX_MIME, key=f"dl_{feat}")

    st.caption("All timing dates come from **Land Gorilla** (request = createdDate, package complete = approvedDate) — "
               "Salesforce collapses these to the wire date, so we never use its dates for timing. The wire date, "
               "notes, dollars and names come from Salesforce. **Borrower wait** = request→package; **Our funding** = "
               "package→wire; **Total** = request→wire.")


def _render_charts(done: pd.DataFrame, roll: pd.DataFrame, gran: str, fg: pd.DataFrame, feat: str):
    """Charts: days-to-fund distribution for the corrected featured period, plus volume/dollars per period
       (which don't need Land Gorilla, so they're instant across the whole range)."""
    d2f = fg["_days"].dropna()               # corrected, featured period only
    if alt is None:
        st.bar_chart(roll.set_index("_raw")["Draws funded"])
        st.bar_chart(roll.set_index("_raw")["Funded ($)"])
        return

    r = roll.drop(columns=[c for c in ["Period"] if c in roll.columns]).rename(columns={"_raw": "Period"})
    r = r.loc[:, ~r.columns.duplicated()]
    order = list(r["Period"])
    c1, c2 = st.columns(2)

    # 1) days-to-fund distribution for the featured period (the true numbers)
    with c1:
        st.markdown(f"**How long draws took in {feat}** (request → wire, business days)")
        if len(d2f):
            capped = d2f.clip(upper=45)
            hist = alt.Chart(pd.DataFrame({"days": capped})).mark_bar(color="#4C78A8").encode(
                x=alt.X("days:Q", bin=alt.Bin(maxbins=30), title="Business days (capped at 45)"),
                y=alt.Y("count()", title="Draws"))
            rule = alt.Chart(pd.DataFrame({"m": [float(d2f.mean())]})).mark_rule(color="#E45756", size=2).encode(x="m:Q")
            st.altair_chart(hist + rule, width="stretch")
            w, f = fg["_wait_bd"].mean(), fg["_fund_bd"].mean()
            side = (" Most of that time is the borrower side (request → complete package)."
                    if pd.notna(w) and pd.notna(f) and w > f else "")
            st.caption(f"Red line = average ({d2f.mean():.0f} business days).{side}")
        else:
            st.info("No days-to-fund data for this period.")

    # 2) funded dollars per period (no LG needed)
    with c2:
        st.markdown("**Amount funded per period**")
        bars = alt.Chart(r[["Period", "Funded ($)"]]).mark_bar(color="#54A24B").encode(
            x=alt.X("Period:N", sort=order, title=""),
            y=alt.Y("Funded ($):Q", title="Funded ($)"),
            tooltip=["Period", alt.Tooltip("Funded ($):Q", format="$,.0f")])
        st.altair_chart(bars, width="stretch")

    # 3) draws funded per period (no LG needed)
    st.markdown("**Draws funded per period**")
    bars = alt.Chart(r[["Period", "Draws funded"]]).mark_bar(color="#B9D7A8").encode(
        x=alt.X("Period:N", sort=order, title=""),
        y=alt.Y("Draws funded:Q", title="Draws"),
        tooltip=["Period", "Draws funded"])
    st.altair_chart(bars, width="stretch")


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
        f"ORDER BY Loan_Number__c, {REQ_FIELD} LIMIT {SEARCH_LIMIT + 1}"))
    if df.empty:
        st.warning("No matching draws."); return
    if len(df) > SEARCH_LIMIT:
        df = df.head(SEARCH_LIMIT)
        st.warning(f"Showing the first {SEARCH_LIMIT} matching draws — narrow the search to see the rest.")
    nloans = df["Loan_Number__c"].nunique() if "Loan_Number__c" in df else len(df)
    st.caption(f"Found **{len(df)}** draw(s) across **{nloans}** loan(s).")

    cfg = lg_config()
    # Timing comes from Land Gorilla (Salesforce collapses the request/package dates onto the wire date).
    df = apply_lg_correction(df, cfg, same_day)
    if cfg:
        st.caption(f"↳ timing from Land Gorilla for **{int(df['_days'].notna().sum())}** of {len(df)} draw(s) "
                   "— open draws have no wire date yet.")
    tmpl = (cfg or {}).get("template_id", LG_TEMPLATE_ID_DEFAULT)
    def first(colname):
        return g[colname].dropna().iloc[0] if colname in g and g[colname].notna().any() else ""

    for loan_no, g in df.groupby("Loan_Number__c", dropna=False):
        deal = first("Deal__r.Name")
        account = first("Deal__r.Account.Name")
        borrower = first("Borrower_Name_Text__c") or first("Borrower_Name__c")
        st.markdown(f"### Loan {loan_no} — {deal}")
        sub = []
        if account: sub.append(f"Account: {account}")
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
    draw_no = next((v for v in (row.get("Loan_Advance_Number__c"), row.get("Name")) if pd.notna(v) and v != ""),
                   "draw")
    top = st.columns([2, 1, 1, 1, 1])
    top[0].markdown(f"**Draw {draw_no}** — {row.get('Status__c','')}")
    if NET_FIELD in row and pd.notna(row.get(NET_FIELD)):
        top[1].metric("Funded", money(row[NET_FIELD]))
    if pd.notna(row.get("_days")):
        top[2].metric("Request → wire", f"{row['_days']:.0f} bd",
                      help="Business days from the Land Gorilla request date (createdDate) to the Salesforce wire.")
    if pd.notna(row.get("_wait_bd")):
        top[3].metric("Borrower wait", f"{row['_wait_bd']:.0f} bd",
                      help="Request → complete package (Land Gorilla approvedDate): borrower / inspection / title.")
    if pd.notna(row.get("_fund_bd")):
        top[4].metric("Our funding", f"{row['_fund_bd']:.0f} bd", help="Complete package → wire: CoreVest's own speed.")

    steps = [{"Milestone": lbl, "Date": (pd.to_datetime(row.get(f)).date() if pd.notna(row.get(f)) else None),
              "": "✅" if pd.notna(row.get(f)) else "⬜"} for lbl, f in MILESTONES]
    tdf = pd.DataFrame(steps)
    recorded = tdf["Date"].notna().sum()
    st.progress(recorded / len(MILESTONES), text=f"Salesforce milestones recorded: {recorded}/{len(MILESTONES)}")
    st.dataframe(tdf, hide_index=True, width="stretch",
                 column_config={"": st.column_config.TextColumn(width="small")})
    if pd.notna(row.get("_lg_req")):
        pkg = row.get("_lg_pkg")
        st.caption(f"Land Gorilla: requested {row['_lg_req']} · package complete {pkg if pd.notna(pkg) else '—'}. "
                   "Salesforce's Requested / package dates are often stamped on the wire date, so timing uses these.")

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
    #   name, type, status, createdDate, submittedDate, approvedDate, effectiveDate, total amount.
    #   effectiveDate collapses onto approvedDate — it is NOT the wire (that's Salesforce Wire_Date__c).
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
                              "Created": None, "Submitted": None, "Approved": None, "Effective": None, "Amount": None})
            continue
        draw_rows.append({
            "Draw": sf_draw,
            "LG draw": det.get("name") or "—",
            "Type": det.get("type") or "—",
            "Status": det.get("status") or "—",
            "Created": det.get("created"),
            "Submitted": det.get("submitted"),
            "Approved": det.get("approved"),
            "Effective": det.get("funded"),
            "Amount": det.get("amount"),
        })
        for p in (det.get("payees") or []):
            all_payees.append({"Draw": det.get("name") or sf_draw, "Payee": p["payee"], "Amount": p.get("amount")})

    if draw_rows:
        st.markdown("**Land Gorilla draws**")
        ddf = pd.DataFrame(draw_rows)
        ddf["Amount"] = ddf["Amount"].map(lambda x: money(x) if pd.notna(x) else "—")
        st.dataframe(
            ddf, hide_index=True, width="stretch",
            column_config={
                "Created": st.column_config.DateColumn("Created", format="MM/DD/YYYY"),
                "Submitted": st.column_config.DateColumn("Submitted", format="MM/DD/YYYY"),
                "Approved": st.column_config.DateColumn("Approved", format="MM/DD/YYYY"),
                "Effective": st.column_config.DateColumn("Effective", format="MM/DD/YYYY",
                                                         help="LG effectiveDate (= approval) — not the wire date."),
            })
        if all_payees:
            with st.expander(f"Payees across these draws ({len(all_payees)})"):
                pdf = pd.DataFrame(all_payees)
                pdf["Amount"] = pdf["Amount"].map(lambda x: money(x) if pd.notna(x) else "—")
                st.dataframe(pdf, hide_index=True, width="stretch")
    else:
        st.caption("No Land Gorilla draw containers on these advances yet.")


# ───────────────────────────── Funding request (weekly file) ────────────────
@st.cache_data(show_spinner=False)
def template_headers(path: str) -> dict:
    """{column letter: header} from the template's header row (for the 'left blank' note)."""
    import openpyxl
    from openpyxl.utils import get_column_letter
    ws = openpyxl.load_workbook(path, read_only=True)[TEMPLATE_SHEET]
    row = next(ws.iter_rows(min_row=TEMPLATE_HEADER_ROW, max_row=TEMPLATE_HEADER_ROW))
    return {get_column_letter(c.column): " ".join(str(c.value).split()) for c in row if c.value}


def render_funding_request(inst, tok, rt):
    st.subheader("Funding-request file — Melanie's format")
    tpl = _find_template()
    if not tpl:
        st.error("Template not found — commit `Advance_template.xlsx` next to draw_dashboard.py."); return
    st.caption("The weekly *Construction/Renovation Funding Request* workbook, filled from Salesforce — one row per "
               "draw, like the Generate File button on the Construction Advance IC Approvals page.")
    c = st.columns([1.2, 1, 1.6])
    who = c[0].text_input("Submitted by", key="fr_by", placeholder="Your name")
    fdate = c[1].date_input("Scheduled funding date", value=date.today(), key="fr_date")
    mode = c[2].radio("Draws", ["Open (not yet wired)", "Wired in a date range"], key="fr_mode", horizontal=True)
    if mode.startswith("Open"):
        where = OPEN_WHERE
    else:
        rng = st.date_input("Wire dates", value=(date.today() - timedelta(days=7), date.today()), key="fr_rng")
        if len(rng) != 2:
            st.info("Pick a start and an end date."); return
        where = f"{WIRE_FIELD}>={rng[0]:%Y-%m-%d} AND {WIRE_FIELD}<={rng[1]:%Y-%m-%d}"
    with st.spinner("Loading draws from Salesforce…"):
        rows, ftypes, unavailable = funding_request_rows(inst, tok, rt, where)
    if rows.empty:
        st.info("No draws in that range."); return

    f = st.columns(2)
    people = sorted({p for col in PERSON_FIELDS if col in rows for p in rows[col].dropna()})
    if people:
        person = f[0].selectbox("Coordinator / analyst", ["Everyone"] + people,
                                help="Salesforce's Generate File lists only the logged-in user's draws — "
                                     "pick that person to match it.")
        if person != "Everyone":
            rows = rows[np.logical_or.reduce([rows[col].eq(person) for col in PERSON_FIELDS if col in rows])]
    if "Status__c" in rows:
        statuses = sorted(rows["Status__c"].dropna().unique())
        default = [s for s in statuses if s in FUNDING_STAGES] or statuses
        rows = rows[rows["Status__c"].isin(f[1].multiselect(
            "Statuses", statuses, default=default,
            help="Starts with the approval / release stages a funding request covers; add others as needed."))]
    total = pd.to_numeric(rows["U"], errors="coerce").sum() if "U" in rows else None
    st.caption(f"**{len(rows)}** draw(s) on **{rows['B'].nunique() if 'B' in rows else 0}** loan(s)"
               + (f" · Current Draw Amount total **{money(total)}**" if total is not None else ""))

    if st.button("Build funding-request file", type="primary", disabled=rows.empty or not who.strip()):
        with st.spinner("Filling the template…"):
            xlsx = fill_melanie_template(tpl, rows, ftypes, who.strip(), fdate)
        st.download_button("⬇️ Download funding-request file", xlsx, mime=XLSX_MIME,
                           file_name=f"Construction-Renovation Funding Request "
                                     f"{fdate.month}_{fdate.day}_{fdate.year}.xlsx")
    if not who.strip():
        st.caption("Enter your name under *Submitted by* to build the file.")
    heads = template_headers(tpl)
    filled = (set(MELANIE_MAP) | set(MELANIE_PROPERTY_MAP) | {"F"}) - set(unavailable)
    blank = [f"{k} {v}" for k, v in heads.items() if k not in filled and k not in MELANIE_FORMULAS]
    st.caption(f"Current Draw Amount (column U, which feeds the total) is Salesforce `{MELANIE_MAP.get('U')}`. "
               f"LTV / equity columns are live formulas where their inputs are filled. Not filled here: "
               f"{', '.join(blank) or 'none'}. Salesforce's own Generate File remains the reference for those.")


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
    if sf is None:                                      # no (working) creds -> OAuth redirect fallback
        try:
            cfg = load_sf_oauth()
        except Exception as exc:
            err = err or str(exc)
        if cfg:
            try:
                finish_oauth(cfg)
                sf = sf_from_session(); mode = "oauth"
            except Exception as exc:
                err = str(exc)
    if sf is not None:
        err = None

    with st.sidebar:
        st.header("Salesforce")
        if err:
            st.error(err)
        elif sf is None:
            st.info("Not connected")
        else:
            st.success(f"Connected · {mode}")
            st.caption(st.session_state.get("salesforce_auth", {}).get("instance_url", ""))
            if mode == "oauth" and st.button("Log out", width="stretch"):
                clear_sf_session(); st.rerun()
        st.divider()
        same_day = 1 if st.radio("Same-day convention", ["0 business days", "1 business day"],
                                 help="A package that's complete and wired the same day counts as this. Applies to "
                                      "Our funding and the request → wire total. Confirm with Melanie — it shifts "
                                      "the averages.").startswith("1") else 0
        st.divider()
        st.caption("Land Gorilla: " + ("configured ✅" if lg_config() else "not set"))

    if sf is None and not cfg:                           # nothing to log in with
        st.error(err or "Salesforce isn't configured.")
        st.caption("For the no-callback path, put username / password / security_token under "
                   "[salesforce] in secrets. For OAuth, provide client_id / client_secret / redirect_uri / auth_host.")
        st.stop()
    if sf is None:                                       # OAuth path: not logged in yet, or the last try failed
        if err:
            st.error(err)
        st.subheader("Step 1 — Log in to Salesforce")
        st.info("Log in to load the construction draw book.")
        st.link_button("Log in to Salesforce", login_url(cfg))
        st.caption(f"Callback URL: {cfg['redirect_uri']}")
        st.stop()

    inst = st.session_state["salesforce_auth"]["instance_url"]
    tok = st.session_state["salesforce_auth"]["access_token"]
    page = st.sidebar.radio("View", ["Pipeline (macro)", "Loan detail (micro)", "Funding request"])
    try:                          # the record-type / describe lookups can hit an expired session too
        rt = construction_rt(inst, tok)
        if not rt:
            st.error("Could not find the 'Construction Advance' record type."); st.stop()
        sel = select_fields(inst, tok)
        if page.startswith("Pipeline"):
            render_pipeline(inst, tok, rt, sel, same_day)
        elif page.startswith("Loan"):
            render_loan_detail(inst, tok, rt, sel, same_day)
        else:
            render_funding_request(inst, tok, rt)
        st.session_state.pop("_sf_retried", None)
    except Exception as exc:
        msg = str(exc)
        if "INVALID_SESSION_ID" in msg or "Session expired" in msg:
            try:
                sf_login_credentials.clear()          # force a fresh login on the credentials path
            except Exception:
                pass
            clear_sf_session()
            if not st.session_state.get("_sf_retried"):  # reload once; a second failure stops with a message
                st.session_state["_sf_retried"] = True
                st.rerun()
            st.warning("Your Salesforce session expired — refresh the page to log in again."); st.stop()
        raise


if __name__ == "__main__":
    main()
