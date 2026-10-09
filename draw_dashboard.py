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
         NET_FIELD, GROSS_FIELD, "LastModifiedBy.Name", "LastModifiedDate"]
        + [f for f, _ in CONSTRUCTION_FIELDS] + [m[1] for m in MILESTONES])
TERMINAL = ["Completed", "Cancelled", "Rescinded", "Rejected by Capital Partner"]
OPEN_WHERE = f"{WIRE_FIELD}=null AND Status__c NOT IN ({','.join(repr(s) for s in TERMINAL)})"
# Status__c stages a weekly funding request covers (approval through release) — the export's default filter.
FUNDING_STAGES = ("Pending Approval", "Approved", "Pending Capital Partner Release", "Pending Release")
SEARCH_LIMIT = 200    # loan-detail search: most draws rendered (and dated in Land Gorilla) per query
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

_HOLS = USFederalHolidayCalendar().holidays("2018-01-01", "2032-12-31").values.astype("datetime64[D]")


def _today() -> date:
    """Today in the ops time zone — Streamlit Cloud's clock is UTC, a day ahead every US evening."""
    return pd.Timestamp.now(tz="America/New_York").date()


def _secrets(name: str) -> dict:
    """One [section] of Streamlit secrets, or {} when there's no secrets.toml (e.g. a fresh local checkout)."""
    try:
        return dict(st.secrets.get(name, {}))
    except Exception:
        return {}


def _txt(v) -> str:
    """Display text for a Salesforce value: blank for None / NaN (pandas 3 strings use NaN, which is truthy)."""
    return "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v).strip()


# ───────────────────────────── Salesforce OAuth ─────────────────────────────
def install_truststore() -> None:
    try:
        import truststore
        truststore.inject_into_ssl()
    except Exception:
        pass


def load_sf_oauth() -> dict[str, str]:
    sec = _secrets("salesforce")
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
    sec = _secrets("salesforce")
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
def describe_full(inst: str, tok: str, obj: str) -> dict:
    """One sObject's describe, trimmed: {"label", "fields": {field: {"type", "ref", "rel", "label"}},
       "children": [(child object, its lookup field)]}."""
    sf = Salesforce(instance_url=inst, session_id=tok)
    d = getattr(sf, obj).describe()
    return {"label": d.get("label") or obj,
            "fields": {f["name"]: {"type": f.get("type"), "ref": list(f.get("referenceTo") or []),
                                   "rel": f.get("relationshipName"), "label": f.get("label") or f["name"]}
                       for f in d["fields"]},
            "children": [(c.get("childSObject"), c.get("field")) for c in d.get("childRelationships") or []]}


def describe_object(inst: str, tok: str, obj: str) -> dict:
    """{field: {"type", "ref" (objects it points to), "rel" (relationship name), "label"}} for one sObject."""
    return describe_full(inst, tok, obj)["fields"]


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
        if (seg + "Id") in present:                                 # standard lookup, e.g. LastModifiedBy.Name
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
    for col in ("_lg_req", "_lg_pkg", "_pkg", "_check"):
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
    """Each draw's timeline, in business days:
         request          = Land Gorilla createdDate (Salesforce's request date is collapsed onto the wire)
         package received = Salesforce "Date Full Draw Package Received" when recorded, else LG approvedDate
         wire             = Salesforce Wire_Date__c
       Turn-time (package -> wire) is the official measure and needs no Land Gorilla link. Out-of-order dates
       are left out of the averages and noted in _check instead of being guessed."""
    if df.empty or WIRE_FIELD not in df:
        return df
    df = df.copy()
    lg = {}
    if cfg and CONTAINER_FIELD in df:
        live = df.loc[df[CONTAINER_FIELD].notna() & df[WIRE_FIELD].notna(), CONTAINER_FIELD]
        try:
            lg = lg_draw_dates(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), live.tolist())
        except Exception as exc:      # no exception text: a malformed secret would be echoed to every viewer
            st.warning(f"Land Gorilla is unavailable right now ({type(exc).__name__}) — request dates are blank.")
    cont = df[CONTAINER_FIELD] if CONTAINER_FIELD in df else pd.Series(None, index=df.index)
    got = cont.map(lambda c: lg.get(str(c)) if pd.notna(c) else None)
    created = got.map(lambda t: t[0] if t else None)
    approved = got.map(lambda t: t[1] if t else None)
    wire = df[WIRE_FIELD]
    sf_pkg = pd.to_datetime(df[PKG_FIELD], errors="coerce") if PKG_FIELD in df else pd.Series(pd.NaT, index=df.index)
    pkg = sf_pkg.dt.strftime("%Y-%m-%d").where(sf_pkg.notna(), approved)
    c, p, w = (pd.to_datetime(s, errors="coerce").dt.normalize() for s in (created, pkg, wire))

    df["_lg_req"], df["_lg_pkg"], df["_pkg"] = created, approved, pkg
    df["_days"] = bdays(created, wire, same_day)                       # request -> wire
    df["_days_src"] = np.where(df["_days"].notna(), "Land Gorilla", "")
    df["_fund_bd"] = bdays(pkg, wire, same_day).where(p <= w)          # turn-time: package received -> wire
    df["_wait_bd"] = bdays(created, pkg, 0).where((c <= p) & (p <= w))  # request -> package (borrower side)
    check = pd.Series(None, index=df.index, dtype=object)
    check[p.notna() & w.notna() & (p > w)] = "Package date is after the wire date"
    check[c.notna() & p.notna() & (c > p)] = "Package date is before the draw was requested"
    check[w.notna() & p.isna()] = "No package-received date"
    df["_check"] = check
    return df


def money(x) -> str:
    try:
        v = float(x)
    except Exception:
        return "—"
    return f"${v:,.0f}"


# ───────────────────────────── Land Gorilla (draw detail only) ──────────────
def lg_config() -> dict | None:
    sec = _secrets("landgorilla")
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
    """Formatted multi-sheet .xlsx (Arial, header band, freeze panes, autofilter) as bytes for download.
       Dates are real Excel dates; "($)" columns are whole dollars and "(bd)" columns whole days."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    HFONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    HFILL = PatternFill("solid", fgColor="1F3864")
    AR = Font(name="Arial", size=10)
    wb = Workbook(); wb.remove(wb.active)
    for name, df in sheets.items():
        ws = wb.create_sheet(str(name)[:31])
        fmts = {j: ('"$"#,##0' if str(col).endswith("($)") else "0" if str(col).endswith("(bd)") else None)
                for j, col in enumerate(df.columns, 1)}
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
                fmt = fmts[j]
                if isna:
                    val = None
                elif isinstance(v, date):                    # datetime / Timestamp are date subclasses
                    val, fmt = (v.date() if isinstance(v, datetime) else v), "m/d/yyyy"
                elif isinstance(v, np.generic):
                    val = v.item()
                else:
                    val = _xl_text(v) if isinstance(v, str) else v
                _put(ws, i, j, val)
                ws.cell(i, j).font = AR
                if fmt and val is not None:
                    ws.cell(i, j).number_format = fmt
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
    extra = [p for p in ("Id", "Deal__c", "Status__c", "Loan_Advance_Number__c", WIRE_FIELD) + PERSON_FIELDS
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


@st.cache_data(ttl=3600, show_spinner=False)
def funded_by_quarter(inst: str, tok: str, rt: str, since: str, until: str) -> pd.DataFrame:
    """Draws and dollars wired per calendar quarter (columns y, q, n, amt) — feeds the Total Summary sheet."""
    base = f"FROM Advance__c WHERE RecordTypeId='{rt}' AND {WIRE_FIELD}>={since} AND {WIRE_FIELD}<={until}"
    try:                                                  # one aggregate query when the org allows it
        agg = run_soql(inst, tok, f"SELECT CALENDAR_YEAR({WIRE_FIELD}) y, CALENDAR_QUARTER({WIRE_FIELD}) q, "
                                  f"COUNT(Id) n, SUM({NET_FIELD}) amt {base} "
                                  f"GROUP BY CALENDAR_YEAR({WIRE_FIELD}), CALENDAR_QUARTER({WIRE_FIELD})")
        if {"y", "q", "n", "amt"} <= set(agg.columns):
            return agg[["y", "q", "n", "amt"]].astype({"y": int, "q": int, "n": int})
    except Exception:
        pass
    raw = run_soql(inst, tok, f"SELECT {WIRE_FIELD},{NET_FIELD} {base}")
    if raw.empty:
        return pd.DataFrame({"y": [], "q": [], "n": [], "amt": []})
    w = pd.to_datetime(raw[WIRE_FIELD], errors="coerce")
    g = pd.DataFrame({"y": w.dt.year, "q": w.dt.quarter, "amt": pd.to_numeric(raw[NET_FIELD], errors="coerce")})
    return g.dropna(subset=["y"]).groupby(["y", "q"]).agg(n=("amt", "size"), amt=("amt", "sum")).reset_index()


def file_summary(inst: str, tok: str, rt: str, rows: pd.DataFrame) -> dict:
    """Figures for the workbook's Total Summary sheet: quarterly totals since 2021 and this month to date."""
    today = _today()
    quarters = funded_by_quarter(inst, tok, rt, "2021-01-01", f"{today:%Y-%m-%d}")
    month = funded_by_quarter(inst, tok, rt, f"{today.replace(day=1):%Y-%m-%d}", f"{today:%Y-%m-%d}")
    pending = 0.0                                         # draws in this file that aren't wired yet
    if "U" in rows and WIRE_FIELD in rows:
        pending = float(pd.to_numeric(rows.loc[rows[WIRE_FIELD].isna(), "U"], errors="coerce").sum())
    return {"quarters": quarters, "mtd": float(month["amt"].sum()), "pending": pending}


def _write_total_summary(wb, summary: dict) -> None:
    """Replace the template's hard-coded 2024 ledger with figures computed from Salesforce."""
    from openpyxl.styles import Font
    if "Total Summary" not in wb.sheetnames:
        return
    ws = wb["Total Summary"]
    money_fmt = ws["D6"].number_format or '"$"#,##0'
    for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 60), max_col=12):
        for c in row:
            c.value = None
    today = _today()
    now = (today.year, (today.month - 1) // 3 + 1)
    bold = Font(bold=True)
    ws["B2"] = "This Month"
    ws["B2"].font = bold
    ws["B3"] = summary["mtd"] + summary["pending"]
    ws["B3"].number_format = money_fmt
    ws["C3"] = "funded this month to date" + (" + this request's pending draws" if summary["pending"] else "")
    q = summary["quarters"]
    years = sorted({int(y) for y in q["y"]} | {today.year}, reverse=True)
    r = 5
    for y in years:
        ws.cell(r, 2, f"{y} TOTALS").font = bold
        ws.cell(r, 3, "Draws").font = bold
        ws.cell(r, 4, "Funded").font = bold
        for n in range(1, 5):
            ws.cell(r + n, 2, f"Q{n}" + (" (running)" if (y, n) == now else ""))
            if (y, n) <= now:
                m = q[(q["y"] == y) & (q["q"] == n)]
                ws.cell(r + n, 3, int(m["n"].sum()))
                ws.cell(r + n, 4, float(m["amt"].sum())).number_format = money_fmt
        ws.cell(r + 5, 2, "TOTAL").font = bold
        ws.cell(r + 5, 3, f"=SUM(C{r + 1}:C{r + 4})")
        ws.cell(r + 5, 4, f"=SUM(D{r + 1}:D{r + 4})").number_format = money_fmt
        r += 8


def fill_melanie_template(template: str, rows: pd.DataFrame, ftypes: dict, submitted_by: str,
                          funding_date: date, summary: dict | None = None) -> bytes:
    """Open the committed template, fill the title block, and write one row per draw in its own column order.
       The other sheets are brought up to date too: Total Summary from Salesforce, the Cap Partners title
       blocks from this request, and the template's stale hidden reference sheet removed."""
    import openpyxl
    from openpyxl.utils import column_index_from_string as ci
    wb = openpyxl.load_workbook(template)
    ws = wb[TEMPLATE_SHEET]
    _put(ws, 3, 5, _xl_text(submitted_by.strip()) or None)                   # E3 (optional)
    for name in ("Cap Partners Funded By CV", "Cap Partners"):              # their title blocks mirror E3:E5
        if name in wb.sheetnames:
            wb[name]["D4"] = f"='{TEMPLATE_SHEET}'!E4"
            wb[name]["D5"] = f"='{TEMPLATE_SHEET}'!E5"
    if "Sheet1" in wb.sheetnames:                                            # March-2024 servicer data, unused
        del wb["Sheet1"]
    if summary:
        _write_total_summary(wb, summary)
    ws["E4"] = _mdy(_today())                                               # submitted today (ET, not server UTC)
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


# ───────────────────────────── approvals ────────────────────────────────────
APPROVAL_COLS = ["Approval", "Approval status", "Approved by", "Approved on", "Approval comments"]


@st.cache_data(ttl=3600, show_spinner=False)
def _approval_source(inst: str, tok: str) -> dict | None:
    """The org's IC approval records for draws (the "B-mmddyyyy…" rows on the Construction Advance IC Approvals
       page): an object tied to Advance__c whose name mentions approval, plus its likeliest fields."""
    adv = describe_full(inst, tok, "Advance__c")
    src = next(({"kind": "parent", "obj": m["ref"][0], "rel": m["rel"]} for m in adv["fields"].values()
                if m["type"] == "reference" and m["ref"] and m["rel"] and "approv" in m["ref"][0].lower()), None)
    src = src or next(({"kind": "child", "obj": ch, "fk": fk} for ch, fk in adv["children"]
                       if ch and fk and "approv" in ch.lower()), None)
    if not src:
        return None
    fields = describe_full(inst, tok, src["obj"])["fields"]

    def pick(test):
        return next((n for n, m in fields.items() if test(n.lower(), str(m.get("label", "")).lower(), m)), None)

    by = pick(lambda n, l, m: m["type"] == "reference" and m["ref"] == ["User"] and "approv" in n + l)
    src.update(status=pick(lambda n, l, m: "status" in l and m["type"] in ("picklist", "string")),
               when=pick(lambda n, l, m: "approv" in l and m["type"] in ("date", "datetime")) or "LastModifiedDate",
               comments=pick(lambda n, l, m: "comment" in l and m["type"] in ("textarea", "string")),
               by=f"{fields[by]['rel']}.Name" if by and fields[by].get("rel") else "LastModifiedBy.Name")
    return src


@st.cache_data(ttl=300, show_spinner=False)
def approvals_for(inst: str, tok: str, ids: tuple) -> pd.DataFrame:
    """Who approved each draw, when, and what they wrote — indexed by Advance Id. Uses the org's IC approval
       records when they can be found, else Salesforce's approval history (approve / reject steps)."""
    out = {i: {} for i in ids if isinstance(i, str) and i}
    if not out:
        return pd.DataFrame(columns=APPROVAL_COLS)
    try:
        steps = _soql_in(inst, tok, "ProcessInstance.TargetObjectId,StepStatus,Comments,Actor.Name,CreatedDate",
                         "ProcessInstanceStep", "ProcessInstance.TargetObjectId", list(out))
    except Exception:
        steps = pd.DataFrame()
    if not steps.empty and "StepStatus" in steps:
        steps = steps[steps["StepStatus"].isin(["Approved", "Rejected"])].sort_values("CreatedDate")
        for aid, g in steps.groupby("ProcessInstance.TargetObjectId"):
            last = g.iloc[-1]
            out.setdefault(aid, {}).update({
                "Approval status": last["StepStatus"], "Approved by": _txt(last.get("Actor.Name")) or None,
                "Approved on": _d(last.get("CreatedDate")),
                "Approval comments": "\n".join(
                    f'{_mdy(_d(r["CreatedDate"]))}: {_txt(r.get("Actor.Name"))} : "{_txt(r.get("Comments"))}"'
                    for _, r in g.iterrows())})
    try:
        src = _approval_source(inst, tok)
        if src:
            cols = [c for c in ("Name", src["status"], src["when"], src["comments"], src["by"]) if c]
            if src["kind"] == "parent":
                pre, key = f"{src['rel']}.", "Id"
                recs = _soql_in(inst, tok, ",".join(["Id"] + [pre + c for c in cols]), "Advance__c", "Id", list(out))
            else:
                pre, key = "", src["fk"]
                recs = _soql_in(inst, tok, ",".join([key] + cols), src["obj"], key, list(out))
            if not recs.empty and pre + src["when"] in recs:
                recs = recs.sort_values(pre + src["when"])
            for _, r in recs.iterrows():                 # latest approval record per draw wins
                if not _txt(r.get(pre + "Name")):
                    continue
                found = {"Approval": _txt(r.get(pre + "Name")),
                         "Approval status": _txt(r.get(pre + src["status"])) if src["status"] else None,
                         "Approved by": _txt(r.get(pre + src["by"])),
                         "Approved on": _d(r.get(pre + src["when"])),
                         "Approval comments": _txt(r.get(pre + src["comments"])) if src["comments"] else None}
                out.setdefault(r[key], {}).update({k: v for k, v in found.items() if v})
    except Exception:
        pass
    return pd.DataFrame.from_dict(out, orient="index").reindex(columns=APPROVAL_COLS)


# ───────────────────────────── display helpers ──────────────────────────────
def _bd(x) -> str:
    """Business days for display: a decimal only where it says something (0.4 bd), whole numbers from 10 up."""
    if x is None or pd.isna(x):
        return "—"
    s = f"{x:.1f}" if abs(x) < 10 else f"{x:.0f}"
    return (s[:-2] if s.endswith(".0") else s) + " bd"


def _d(v):
    """A date-ish value as a date, in Eastern time for timestamps (None when blank)."""
    t = pd.to_datetime(v, errors="coerce")
    if pd.isna(t):
        return None
    return (t.tz_convert("America/New_York") if t.tzinfo is not None else t).date()


def _dates(s) -> pd.Series:
    """A column of Salesforce / Land Gorilla dates as dates."""
    return pd.to_datetime(s, errors="coerce").dt.normalize()


def _col(df: pd.DataFrame, name: str) -> pd.Series:
    return df[name] if name in df else pd.Series(None, index=df.index, dtype=object)


def _draw_table(d: pd.DataFrame, inst: str, ap: pd.DataFrame | None = None) -> pd.DataFrame:
    """The draws in plain, business-friendly columns (shown in the app and used for the Excel export)."""
    ids = _col(d, "Id")
    t = pd.DataFrame({
        "Loan #": _col(d, "Loan_Number__c"), "Draw #": _col(d, "Loan_Advance_Number__c"),
        "Deal": _col(d, "Deal__r.Name"), "Borrower": _col(d, "Borrower_Name_Text__c"), "Status": _col(d, "Status__c"),
        "Requested": _dates(_col(d, "_lg_req")), "Package received": _dates(_col(d, "_pkg")),
        "Wired": _dates(_col(d, WIRE_FIELD)),
        "Turn-time (bd)": pd.to_numeric(_col(d, "_fund_bd"), errors="coerce"),
        "Borrower side (bd)": pd.to_numeric(_col(d, "_wait_bd"), errors="coerce"),
        "Request → wire (bd)": pd.to_numeric(_col(d, "_days"), errors="coerce"),
        "Funded ($)": pd.to_numeric(_col(d, NET_FIELD), errors="coerce"),
    })
    for c in ("Approved by", "Approved on", "Approval comments"):
        t[c] = ids.map(ap[c]) if ap is not None and c in ap else None
    t["Approved on"] = _dates(t["Approved on"])
    t["Coordinator"] = _col(d, "Advance_Coordinator__r.Name")
    t["Last updated by"] = _col(d, "LastModifiedBy.Name")
    t["Notes"] = _col(d, NOTES_FIELD)
    t["Salesforce"] = ids.map(lambda i: f"{inst.rstrip('/')}/{i}" if isinstance(i, str) and i else None)
    return t


def _table_config() -> dict:
    C = st.column_config
    return {
        "Loan #": C.TextColumn(width="small"), "Draw #": C.TextColumn(width="small"),
        "Requested": C.DateColumn(format="M/D/YYYY", help="When the borrower requested the draw"),
        "Package received": C.DateColumn(format="M/D/YYYY", help="Full draw package received"),
        "Wired": C.DateColumn(format="M/D/YYYY"),
        "Turn-time (bd)": C.NumberColumn("Turn-time", format="%d bd",
                                         help="Business days from full draw package received to wire"),
        "Borrower side (bd)": C.NumberColumn("Borrower side", format="%d bd",
                                             help="Business days from the request to a full package"),
        "Request → wire (bd)": C.NumberColumn("Request → wire", format="%d bd"),
        "Funded ($)": C.NumberColumn("Funded", format="dollar", step=1),
        "Approved on": C.DateColumn(format="M/D/YYYY"),
        "Approval comments": C.TextColumn(width="large"),
        "Notes": C.TextColumn(width="large"),
        "Salesforce": C.LinkColumn(display_text="Open", width="small"),
    }


def _excel_table(fg: pd.DataFrame, table: pd.DataFrame) -> pd.DataFrame:
    """The turn-time data dump: every column in the table plus the Salesforce milestone dates in between."""
    x = table.drop(columns=["Salesforce"]).copy()
    pos = x.columns.get_loc("Package received")
    for i, (label, field) in enumerate(MILESTONES[1:-2]):    # inspection ordered … internal review complete
        x.insert(pos + i, label, _dates(_col(fg, field)))
    return x


# ───────────────────────────── Turn-time (overview) ─────────────────────────
def _period_key(w: pd.Series, gran: str) -> pd.Series:
    if gran == "Monthly":
        return w.dt.to_period("M").astype(str)
    if gran == "Quarterly":
        return w.dt.to_period("Q").astype(str)
    return w.dt.year.astype(str)


def _span_start(gran: str, n: int) -> date:
    today = _today()
    if gran == "Monthly":
        return (pd.Timestamp(today).to_period("M") - (n - 1)).start_time.date()
    if gran == "Quarterly":
        return (pd.Timestamp(today).to_period("Q") - (n - 1)).start_time.date()
    return date(today.year - (n - 1), 1, 1)


def render_pipeline(inst, tok, rt, sel, same_day):
    st.header("Draw turn-time")
    st.caption("How long construction draws take, and where the time goes. All figures are business days.")
    c = st.columns([2, 1.3, 1.4])
    gran = c[0].segmented_control("View by", ["Month", "Quarter", "Year", "Date range"], default="Quarter",
                                  key="ov_gran") or "Quarter"
    if gran == "Date range":
        rng = c[1].date_input("Wire dates", value=(_today() - timedelta(days=30), _today()), key="ov_rng")
        if len(rng) != 2:
            st.info("Pick a start and an end date."); return
        start, end, gi = rng[0], min(rng[1], _today()), "Monthly"
    else:
        gi = {"Month": "Monthly", "Quarter": "Quarterly", "Year": "Yearly"}[gran]
        opts = {"Month": [6, 12, 24], "Quarter": [4, 8, 12], "Year": [3, 5]}[gran]
        nper = c[1].selectbox("How many", opts, index=1 if gran != "Year" else 0, key=f"ov_n_{gran}")
        start, end = _span_start(gi, nper), _today()

    # The rollup only needs four fields; full detail is fetched for the period being summarised alone.
    span = ",".join(f for f in ("Id", "Loan_Number__c", WIRE_FIELD, NET_FIELD) if f in sel.split(","))
    with st.spinner("Loading draws…"):
        done = run_soql(inst, tok, f"SELECT {span} FROM Advance__c WHERE RecordTypeId='{rt}' "
                                   f"AND {WIRE_FIELD}>={start:%Y-%m-%d} AND {WIRE_FIELD}<={end:%Y-%m-%d}")
        flight = run_soql(inst, tok, f"SELECT Id,Status__c FROM Advance__c WHERE RecordTypeId='{rt}' AND {OPEN_WHERE}")
    if done.empty:
        st.info("No draws were wired in this range."); return
    done["Period"] = _period_key(pd.to_datetime(done[WIRE_FIELD]), gi)
    AMT = NET_FIELD if NET_FIELD in done else None
    periods = sorted(done["Period"].unique())
    cur_key = _period_key(pd.Series([pd.Timestamp(_today())]), gi).iloc[0]
    if gran == "Date range":
        feat, p_start, p_end = f"{_mdy(start)} – {_mdy(end)}", start, end
    else:
        complete = [p for p in periods if p < cur_key]
        default_feat = complete[-1] if complete else periods[-1]
        rev = list(reversed(periods))
        feat = c[2].selectbox("Summary for", rev, index=rev.index(default_feat),
                              format_func=lambda p: p + ("  (in progress)" if p == cur_key else ""))
        per = pd.Period(feat, freq={"Monthly": "M", "Quarterly": "Q", "Yearly": "Y"}[gi])
        p_start, p_end = per.start_time.date(), min(per.end_time.date(), _today())
    period_where = f"{WIRE_FIELD}>={p_start:%Y-%m-%d} AND {WIRE_FIELD}<={p_end:%Y-%m-%d}"

    cfg = lg_config()
    with st.spinner(f"Loading {feat}…"):
        fg = apply_lg_correction(add_intervals(run_soql(
            inst, tok, f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND {period_where}")), cfg, same_day)
    if fg.empty:
        st.info(f"No draws were wired in {feat}."); return

    turn, wait, total = fg["_fund_bd"].dropna(), fg["_wait_bd"].dropna(), fg["_days"].dropna()
    funded = pd.to_numeric(fg.get(AMT), errors="coerce").sum() if AMT else None
    same = (turn == same_day).mean() * 100 if len(turn) else None
    fast = (turn <= 3 + same_day).mean() * 100 if len(turn) else None
    issues = int(fg["_check"].isin(["Package date is after the wire date",
                                    "Package date is before the draw was requested"]).sum())
    holds = int(flight["Status__c"].astype(str).str.contains("Hold|Pending Borrower|Revision", case=False,
                                                              na=False).sum()) if "Status__c" in flight else 0

    k = st.columns(4)
    k[0].metric("Turn-time · package → wire", _bd(turn.mean()) if len(turn) else "—",
                f"{same:.0f}% same day" if same is not None else None, delta_color="off", delta_arrow="off",
                border=True, help="Average business days from a full draw package received to the wire — "
                                  "the official turn-time.")
    k[1].metric("Borrower side · request → package", _bd(wait.mean()) if len(wait) else "—", border=True,
                help="Average business days from the borrower's request to a full package: inspections, "
                     "title, missing documents and other conditions.")
    k[2].metric("Request → wire", _bd(total.mean()) if len(total) else "—", border=True,
                help="Average business days from the borrower's request to the wire.")
    k[3].metric(f"Draws funded · {feat}", f"{len(fg):,}", money(funded) if funded is not None else None,
                delta_color="off", delta_arrow="off", border=True)
    k2 = st.columns(4)
    k2[0].metric("Funded within 3 days of a full package", f"{fast:.0f}%" if fast is not None else "—",
                 border=True, help="Share of draws wired within 3 business days of the full package.")
    k2[1].metric("Package received on file", f"{fg['_pkg'].notna().mean() * 100:.0f}%",
                 f"{issues} to check" if issues else None, delta_color="off", delta_arrow="off", border=True,
                 help="Share of these draws with a full-draw-package-received date — what turn-time is measured "
                      "from. 'To check' = dates out of order, left out of the averages.")
    k2[2].metric("Median turn-time", _bd(turn.median()) if len(turn) else "—", border=True)
    k2[3].metric("Open draws now", f"{len(flight):,}", f"{holds} on hold or waiting" if holds else None,
                 delta_color="off", delta_arrow="off", border=True)
    if len(turn) and len(wait):
        st.caption(f"In **{feat}**, a full package was wired in **{_bd(turn.mean())}** on average"
                   + (f" ({same:.0f}% the same day)" if same else "") + f"; getting to a full package took "
                   f"borrowers **{_bd(wait.mean())}**.")
    if not cfg:
        st.info("Request dates aren't available right now, so only package → wire timing is shown.")

    rows = []
    for p in periods:
        g = done[done["Period"] == p]
        amt = float(pd.to_numeric(g.get(AMT), errors="coerce").sum()) if AMT else None
        rows.append({"_raw": p, "Period": p + ("  (in progress)" if p == cur_key else ""), "Draws": len(g),
                     "Loans": g["Loan_Number__c"].nunique() if "Loan_Number__c" in g else None,
                     "Funded ($)": amt, "Average draw ($)": round(amt / len(g)) if amt is not None and len(g) else None})
    roll = pd.DataFrame(rows)
    ap = approvals_for(inst, tok, tuple(_col(fg, "Id").dropna()))
    table = _draw_table(fg, inst, ap)

    t_sum, t_draws, t_dl = st.tabs([":material/insights: Summary", f":material/list: Draws ({len(fg):,})",
                                    ":material/download: Download"])
    with t_sum:
        left, right = st.columns([1, 1.3])
        with left:
            st.markdown(f"**By {'month' if gran == 'Date range' else gran.lower()}**")
            st.dataframe(roll.drop(columns="_raw").iloc[::-1], hide_index=True, width="stretch",
                         column_config={"Funded ($)": st.column_config.NumberColumn("Funded", format="dollar", step=1),
                                        "Average draw ($)": st.column_config.NumberColumn("Average draw",
                                                                                          format="dollar", step=1)})
        with right:
            _render_charts(roll, fg, feat, same_day)
        with st.expander("How these numbers work"):
            st.markdown(
                "- **Requested** — when the borrower submitted the draw (Land Gorilla).\n"
                "- **Package received** — the *Full Draw Package Received* date in Salesforce; when it's blank, the "
                "date the draw was approved in Land Gorilla.\n"
                "- **Wired** — the Salesforce wire date.\n"
                "- **Turn-time** = package received → wired. **Borrower side** = requested → package received.\n"
                "- Business days skip weekends and federal holidays. Draws whose dates are out of order are left out "
                "of the averages and listed under *Data checks* on the Draws tab.")
    with t_draws:
        st.dataframe(table, hide_index=True, width="stretch", height=460, column_config=_table_config())
        flagged = table.assign(Issue=fg["_check"].values)
        flagged = flagged[flagged["Issue"].isin(["Package date is after the wire date",
                                                 "Package date is before the draw was requested"])]
        missing = int((fg["_check"] == "No package-received date").sum())
        if len(flagged) or missing:
            with st.expander(f"Data checks ({len(flagged) + missing})", icon=":material/fact_check:"):
                if missing:
                    st.caption(f"{missing} draw(s) have no package-received date yet, so they have no turn-time.")
                if len(flagged):
                    st.dataframe(flagged[["Loan #", "Draw #", "Deal", "Requested", "Package received", "Wired",
                                          "Issue", "Salesforce"]], hide_index=True, width="stretch",
                                 column_config=_table_config())
    with t_dl:
        _period_downloads(inst, tok, rt, feat, period_where, p_end, _excel_table(fg, table),
                          roll.drop(columns="_raw").rename(columns={"Period": "Period", "Draws": "Draws funded"}))


@st.fragment
def _period_downloads(inst, tok, rt, feat, period_where, p_end, xtable, roll):
    """Both downloads for the period. A fragment, so building a file doesn't reload the whole page."""
    tpl = _find_template()
    with st.container(border=True):
        st.markdown(f"**Funding-request file** · Melanie's workbook, one row per draw wired in {feat}")
        e = st.columns([1.4, 1])
        who = e[0].text_input("Submitted by (optional)", key=f"pby_{feat}")
        fdate = e[1].date_input("Scheduled funding date", value=p_end, key=f"pdate_{feat}")
        if not tpl:
            st.error("The funding-request template is missing from the app.")
        elif st.button("Build funding-request file", type="primary", icon=":material/description:", key=f"mel_{feat}"):
            with st.spinner("Filling the template…"):
                rows, ftypes, _ = funding_request_rows(inst, tok, rt, period_where)
                xlsx = fill_melanie_template(tpl, rows, ftypes, who, fdate, file_summary(inst, tok, rt, rows))
            st.download_button(f"Download funding-request file ({len(rows):,} draws)", xlsx, mime=XLSX_MIME,
                               file_name=f"Construction-Renovation Funding Request {feat}.xlsx",
                               icon=":material/download:", key=f"meldl_{feat}", on_click="ignore")
            st.toast("Funding-request file is ready.", icon="✅")
    with st.container(border=True):
        st.markdown("**Turn-time data** · every draw with its milestone dates, business days, approvals and notes")
        if st.button("Build turn-time workbook", icon=":material/table_view:", key=f"xls_{feat}"):
            with st.spinner("Building the workbook…"):
                xlsx = build_excel({"Summary": roll, "Draws": xtable})
            st.download_button("Download turn-time workbook", xlsx, mime=XLSX_MIME, icon=":material/download:",
                               file_name=f"Draw turn-time {feat}.xlsx", key=f"dl_{feat}", on_click="ignore")


def _render_charts(roll: pd.DataFrame, fg: pd.DataFrame, feat: str, same_day: int):
    """Turn-time distribution for the period, plus dollars and draws per period."""
    if alt is None:
        st.bar_chart(roll.set_index("_raw")["Funded ($)"])
        return
    turn = fg["_fund_bd"].dropna()
    if len(turn):
        top = 10
        days = turn.clip(upper=top).astype(int)
        dist = days.value_counts().reindex(range(same_day, top + 1), fill_value=0).rename_axis("d").reset_index(
            name="Draws")
        dist["Business days"] = dist["d"].map(lambda d: f"{d}+" if d == top else str(d))
        st.markdown(f"**Package → wire, {feat}**")
        st.altair_chart(alt.Chart(dist).mark_bar(color="#4C78A8").encode(
            x=alt.X("Business days:N", sort=list(dist["Business days"]), axis=alt.Axis(labelAngle=0)),
            y=alt.Y("Draws:Q"), tooltip=["Business days", "Draws"]), width="stretch")
    r = roll.drop(columns="Period").rename(columns={"_raw": "Period"})
    st.markdown("**Funded per period**")
    st.altair_chart(alt.Chart(r).mark_bar(color="#54A24B").encode(
        x=alt.X("Period:N", sort=list(r["Period"]), title=None),
        y=alt.Y("Funded ($):Q", title=None, axis=alt.Axis(format="$,.0s")),
        tooltip=["Period", alt.Tooltip("Funded ($):Q", format="$,.0f"), "Draws"]), width="stretch")


# ───────────────────────────── Find a draw ──────────────────────────────────
def lg_filenumber(loan_no) -> str | None:
    """LG file number is 'rb0'+the SF loan number, but only for the clean numeric RB0 book."""
    s = str(loan_no).strip()
    return f"rb0{s}" if s.isdigit() else None


def render_loan_detail(inst, tok, rt, sel, same_day):
    st.header("Find a draw")
    text = st.text_input("Search", key="ld_q", placeholder="Loan #, property or deal name, borrower, or draw #",
                         help="Type any part of it.")
    if not text:
        st.info("Search by loan number, property / deal name, borrower or draw number to see each draw's full "
                "cycle.", icon=":material/search:")
        return
    q = soql_escape(text)
    searchable = [f for f in ["Loan_Number__c", "Deal__r.Name", "Borrower_Name_Text__c",
                              "Loan_Advance_Number__c"] if f in sel]
    ors = " OR ".join(f"{f} LIKE '%{q}%'" for f in searchable) or "Id != null"
    with st.spinner("Searching…"):
        df = add_intervals(run_soql(inst, tok, f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' "
                                               f"AND ({ors}) ORDER BY Loan_Number__c, {REQ_FIELD} "
                                               f"LIMIT {SEARCH_LIMIT + 1}"))
    if df.empty:
        st.warning("No matching draws."); return
    if len(df) > SEARCH_LIMIT:
        df = df.head(SEARCH_LIMIT)
        st.warning(f"Showing the first {SEARCH_LIMIT} matching draws — narrow the search to see the rest.")
    cfg = lg_config()
    with st.spinner("Loading draw history…"):
        df = apply_lg_correction(df, cfg, same_day)
        ap = approvals_for(inst, tok, tuple(_col(df, "Id").dropna()))
    nloans = df["Loan_Number__c"].nunique() if "Loan_Number__c" in df else len(df)
    st.caption(f"{len(df)} draw(s) on {nloans} loan(s)")
    tmpl = (cfg or {}).get("template_id", LG_TEMPLATE_ID_DEFAULT)

    for loan_no, g in df.groupby("Loan_Number__c", dropna=False):
        def first(col):
            return _txt(g[col].dropna().iloc[0]) if col in g and g[col].notna().any() else ""
        with st.container(border=True):
            st.subheader(f"Loan {loan_no} · {first('Deal__r.Name')}")
            sub = [x for x in (first("Deal__r.Account.Name"), first("Borrower_Name_Text__c")) if x]
            st.caption("  ·  ".join(sub + [f"{len(g)} draw(s)"]))
            for _, row in g.iterrows():
                _render_draw(row, ap, inst)
            if cfg:
                with st.expander("Construction progress", icon=":material/construction:"):
                    _render_lg_for_loan(cfg, tmpl, loan_no, g)


def _render_draw(row: pd.Series, ap: pd.DataFrame, inst: str):
    draw_no = _txt(row.get("Loan_Advance_Number__c")) or _txt(row.get("Name")) or "draw"
    rid = _txt(row.get("Id"))
    head = st.columns([4, 1])
    status = _txt(row.get("Status__c"))
    head[0].markdown(f"**Draw {draw_no}**" + (f" &nbsp; :blue-badge[{status}]" if status else ""))
    if rid:
        head[1].link_button("Open in Salesforce", f"{inst.rstrip('/')}/{rid}", icon=":material/open_in_new:",
                            width="stretch")
    m = st.columns(4)
    m[0].metric("Funded", money(row.get(NET_FIELD)) if pd.notna(row.get(NET_FIELD)) else "—", border=True)
    m[1].metric("Turn-time", _bd(row.get("_fund_bd")), border=True,
                help="Business days from full draw package received to wire")
    m[2].metric("Borrower side", _bd(row.get("_wait_bd")), border=True,
                help="Business days from the request to a full package")
    m[3].metric("Request → wire", _bd(row.get("_days")), border=True)

    dates = {label: _d(row.get(field)) for label, field in MILESTONES}
    dates["Requested"] = _d(row.get("_lg_req"))                 # the real request date
    dates["Full draw package received"] = _d(row.get("_pkg"))   # Salesforce's date, or the LG approval
    steps = pd.DataFrame({"Milestone": list(dates), "Date": pd.to_datetime(list(dates.values()))})
    steps["Done"] = steps["Date"].notna()
    st.dataframe(steps, hide_index=True, width="stretch",
                 column_config={"Date": st.column_config.DateColumn(format="M/D/YYYY"),
                                "Done": st.column_config.CheckboxColumn(width="small")})
    if _txt(row.get("_check")):
        st.caption(f":material/error: {row['_check']}")
    a = ap.loc[rid] if rid and rid in ap.index else None
    if a is not None and _txt(a.get("Approved by")):
        when = f" on {_mdy(a['Approved on'])}" if pd.notna(a.get("Approved on")) else ""
        st.markdown(f":material/verified: **Approved by {a['Approved by']}**{when}")
        if _txt(a.get("Approval comments")):
            st.caption(_txt(a["Approval comments"]).replace("\n", "  \n"))
    elif _txt(row.get("LastModifiedBy.Name")):
        st.caption(f"Last updated by {row['LastModifiedBy.Name']} on {_mdy(_d(row.get('LastModifiedDate')))}"
                   if _d(row.get("LastModifiedDate")) else f"Last updated by {row['LastModifiedBy.Name']}")
    vals = pd.to_numeric(pd.Series([row.get(f) for f, _ in CONSTRUCTION_FIELDS]), errors="coerce")
    money_bits = [f"{lbl}: {money(v)}" for (_, lbl), v in zip(CONSTRUCTION_FIELDS, vals) if pd.notna(v) and v]
    if money_bits:
        st.caption("  ·  ".join(money_bits))
    note = _txt(row.get(NOTES_FIELD))
    if note:
        with st.expander("Notes", icon=":material/sticky_note_2:"):
            st.text(note)


def _render_lg_for_loan(cfg: dict, tmpl: str, loan_no, g: pd.DataFrame):
    fn = lg_filenumber(loan_no)
    if fn:
        ov = lg_loan_overview(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), str(tmpl), str(loan_no))
        if ov and not ov.get("_error"):
            m = st.columns(4)
            if ov.get("project_pct") is not None:
                pct = float(ov["project_pct"])
                m[0].metric("Project complete", f"{pct:.0f}%", border=True)
                m[0].progress(min(max(pct / 100, 0), 1.0))
            if ov.get("to_finish") is not None:
                m[1].metric("Remaining to fund", money(ov["to_finish"]), border=True)
            if ov.get("balance") is not None:
                m[2].metric("Loan balance", money(ov["balance"]), border=True)
            if ov.get("last_draw"):
                m[3].metric("Last draw", _mdy(ov["last_draw"]), border=True)
            bits = [f"**{lbl}:** {_mdy(ov[key]) if isinstance(ov[key], date) else ov[key]}"
                    for lbl, key in [("Status", "status"), ("Program", "program"), ("Funded", "funded_date"),
                                     ("Due", "due_date")] if ov.get(key)]
            loc = " ".join(x for x in [ov.get("city"), ov.get("state")] if x)
            if loc:
                bits.append(f"**Location:** {loc}")
            if bits:
                st.caption("  ·  ".join(bits))
            if ov.get("risk"):
                st.caption("**Risk:** " + " · ".join(ov["risk"]))
        else:
            st.caption("No construction file found for this loan.")
    draw_rows, payees = [], []
    for _, row in g.iterrows():
        container = row.get(CONTAINER_FIELD)
        if pd.isna(container) or not container:
            continue
        det = lg_draw_detail(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), str(container))
        sf_draw = _txt(row.get("Loan_Advance_Number__c")) or _txt(row.get("Name"))
        if not det or det.get("_error"):
            continue
        draw_rows.append({"Draw": sf_draw, "Type": det.get("type"), "Status": det.get("status"),
                          "Requested": det.get("created"), "Submitted": det.get("submitted"),
                          "Approved": det.get("approved"), "Amount": det.get("amount")})
        payees += [{"Draw": sf_draw, "Payee": p["payee"], "Amount": p.get("amount")} for p in det.get("payees") or []]
    if draw_rows:
        st.markdown("**Draws**")
        ddf = pd.DataFrame(draw_rows)
        ddf["Amount"] = pd.to_numeric(ddf["Amount"], errors="coerce")
        dcol = st.column_config.DateColumn(format="M/D/YYYY")
        st.dataframe(ddf, hide_index=True, width="stretch",
                     column_config={"Requested": dcol, "Submitted": dcol, "Approved": dcol,
                                    "Amount": st.column_config.NumberColumn(format="dollar", step=1)})
        if payees:
            with st.expander(f"Payees ({len(payees)})"):
                pdf = pd.DataFrame(payees)
                pdf["Amount"] = pd.to_numeric(pdf["Amount"], errors="coerce")
                st.dataframe(pdf, hide_index=True, width="stretch",
                             column_config={"Amount": st.column_config.NumberColumn(format="dollar", step=1)})


# ───────────────────────────── Funding request (weekly file) ────────────────
def render_funding_request(inst, tok, rt):
    st.header("Funding request")
    st.caption("Melanie's weekly *Construction/Renovation Funding Request* workbook, filled from Salesforce — "
               "one row per draw.")
    tpl = _find_template()
    if not tpl:
        st.error("The funding-request template is missing from the app."); return
    with st.container(border=True):
        c = st.columns([1.3, 1, 1.7])
        who = c[0].text_input("Submitted by (optional)", key="fr_by")
        fdate = c[1].date_input("Scheduled funding date", value=_today(), key="fr_date")
        mode = c[2].segmented_control("Draws", ["Open (not yet wired)", "Wired in a date range"],
                                      default="Open (not yet wired)", key="fr_mode") or "Open (not yet wired)"
        if mode.startswith("Open"):
            where = OPEN_WHERE
        else:
            rng = st.date_input("Wire dates", value=(_today() - timedelta(days=7), _today()), key="fr_rng")
            if len(rng) != 2:
                st.info("Pick a start and an end date."); return
            where = f"{WIRE_FIELD}>={rng[0]:%Y-%m-%d} AND {WIRE_FIELD}<={rng[1]:%Y-%m-%d}"
    with st.spinner("Loading draws…"):
        rows, ftypes, _ = funding_request_rows(inst, tok, rt, where)
    if rows.empty:
        st.info("No draws in that range."); return

    f = st.columns(2)
    people = sorted({p for col in PERSON_FIELDS if col in rows for p in rows[col].dropna()})
    if people:
        person = f[0].selectbox("Coordinator / analyst", ["Everyone"] + people, key="fr_person",
                                help="Everyone gives the full list. Pick a person to match the file Salesforce "
                                     "generates for them.")
        if person != "Everyone":
            rows = rows[np.logical_or.reduce([rows[col].eq(person) for col in PERSON_FIELDS if col in rows])]
    if "Status__c" in rows:
        statuses = sorted(rows["Status__c"].dropna().unique())
        rows = rows[rows["Status__c"].isin(f[1].multiselect("Statuses", statuses, default=statuses,
                                                            key=f"fr_status_{mode}"))]
    total = pd.to_numeric(rows["U"], errors="coerce").sum() if "U" in rows else None
    m = st.columns(3)
    m[0].metric("Draws", f"{len(rows):,}", border=True)
    m[1].metric("Loans", f"{rows['B'].nunique() if 'B' in rows else 0:,}", border=True)
    m[2].metric("Current draw amount", money(total) if total is not None else "—", border=True)
    preview = pd.DataFrame({"Loan #": _col(rows, "B"), "Deal": _col(rows, "C"), "Borrower": _col(rows, "D"),
                            "Status": _col(rows, "Status__c"), "Current draw": pd.to_numeric(_col(rows, "U"),
                                                                                            errors="coerce"),
                            "Coordinator": _col(rows, "Advance_Coordinator__r.Name")})
    st.dataframe(preview, hide_index=True, width="stretch", height=320,
                 column_config={"Current draw": st.column_config.NumberColumn(format="dollar", step=1)})
    _funding_download(inst, tok, rt, tpl, rows, ftypes, who, fdate)


@st.fragment
def _funding_download(inst, tok, rt, tpl, rows, ftypes, who, fdate):
    if st.button("Build funding-request file", type="primary", icon=":material/description:",
                 disabled=rows.empty, key="fr_build"):
        with st.spinner("Filling the template…"):
            xlsx = fill_melanie_template(tpl, rows, ftypes, who, fdate, file_summary(inst, tok, rt, rows))
        st.download_button(f"Download funding-request file ({len(rows):,} draws)", xlsx, mime=XLSX_MIME,
                           icon=":material/download:", on_click="ignore", key="fr_dl",
                           file_name=f"Construction-Renovation Funding Request {fdate.month}_{fdate.day}_{fdate.year}.xlsx")
        st.toast("Funding-request file is ready.", icon="✅")


# ───────────────────────────────── main ─────────────────────────────────────
def main():
    st.set_page_config(page_title="Construction Draw Tracker", page_icon="🏗️", layout="wide")

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

    if sf is None and not cfg:                           # nothing to log in with
        st.title("🏗️ Construction Draw Tracker")
        st.error(err or "Salesforce isn't set up for this app yet.")
        st.caption("Add the Salesforce login to the app's secrets (username / password / security token, or the "
                   "OAuth client settings).")
        st.stop()
    if sf is None:                                       # OAuth path: not logged in yet, or the last try failed
        st.title("🏗️ Construction Draw Tracker")
        if err:
            st.error(err)
        st.info("Log in to Salesforce to load the construction draws.")
        st.link_button("Log in to Salesforce", login_url(cfg), type="primary", icon=":material/login:")
        st.stop()

    with st.sidebar:
        if mode == "oauth" and st.button("Log out", icon=":material/logout:", width="stretch"):
            clear_sf_session(); st.rerun()
        with st.expander("Settings", icon=":material/tune:"):
            same_day = 1 if st.radio("A same-day wire counts as", ["0 business days", "1 business day"],
                                     help="Applies to turn-time and request → wire.").startswith("1") else 0

    inst = st.session_state["salesforce_auth"]["instance_url"]
    tok = st.session_state["salesforce_auth"]["access_token"]
    start = st.session_state.get("_start_page", "turn-time")       # lets tests open a given page
    try:                          # the record-type / describe lookups can hit an expired session too
        rt = construction_rt(inst, tok)
        if not rt:
            st.error("Couldn't find the Construction Advance draws in Salesforce."); st.stop()
        sel = select_fields(inst, tok)
        pg = st.navigation([
            st.Page(lambda: render_pipeline(inst, tok, rt, sel, same_day), title="Turn-time",
                    icon=":material/timer:", url_path="turn-time", default=start == "turn-time"),
            st.Page(lambda: render_loan_detail(inst, tok, rt, sel, same_day), title="Find a draw",
                    icon=":material/search:", url_path="find", default=start == "find"),
            st.Page(lambda: render_funding_request(inst, tok, rt), title="Funding request",
                    icon=":material/request_quote:", url_path="funding-request", default=start == "funding-request"),
        ])
        pg.run()
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
