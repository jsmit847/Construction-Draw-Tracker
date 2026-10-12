"""
Construction Draw Tracker  (Streamlit)
======================================
ONE fused tracker — not separate Salesforce vs Land Gorilla views.

Salesforce is the spine: the Construction Advance record type IS the small-balance
RTL / fix-and-flip (RB0) book. Queried from the Advance__c OBJECT (not the pipeline
report), so paid-off draws are retained — closing the "paid off today drops off the
report" gap Melanie flagged. It carries the rb number (Loan_Number__c), the milestone
dates, Notes__c, and the real dollars.

Land Gorilla is the source of truth wherever it has the data; Salesforce's own dates are
unreliable. Through DrawContainerId__c, each draw's request date is Land Gorilla's createdDate
and its wire date Land Gorilla's funding (effective) date; Salesforce's wire date fills in only
where there's no Land Gorilla link. The full-package date is Salesforce's "Date Full Draw
Package Received", recorded since Aug 3, 2026. Turn-time = full package -> wire, in BUSINESS days.

Three pages on the same data:
  • Turn-time — monthly / quarterly / yearly turn-time KPIs, rollup, charts, the package-date
    check, a draw & deal details view, and the turn-time and funding-request downloads.
  • Find a draw — search a borrower / property / loan # and see each draw's full cycle, most
    recent request first; "Who's falling short" shows how long each account takes to reach a
    full package and what its draws were waiting on (delay reasons read from the notes).
  • Funding request — Melanie's weekly "Construction/Renovation Funding Request" workbook,
    filled from Land Gorilla and Salesforce into the committed template (Advance_template.xlsx).

Secrets (.streamlit/secrets.toml — see .streamlit/secrets.toml.example): [salesforce]
(username/password/security_token, or OAuth client_id/client_secret/redirect_uri/auth_host),
[landgorilla] (user, password, verify), and — only on the deployment whose URL is the approved
Salesforce callback, which runs this file — [relay] forward_to. The main deployment runs streamlit_app.py,
which imports this file and so never relays.  Run:  streamlit run draw_dashboard.py
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
from functools import wraps
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

# ─────────────── OAuth relay (only on the deployment Salesforce redirects back to) ───────────────
# The Salesforce connected app approves one callback URL: the original deployment's. When that deployment has
# a [relay] forward_to secret, it hands this app's logins (state starting OAUTH_STATE_PREFIX) on to forward_to,
# which holds the PKCE verifier and exchanges the code itself. The destination comes only from secrets, never
# from the request, so this can't be used as an open redirect. Deployments without [relay] skip it entirely, and
# so does a deployment that imports this file instead of running it (streamlit_app.py).
OAUTH_STATE_PREFIX = "cdt"
if __name__ == "__main__":
    try:
        _forward_to = str(st.secrets.get("relay", {}).get("forward_to") or "").strip()
    except Exception:                                   # no secrets.toml at all (e.g. a local checkout)
        _forward_to = ""
    _code, _state = st.query_params.get("code"), str(st.query_params.get("state") or "")
    if _forward_to and _code and _state.startswith(OAUTH_STATE_PREFIX):
        st.query_params.clear()                         # the unused code leaves the address bar; only the link has it
        st.link_button("Open Construction Draw Tracker", _forward_to + ("&" if "?" in _forward_to else "?")
                       + urlencode({"code": _code, "state": _state}))
        st.stop()

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
         NET_FIELD, GROSS_FIELD, "LastModifiedBy.Name", "LastModifiedDate", "CreatedDate"]
        + [f for f, _ in CONSTRUCTION_FIELDS] + [m[1] for m in MILESTONES])
TERMINAL = ["Completed", "Cancelled", "Rescinded", "Rejected by Capital Partner"]
OPEN_WHERE = f"{WIRE_FIELD}=null AND Status__c NOT IN ({','.join(repr(s) for s in TERMINAL)})"
# Status__c stages a weekly funding request covers (approval through release) — the export's default filter.
FUNDING_STAGES = ("Pending Approval", "Approved", "Pending Capital Partner Release", "Pending Release")
SEARCH_LIMIT = 200    # loan-detail search: most draws rendered (and dated in Land Gorilla) per query
MAX_LOAN_SECTIONS = 40   # ... and most loans shown with their track record and construction progress
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
    state = OAUTH_STATE_PREFIX + secrets.token_urlsafe(24)    # the prefix lets the callback deployment relay it
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


def _session_expired(exc: BaseException) -> bool:
    msg = str(exc)
    return "INVALID_SESSION_ID" in msg or "Session expired" in msg or type(exc).__name__ == "SalesforceExpiredSession"


def _restart_session():
    """Forget an expired Salesforce session and rerun the whole app once, so main() logs in again (a second failure
       stops with a message)."""
    try:
        sf_login_credentials.clear()          # force a fresh login on the credentials path
    except Exception:
        pass
    clear_sf_session()
    if not st.session_state.get("_sf_retried"):
        st.session_state["_sf_retried"] = True
        st.rerun(scope="app")
    st.warning("Your Salesforce session expired — refresh the page to log in again."); st.stop()


def sf_fragment(fn):
    """st.fragment that recovers from an expired Salesforce session the way a full rerun does: a fragment rerun
       doesn't pass through main(), where that recovery lives."""
    @wraps(fn)
    def run(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:              # st.rerun / st.stop are BaseExceptions and pass straight through
            if _session_expired(exc):
                _restart_session()
            raise
    return st.fragment(run)


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


LG_TTL_DONE, LG_TTL_OPEN, LG_TTL_LOAN = 86400, 1800, 6 * 3600   # funded draws are settled; open ones move


@st.cache_resource(show_spinner=False)
def _lg_store() -> dict:
    """Process-wide cache of successful Land Gorilla reads: {(kind, key): (fetched_at, value)}. Failures aren't
       stored, so a draw that hit a rate limit or timeout is simply fetched again next time."""
    return {}


def _lg_batch(client: "LGClient", paths: dict) -> dict:
    """GET many Land Gorilla paths CONCURRENTLY — {key: path} -> {key: json or None}. One token, one session per
       worker thread. ~635 serial calls took ~80s; in parallel it's a few seconds."""
    from concurrent.futures import ThreadPoolExecutor
    import requests, threading
    if not paths:
        return {}
    token = client.token()                              # pre-warm so threads don't race to refresh it
    tl = threading.local()

    def one(item):
        key, path = item
        s = getattr(tl, "s", None)
        if s is None:
            s = tl.s = requests.Session()
        try:
            r = s.get(f"{client.BASE}{path}", timeout=30, verify=client.verify,
                      headers={"Authorization": f"Bearer {token}", "Accept": "application/json"})
            return key, (r.json() if r.ok else None)
        except Exception:
            return key, None

    with ThreadPoolExecutor(max_workers=16) as ex:
        return dict(ex.map(one, paths.items()))


def _iso(d) -> str | None:
    return d.isoformat() if isinstance(d, date) else None


def lg_draws(user: str, password: str, verify: bool, containers) -> dict:
    """{container: draw} from GET /api/clm/draw/{id}: every date on the draw (ISO), amount, requested amount,
       payee names, status and type. Fetched in parallel; cached per draw."""
    store, now = _lg_store(), time.time()
    want = [str(c) for c in dict.fromkeys(containers) if isinstance(c, str) and c]

    def fresh(c):
        hit = store.get(("draw", c))
        return hit and now - hit[0] < (LG_TTL_DONE if hit[1].get("effective") else LG_TTL_OPEN)

    need = {c: f"/api/clm/draw/{c}" for c in want if not fresh(c)}
    if need:
        for c, payload in _lg_batch(LGClient(user, password, verify=verify), need).items():
            det = parse_draw_detail(payload) if payload else {}
            if det.get("created"):
                store[("draw", c)] = (time.time(), {
                    "created": _iso(det["created"]), "submitted": _iso(det.get("submitted")),
                    "approved": _iso(det.get("approved")), "effective": _iso(det.get("funded")),
                    "amount": det.get("amount"), "requested": det.get("requested_total"),
                    "payees": [p["payee"] for p in det.get("payees") or [] if p.get("payee")],
                    "status": det.get("status"), "type": det.get("type"), "name": det.get("name"),
                    "requested_by": det.get("requested_by"), "dates": det.get("dates") or {}})
    return {c: store[("draw", c)][1] for c in want if ("draw", c) in store}


def lg_loans(user: str, password: str, verify: bool, template_id: str, loan_numbers) -> dict:
    """{loan #: overview} for many loans at once (Land Gorilla file number rb0 + loan #): project completion,
       funded / due dates, balance, remaining to fund, team, risk. Fetched in parallel; cached per loan."""
    store, now = _lg_store(), time.time()
    want = [str(n).strip() for n in dict.fromkeys(loan_numbers) if lg_filenumber(n)]

    def fresh(n):                          # a loan Land Gorilla doesn't have is remembered for a while too
        hit = store.get(("loan", n))
        return hit and now - hit[0] < (LG_TTL_LOAN if hit[1] is not None else LG_TTL_OPEN)

    need = [n for n in want if not fresh(n)]
    if need:
        client = LGClient(user, password, verify=verify)
        found = _lg_batch(client, {n: "/api/clm/loan?" + urlencode({"fileNumber": f"rb0{n}"}) for n in need})
        for n, p in found.items():
            if isinstance(p, (dict, list)) and not items_of(p):       # answered, with no such loan (not an error)
                store[("loan", n)] = (time.time(), None)
        ids = {n: items_of(p)[0].get("id") for n, p in found.items() if p and items_of(p)}
        ids = {n: i for n, i in ids.items() if i}
        tmpl = _lg_batch(client, {n: f"/api/clm/pipelineReportTemplates/{template_id}/loans/{i}" for n, i in ids.items()})
        summ = _lg_batch(client, {n: f"/api/clm/loan/{i}" for n, i in ids.items()})
        for n, i in ids.items():
            ov = {"lg_loan_id": i, **(parse_template_loan(tmpl[n]) if tmpl.get(n) else {})}
            sd = (summ.get(n) or {}).get("data", {}) if isinstance(summ.get(n), dict) else {}
            ov.update(balance=sd.get("loanBalance"), to_finish=sd.get("balanceToFinish"),
                      last_approved_draw=_lg_date(sd.get("lastApprovedDrawEffectiveDate")))
            store[("loan", n)] = (time.time(), ov)
    return {n: store[("loan", n)][1] for n in want if store.get(("loan", n), (0, None))[1] is not None}


PKG_LAUNCH = pd.Timestamp("2026-08-03")   # Salesforce's "Date Full Draw Package Received" went live this day
CHECK_EXCLUDED = ("Package date is after the wire date", "Package date is before the draw was requested")


def apply_lg_correction(df: pd.DataFrame, cfg: dict | None, same_day: int) -> pd.DataFrame:
    """Each draw's timeline in business days, Land Gorilla first:
         requested    = LG createdDate
         wired        = LG effectiveDate (its funding date); Salesforce's wire date only when there's no LG link.
                        Only draws Salesforce shows as wired count, so an approved-but-unfunded draw isn't "wired".
         full package = Salesforce "Date Full Draw Package Received" — Land Gorilla has no such date (its approval
                        is stamped on funding day). Tracked from Aug 3, 2026, and used only when requested <=
                        package <= wired. Never filled in from the approval date.
       Turn-time = full package -> wired. Draws whose dates don't line up are left out and named in _check."""
    if df.empty:
        return df
    df = df.copy()
    lg = {}
    if cfg and CONTAINER_FIELD in df:
        try:
            lg = lg_draws(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), df[CONTAINER_FIELD].tolist())
        except Exception as exc:      # no exception text: a malformed secret would be echoed to every viewer
            st.warning(f"Land Gorilla is unavailable right now ({type(exc).__name__}) — timing is incomplete.")
    info = _col(df, CONTAINER_FIELD).map(lambda c: lg.get(c) if isinstance(c, str) else None)
    created, approved, effective, submitted, requested_by = (
        info.map(lambda d, k=k: d.get(k) if d else None)
        for k in ("created", "approved", "effective", "submitted", "requested_by"))
    sf_wire = _col(df, WIRE_FIELD)
    lg_wire = effective.where(effective.notna(), approved)
    wire = lg_wire.where(lg_wire.notna(), sf_wire).where(sf_wire.notna())
    c, w = (pd.to_datetime(s, errors="coerce").dt.normalize() for s in (created, wire))
    raw = pd.to_datetime(_col(df, PKG_FIELD), errors="coerce").dt.normalize()
    ok = raw.notna() & (raw >= PKG_LAUNCH) & (raw <= w) & (c.isna() | (c <= raw))
    pkg = raw.where(ok)

    df["_lg_req"], df["_lg_appr"], df["_wire"], df["_pkg"] = created, approved, w, pkg
    df["_lg_pkg"] = approved                                           # kept for older callers
    df["_lg_sub"], df["_lg_by"] = submitted, requested_by              # for the draw's timeline and card
    for col, key in (("_lg_amount", "amount"), ("_lg_requested", "requested"), ("_lg_status", "status"),
                     ("_lg_payees", "payees")):
        df[col] = info.map(lambda d, k=key: d.get(k) if d else None)
    df["_days"] = bdays(created, w, same_day)                          # request -> wired (per draw / account)
    df["_days_src"] = np.where(c.notna(), "Land Gorilla", "")
    df["_fund_bd"] = bdays(pkg, w, same_day)                           # turn-time: full package -> wired
    df["_wait_bd"] = bdays(created, pkg, 0)                            # getting to a full package (account view)
    check = pd.Series(None, index=df.index, dtype=object)
    live = raw.notna() & (raw >= PKG_LAUNCH)                          # order checks only apply to real values
    check[raw.notna() & (raw < PKG_LAUNCH)] = "Package date predates the field (before Aug 3, 2026) — ignored"
    check[live & w.notna() & (raw > w)] = CHECK_EXCLUDED[0]
    check[live & c.notna() & (raw < c)] = CHECK_EXCLUDED[1]
    check[w.notna() & (w >= PKG_LAUNCH) & raw.isna()] = "No package-received date"
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
    cols, ftypes = {}, {}                  # every candidate that exists; each row takes the first that has a value
    for col, cands in MELANIE_MAP.items():
        for path in _cands(cands):
            meta = resolve_path(inst, tok, "Advance__c", path)
            if _usable(meta):
                cols.setdefault(col, []).append(path)
                ftypes.setdefault(col, meta["type"])
    unavailable = [col for col in MELANIE_MAP if col not in cols]
    prop_ref = next((f for f, m in adv_meta.items() if m["ref"] == ["Property__c"]), None)  # Advance -> Property
    extra = [p for p in ("Id", "Name", "Deal__c", "Status__c", "Loan_Advance_Number__c", WIRE_FIELD, CONTAINER_FIELD)
             + PERSON_FIELDS if resolve_path(inst, tok, "Advance__c", p)] + ([prop_ref] if prop_ref else [])
    checks = {label: path for label, path in DEAL_CHECKS.items() if _usable(resolve_path(inst, tok, "Advance__c", path))}
    paths = [p for ps in cols.values() for p in ps] + list(checks.values())
    adv = run_soql(inst, tok, f"SELECT {','.join(dict.fromkeys(extra + paths))} FROM Advance__c "
                              f"WHERE RecordTypeId='{rt}' AND {where} ORDER BY Loan_Number__c")
    if adv.empty:
        return adv, ftypes, unavailable
    out = adv[[c for c in extra if c in adv]].copy()
    for col, ps in cols.items():
        out[col] = _first_filled(adv, ps)
    for label, path in checks.items():                        # candidates for the columns no field is confirmed for
        out[f"chk:{label}"] = _col(adv, path).values
        ftypes[f"chk:{label}"] = resolve_path(inst, tok, "Advance__c", path)["type"]
    out["_multi_property"] = False

    # One row per property per draw, like Salesforce's own file: Property_Advance__c ties each draw to each
    # property with that property's approved amount. Portfolio rows keep that per-property amount.
    pa = _property_advances(inst, tok, _col(out, "Id").tolist())
    out["_prop_id"], out["_pa_amount"] = None, np.nan
    if not pa.empty and "Id" in out:
        out = out.drop(columns=["_prop_id", "_pa_amount"]).merge(pa, on="Id", how="left")
    out["_pa_count"] = out.groupby(_col(out, "Id").fillna("").values)["_prop_id"].transform("count").values
    out["_u_locked"] = out["_pa_count"] > 1                     # Land Gorilla's draw total can't be split by property
    out["U"] = [a if (lock or pd.notna(a)) else u               # a split draw never repeats the whole amount
                for a, u, lock in zip(out["_pa_amount"], _col(out, "U"), out["_u_locked"])]

    # Property columns — validated field by field, so one missing field can't blank the rest.
    try:
        pmeta = describe_object(inst, tok, "Property__c")
    except Exception:
        pmeta = {}
    pcols = {col: [f for f in _cands(cands) if _usable(pmeta.get(f))] for col, cands in MELANIE_PROPERTY_MAP.items()}
    pcols = {col: fs for col, fs in pcols.items() if fs}
    street = [f for f in STREET_PARTS if f in pmeta]
    unavailable += [col for col in MELANIE_PROPERTY_MAP if col not in pcols] + ([] if street else ["F"])
    pchecks = {label: f for label, f in PROPERTY_CHECKS.items() if _usable(pmeta.get(f))}
    pfields = list(dict.fromkeys([f for fs in pcols.values() for f in fs] + street + list(pchecks.values())
                                 + (["Name"] if "Name" in pmeta else [])))
    link = "Advance__c" if (pmeta.get("Advance__c") or {}).get("ref") == ["Advance__c"] else None  # Property -> draw
    direct = [p for p in list(_col(out, "_prop_id")) + (list(out[prop_ref]) if prop_ref and prop_ref in out else [])
              if isinstance(p, str) and p]
    by_id = {r["Id"]: r for r in _soql_in(inst, tok, ",".join(["Id"] + pfields), "Property__c", "Id",
                                          direct).to_dict("records")} if pfields and direct else {}
    per_deal, per_draw = {}, {}
    if pfields and "Deal__c" in pmeta and "Deal__c" in out:
        sel = ",".join(dict.fromkeys(["Id", "Deal__c"] + ([link] if link else []) + pfields))
        for r in _soql_in(inst, tok, sel, "Property__c", "Deal__c", out["Deal__c"]).to_dict("records"):
            per_deal.setdefault(r["Deal__c"], []).append(r)
            if link and pd.notna(r.get(link)):
                per_draw.setdefault(r[link], []).append(r)
    match = []
    for _, row in out.iterrows():
        pid = row.get("_prop_id") if isinstance(row.get("_prop_id"), str) else row.get(prop_ref) if prop_ref else None
        props, hits = per_deal.get(row.get("Deal__c"), []), per_draw.get(row.get("Id"), [])
        if isinstance(pid, str) and pid in by_id:              # the exact property (Property_Advance__c / lookup)
            match.append(by_id[pid])
        elif len(props) == 1:                                   # the deal's only property
            match.append(props[0])
        elif len(hits) == 1:                                    # the one property pointing at this draw
            match.append(hits[0])
        else:                                                   # a draw named for its property's address
            key = _addr_key(row.get("Name"))
            same = [p for p in props if key and key in (_addr_key(p.get("Name")), _addr_key(
                " ".join(_txt(p.get(f)) for f in street)))]
            match.append(same[0] if len(same) == 1 else None)
    out["_multi_property"] = [len(per_deal.get(d, [])) > 1 and m is None for d, m in zip(_col(out, "Deal__c"), match)]
    for col, fs in pcols.items():
        out[col] = [next((m.get(f) for f in fs if m and _txt(m.get(f))), None) for m in match]
        ftypes[col] = pmeta[fs[0]]["type"]
    for label, f in pchecks.items():
        out[f"chk:{label}"] = [m.get(f) if m else None for m in match]
        ftypes[f"chk:{label}"] = pmeta[f]["type"]
    out["F"] = [" ".join(_txt(m.get(f)) for f in street if _txt(m.get(f))) or None if m else None for m in match]
    return out.reset_index(drop=True), ftypes, sorted(set(unavailable), key=lambda c: (len(c), c))


# Candidates for the columns no Salesforce field is confirmed for yet. They go on the Prep worksheet (not the
# request sheet), so they can be checked against Melanie's own file before any is promoted to a real column.
DEAL_CHECKS = {
    "Completion LTC % — deal Max LTC": "Deal__r.LTC__c",
    "Completion LTC % — deal Total Loan LTC": "Deal__r.Total_Loan_LTC__c",
    "Asset Manager — deal Loan Manager": "Deal__r.Loan_Coordinator__r.Name",
}
PROPERTY_CHECKS = {
    "Completion LTC % — property Total Loan LTC": "Total_Loan_LTC__c",
    "Asset Manager — property Servicer Asset Manager": "Asset_Manager__c",
    "Estimated Completion — property Rehab Completion Date": "Rehab_Completion_Date__c",
}


def _property_advances(inst: str, tok: str, adv_ids) -> pd.DataFrame:
    """Property_Advance__c rows for these draws — columns Id (the draw), _prop_id, _pa_amount. Empty when the org
       doesn't have the object or the draws have no property split."""
    try:
        meta = describe_object(inst, tok, "Property_Advance__c")
    except Exception:
        return pd.DataFrame()
    adv_fk = next((f for f, m in meta.items() if m["ref"] == ["Advance__c"]), None)
    prop_fk = next((f for f, m in meta.items() if m["ref"] == ["Property__c"]), None)
    if not adv_fk or not prop_fk:
        return pd.DataFrame()
    amt = "Approved_Advance_Amount_Calc__c" if "Approved_Advance_Amount_Calc__c" in meta else next(
        (f for f, m in meta.items() if m["type"] == "currency" and "amount" in str(m.get("label", "")).lower()), None)
    try:
        pa = _soql_in(inst, tok, ",".join(["Id", adv_fk, prop_fk] + ([amt] if amt else [])), "Property_Advance__c",
                      adv_fk, adv_ids)
    except Exception:
        return pd.DataFrame()
    if pa.empty or adv_fk not in pa or prop_fk not in pa:
        return pd.DataFrame()
    return pd.DataFrame({"Id": pa[adv_fk].values, "_prop_id": pa[prop_fk].values,
                         "_pa_amount": pd.to_numeric(pa[amt], errors="coerce").values if amt in pa else np.nan})


def _first_filled(df: pd.DataFrame, paths: list) -> list:
    """Per row, the first of these columns that has a value (blank strings don't count)."""
    cols = [df[p] for p in paths if p in df]
    return [next((v for v in vals if _txt(v)), None) for vals in zip(*cols)] if cols else [None] * len(df)


_DIRECTIONS = {"n", "s", "e", "w", "north", "south", "east", "west", "ne", "nw", "se", "sw"}


def _addr_key(s) -> tuple | None:
    """'2902 East McLoughlin Blvd, Vancouver' -> ('2902', 'mcloughlin'): street number + first street word, so a
       draw named for its property's address can be matched to that property."""
    words = re.findall(r"[a-z0-9]+", _txt(s).lower())
    if len(words) < 2 or not re.fullmatch(r"\d+[a-z]?", words[0]):
        return None
    rest = [w for w in words[1:] if w not in _DIRECTIONS] or words[1:]
    return words[0], rest[0]


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
        return d.date() if pd.notna(d) else v              # a real Excel date (shown m/d/yyyy)
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


LG_COLUMNS = ("U", "AG", "M", "N", "Y", "AB", "W")       # columns Land Gorilla fills first


def lg_overlay(rows: pd.DataFrame, ftypes: dict, cfg: dict | None,
               loans: bool = True) -> tuple[pd.DataFrame, dict, dict]:
    """Land Gorilla first: each draw's amount and payees, and each loan's funded date, current due date, project
       completion, balance and remaining to fund. Salesforce stays only where Land Gorilla has nothing.
       Returns (rows, ftypes, {"draws": {...}, "loans": {...}}) — the raw LG data feeds the prep worksheet."""
    lg = {"draws": {}, "loans": {}}
    if rows.empty or not cfg:
        return rows, ftypes, lg
    rows, ftypes = rows.copy(), dict(ftypes)
    u, p, v = cfg["user"], cfg["password"], bool(cfg.get("verify", True))
    try:
        lg["draws"] = lg_draws(u, p, v, _col(rows, CONTAINER_FIELD).tolist())
        if loans:
            lg["loans"] = lg_loans(u, p, v, str(cfg.get("template_id", LG_TEMPLATE_ID_DEFAULT)),
                                   _col(rows, "B").tolist())
    except Exception as exc:          # no exception text: a malformed secret would be echoed to every viewer
        st.warning(f"Land Gorilla is unavailable right now ({type(exc).__name__}) — using Salesforce values only.")
        return rows, ftypes, lg
    dr = _col(rows, CONTAINER_FIELD).map(lambda c: lg["draws"].get(c) if isinstance(c, str) else None)
    lo = _col(rows, "B").map(lambda n: lg["loans"].get(_txt(n)) if _txt(n) else None)
    payees = dr.map(lambda d: "; ".join(d["payees"]) if d and d.get("payees") else None)
    for col, src, key, ftype in (("U", dr, "amount", "currency"), ("M", lo, "funded_date", "date"),
                                 ("N", lo, "due_date", "date"), ("Y", lo, "project_pct", "percent"),
                                 ("AB", lo, "balance", "currency"), ("W", lo, "to_finish", "currency")):
        got = src.map(lambda d: d.get(key) if d else None)
        locks = [col == "U" and x is not None and not pd.isna(x) and bool(x) for x in _col(rows, "_u_locked")]
        rows[col] = [b if lock else (a if _txt(a) else b) for a, b, lock in zip(got, _col(rows, col), locks)]
        if got.map(lambda x: _txt(x) != "").any():
            ftypes[col] = ftype
    rows["AG"] = payees.tolist()
    ftypes["AG"] = "string"
    return rows, ftypes, lg


def _prep_table(rows: pd.DataFrame, lg: dict, ap: pd.DataFrame | None, ftypes: dict | None = None) -> pd.DataFrame:
    """One line per row of the request with everything at hand for finishing it by hand: Land Gorilla's draw and
       loan details, the team on the loan, approvals and notes."""
    ftypes = ftypes or {}
    dr = _col(rows, CONTAINER_FIELD).map(lambda c: lg["draws"].get(c) if isinstance(c, str) else None)
    lo = _col(rows, "B").map(lambda n: lg["loans"].get(_txt(n)) if _txt(n) else None)
    g = lambda s, k: s.map(lambda d: d.get(k) if d else None)            # noqa: E731
    ids = _col(rows, "Id")
    t = pd.DataFrame({
        "Row": range(TEMPLATE_HEADER_ROW + 1, TEMPLATE_HEADER_ROW + 1 + len(rows)),
        "CV Loan #": _col(rows, "B").values, "Draw #": _col(rows, "Loan_Advance_Number__c").values,
        "Property": _col(rows, "F").values, "Deal": _col(rows, "C").values,
        "Status": _col(rows, "Status__c").values, "Coordinator": _col(rows, "Advance_Coordinator__r.Name").values,
        "This property's amount ($)": pd.to_numeric(_col(rows, "_pa_amount"), errors="coerce").values,
        "Whole draw amount ($)": pd.to_numeric(g(dr, "amount"), errors="coerce").values,
        "Requested amount ($)": pd.to_numeric(g(dr, "requested"), errors="coerce").values,
        "Paid to": dr.map(lambda d: "; ".join(d["payees"]) if d and d.get("payees") else None).values,
        "Draw type": g(dr, "type").values, "Requested": _dates(g(dr, "created")).values,
        "Approved / funded": _dates(g(dr, "effective")).values,
        "Project complete %": (pd.to_numeric(g(lo, "project_pct"), errors="coerce") / 100).values,
        "Term elapsed %": (pd.to_numeric(g(lo, "duration_pct"), errors="coerce") / 100).values,
        "Remaining to fund ($)": pd.to_numeric(g(lo, "to_finish"), errors="coerce").values,
        "Loan balance ($)": pd.to_numeric(g(lo, "balance"), errors="coerce").values,
        "Loan funded": _dates(g(lo, "funded_date")).values, "Current due date": _dates(g(lo, "due_date")).values,
        "Original due date": _dates(g(lo, "original_due")).values, "Program": g(lo, "program").values,
        "Team on the loan": lo.map(lambda d: ", ".join(d.get("users") or []) if d else None).values,
        "Risk flags": lo.map(lambda d: ", ".join(d.get("risk") or []) if d else None).values,
        "Property matched": ["No — portfolio loan" if bool(m) else "" for m in _col(rows, "_multi_property")],
    })
    if ap is not None and not ap.empty:
        for c in ("Approved by", "Approved on", "Approval comments"):
            t[c] = ids.map(ap[c]).values if c in ap else None
        t["Approved on"] = _dates(t["Approved on"])
    for c in [c for c in rows.columns if str(c).startswith("chk:")]:     # candidates to confirm against Melanie's file
        vals = pd.Series(rows[c].values)
        if ftypes.get(c) == "percent":
            vals = pd.to_numeric(vals, errors="coerce") / 100
        elif ftypes.get(c) in ("date", "datetime"):
            vals = _dates(vals)
        t["Check: " + c[4:] + (" %" if ftypes.get(c) == "percent" else "")] = vals.values
    t["Notes"] = _col(rows, "AF").values
    return t


def fill_melanie_template(template: str, rows: pd.DataFrame, ftypes: dict, submitted_by: str,
                          funding_date: date, summary: dict | None = None, prep: pd.DataFrame | None = None) -> bytes:
    """Open the committed template, fill the title block, and write one row per draw in its own column order.
       The other sheets are brought up to date too: Total Summary from Salesforce, the Cap Partners title
       blocks from this request, the template's stale hidden reference sheet removed, and a Prep worksheet
       added with everything needed to finish the columns no system holds."""
    import openpyxl
    from openpyxl.utils import column_index_from_string as ci
    wb = openpyxl.load_workbook(template)
    ws = wb[TEMPLATE_SHEET]
    _put(ws, 3, 5, _xl_text(submitted_by.strip()) or None)                   # E3 (optional)
    for name in ("Cap Partners Funded By CV", "Cap Partners"):              # their title blocks mirror E3:E5
        if name in wb.sheetnames:
            wb[name]["D4"] = f"='{TEMPLATE_SHEET}'!E4"
            wb[name]["D5"] = f"='{TEMPLATE_SHEET}'!E5"
            wb[name]["D4"].number_format = wb[name]["D5"].number_format = "m/d/yyyy"
    if "Sheet1" in wb.sheetnames:                                            # March-2024 servicer data, unused
        del wb["Sheet1"]
    if summary:
        _write_total_summary(wb, summary)
    ws["E4"], ws["E5"] = _today(), funding_date                              # real dates (today in ET)
    ws["E4"].number_format = ws["E5"].number_format = "m/d/yyyy"
    letters = list(dict.fromkeys(list(MELANIE_MAP) + list(MELANIE_PROPERTY_MAP) + ["F"] + list(LG_COLUMNS)))
    for r, (_, row) in enumerate(rows.iterrows(), TEMPLATE_HEADER_ROW + 1):
        for col in letters:
            v = _xl(row.get(col), ftypes.get(col))
            if v is not None:
                _put(ws, r, ci(col), v)
                if isinstance(v, date):
                    ws.cell(r, ci(col)).number_format = "m/d/yyyy"
        ws.cell(r, ci("Z")).value = 1                                        # Borrower EC % — 100%, as Salesforce's file
        if bool(row.get("_multi_property", False)):
            ws.cell(r, ci("A")).value = "Portfolio loan — couldn't match this draw to one property"
        for col, pattern in MELANIE_FORMULAS.items():     # every input filled and non-zero (no #DIV/0!)
            if all(ws.cell(r, ci(c)).value not in (None, "", 0) for c in re.findall(r"([A-Z]+)\{r\}", pattern)):
                ws.cell(r, ci(col)).value = pattern.format(r=r)
    if prep is not None and not prep.empty:
        _write_prep_sheet(wb, prep)
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def _write_prep_sheet(wb, prep: pd.DataFrame) -> None:
    """Append the Prep worksheet: a formatted helper table (header band, filters, frozen header)."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    ws = wb.create_sheet("Prep worksheet")
    head, band = Font(name="Calibri", size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3864")
    for j, col in enumerate(prep.columns, 1):
        c = ws.cell(1, j, col)
        c.font, c.fill = head, band
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, (_, row) in enumerate(prep.iterrows(), 2):
        for j, (col, v) in enumerate(row.items(), 1):
            try:
                if v is None or pd.isna(v):
                    continue
            except (TypeError, ValueError):
                pass
            if isinstance(v, datetime):
                v = v.date()
            if isinstance(v, np.generic):
                v = v.item()
            _put(ws, i, j, _xl_text(v) if isinstance(v, str) else v)
            cell = ws.cell(i, j)
            cell.font = Font(name="Calibri", size=10)
            if isinstance(v, date):
                cell.number_format = "m/d/yyyy"
            elif col.endswith("($)"):
                cell.number_format = '"$"#,##0'
            elif col.endswith("%"):
                cell.number_format = "0%"
            if col in ("Notes", "Approval comments"):
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    for j, col in enumerate(prep.columns, 1):                     # wide enough for what's in the column
        sample = [len(_txt(v)) for v in prep[col].head(200)]
        width = 48 if col in ("Notes", "Approval comments") else min(40, max([12, len(col) * 0.9 + 2]
                                                                              + [n + 2 for n in sample]))
        ws.column_dimensions[get_column_letter(j)].width = width
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(prep.columns))}{len(prep) + 1}"


def build_funding_file(inst: str, tok: str, rt: str, tpl: str, rows: pd.DataFrame, ftypes: dict,
                       submitted_by: str, funding_date: date) -> tuple[bytes, int]:
    """Everything that goes into one funding-request workbook: Land Gorilla first, Salesforce to fill gaps,
       the summary sheet, and the prep worksheet. Returns (xlsx bytes, number of draws)."""
    rows, ftypes, lg = lg_overlay(rows, ftypes, lg_config())
    ap = approvals_for(inst, tok, tuple(_col(rows, "Id").dropna()))
    xlsx = fill_melanie_template(tpl, rows, ftypes, submitted_by, funding_date,
                                 file_summary(inst, tok, rt, rows), _prep_table(rows, lg, ap, ftypes))
    return xlsx, len(rows)


# ───────────────────────────── approvals ────────────────────────────────────
APPROVAL_COLS = ["Approval", "Approval status", "Approved by", "Approved on", "Approval comments", "_steps"]
STEP_NAMES = {"Started": "Sent for approval", "Approved": "Approved", "Rejected": "Rejected"}


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
    """Who approved each draw, when, and what they wrote — indexed by Advance Id. Salesforce's approval history gives
       every step (sent for approval, approved, rejected) and the person who took it; a draw can need more than one
       approver. The org's IC approval record adds its name and status. _steps = [(date, step, person), ...]."""
    out = {i: {} for i in ids if isinstance(i, str) and i}
    if not out:
        return pd.DataFrame(columns=APPROVAL_COLS)
    try:
        steps = _soql_in(inst, tok, "ProcessInstance.TargetObjectId,StepStatus,Comments,Actor.Name,CreatedDate",
                         "ProcessInstanceStep", "ProcessInstance.TargetObjectId", list(out))
    except Exception:
        steps = pd.DataFrame()
    if not steps.empty and "StepStatus" in steps:
        steps = steps[steps["StepStatus"].isin(list(STEP_NAMES))].sort_values("CreatedDate", kind="stable")
        for aid, g in steps.groupby("ProcessInstance.TargetObjectId"):
            rec = {"_steps": [(_d(r["CreatedDate"]), STEP_NAMES[r["StepStatus"]], _txt(r.get("Actor.Name")))
                              for _, r in g.iterrows()]}
            done = g[g["StepStatus"].isin(["Approved", "Rejected"])]
            if len(done):
                approvers = dict.fromkeys(_txt(x) for x in done.loc[done["StepStatus"] == "Approved", "Actor.Name"])
                rec.update({"Approval status": done.iloc[-1]["StepStatus"],
                            "Approved by": ", ".join(a for a in approvers if a) or None,
                            "Approved on": _d(done.iloc[-1].get("CreatedDate")),
                            "Approval comments": "\n".join(
                                f'{_mdy(_d(r["CreatedDate"]))}: {_txt(r.get("Actor.Name"))} : "{_txt(r.get("Comments"))}"'
                                for _, r in done.iterrows() if _txt(r.get("Comments"))) or None})
            out.setdefault(aid, {}).update(rec)
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
                # Without a real approver / approval-date field, "by" and "when" fall back to the record's last edit,
                # which is often an integration user — never let that stand in for the people who approved.
                found = {"Approval": _txt(r.get(pre + "Name")),
                         "Approval status": _txt(r.get(pre + src["status"])) if src["status"] else None,
                         "Approved by": _txt(r.get(pre + src["by"])) if src["by"] != "LastModifiedBy.Name" else None,
                         "Approved on": _d(r.get(pre + src["when"])) if src["when"] != "LastModifiedDate" else None,
                         "Approval comments": _txt(r.get(pre + src["comments"])) if src["comments"] else None}
                out.setdefault(r[key], {}).update({k: v for k, v in found.items() if v})
    except Exception:
        pass
    return pd.DataFrame.from_dict(out, orient="index").reindex(columns=APPROVAL_COLS)


# ───────────────────────────── delay reasons ────────────────────────────────
# What a draw's notes say held it up. Notes__c is a running log, one entry per line, mostly newest first:
# "AB 09/28 - Pending response on reductions", "9/30/2026 CD: report recd; funding initiated", "EF 09/30 approved".
#     note_entries(text) -> [(date or None, entry text), ...]   in the order written, header stripped
#     delay_reasons(text) -> [key, ...]                          DELAY_REASONS keys found, in display order
# DELAY_REASONS = {key: (label, side, compiled regex)}. Sides borrower, third_party, lender and admin are delay
# reasons; status_marker keys are progress markers, not reasons. Each pattern runs on one entry with its header
# removed, so staff initials and dates never fire on their own. Borrower documents count only while the notes show
# them outstanding (pending, needed, asked for, missing); third-party and internal steps count whenever they
# appear. Checked by hand against 189 real notes in Oct 2026: about 99% of tags right; on notes the rules had never
# seen (before the final fixes) 96% right, with 86% of reasons caught. Keep the patterns in step with that review.
# ---------------------------------------------------------------------------------------------------------
# Shared pieces
# ---------------------------------------------------------------------------------------------------------
# A word that says something is still outstanding, then a short window (same sentence) that may not cross a
# receipt / approval word or Land Gorilla / the inspector (so "Received lw", "Pending Approval - got LW" and
# "messaged LG for photos" do not count as borrower items).
_CUE = (r"(?<!'s\s)(?<!date\s)(?<!lg\s)(?<!inspector\s)\b(?:pending|need(?:s|ed)?|still\s+need|"
        r"still\s+waiting\s+(?:on|for)|waiting\s+(?:on|for)|wait\s+(?:on|for)|awaiting|follow(?:ed|ing)?[\s-]*up|f/u|"
        r"asked|asking|ask|request(?:ed|ing|s)?|req|(?<!borrower\s)(?<!borr\s)(?<!he\s)(?<!she\s)(?<!they\s)(?<!gc\s)sent|"
        r"send|emailed|missing|without|until\s+(?:we\s+)?(?:receive|get)|until|provid(?:e|es|ing)|expects?)\b")
_WIN = r"(?:(?!\b(?:received|receive|recd|rcvd|got|approv\w*|lg|inspector)\b)[^.;\n]){0,45}?"

_LW = r"(?:l/?ws?|lien[\s-]*waivers?|(?<!fee\s)(?<!tax\s)(?<!taxes\s)(?<!interest\s)(?<!penalty\s)waivers?(?!\s+of\s+(?:the\s+)?(?:fees?|inspection|requirements?)))"
_BACKUP = (r"(?:invoices?|pops?|proof(?:\s+of\s+payment)?|receipts|back[\s-]?up(?!\s+(?:generator|power|plan))|"
           r"back\s+of\s+(?:the\s+)?check|photos?|pictures?|pics|sow|scope(?:\s+of\s+work)?|"
           r"clarification(?![^.;\n]{0,50}?\b(?:lien|title|tax\w*|single\s*source|abstractor)\b)|"
           r"confirmation\s+(?:on|of|for)\b(?!\s+(?:the\s+)?(?:lien|title|(?:prop\w*\s+)?tax\w*|insurance|payment|inspection|wire|signed|documents?|docs?|loan))|"
           r"line\s+items?|soft[\s-]?costs?(?!\s+(?:draw|only)))")
_DOC_OUT = (r"(?:w-?9s?|wir(?:e|ing)\s+(?:instructions|info)|initial\s+docs?|"
            r"c\s?of\s?o(?!\s+(?:is\s+)?on\s+file)|certificates?\s+of\s+occupancy|"
            r"(?<!no\s)(?<!&\s)(?<!&)co(?![-'/\w])(?!\s+(?:is\s+)?on\s+file)(?!\s+\d{5})|bloc|letter\s+of\s+completion|"
            r"(?<!lw\s)(?<!waiver\s)e[\s-]?sig(?:s|natures?)?(?!\s+(?:for|on)\s+(?:the\s+)?(?:final\s+)?(?:l/?ws?|lien|waivers?))|docusign)")
# Optional entry header, so the line-start patterns also work on a raw note
_HDR = (r"(?:['\"*\-]*\s*(?:[a-z]{2}\s*-?\s*)?\d{1,2}[/.]\d{1,2}(?:/\d{2,4})?(?:\s*[a-z]{2}\b)?\s*[-\N{EN DASH}:.]*\s*)?"
        r"(?:(?!(?:co|lw|sc|gc|dq|ir|pf)\b)[a-z]{2}\s+(?=(?:approved|apv)\b))?")
_LINE_START = r"(?:^|(?<=\n))\s*"


def _any(*parts):
    return "|".join("(?:%s)" % p for p in parts)


# ---------------------------------------------------------------------------------------------------------
# Categories, in display order: (key, label, side, pattern)
# ---------------------------------------------------------------------------------------------------------
_CATEGORIES = [
    # ------------------------------------------------ borrower
    ("lien_waivers", "Lien waivers", "borrower", _any(
        _CUE + _WIN + r"\b" + _LW + r"\b",
        r"(?<!no\s)\b(?:l/?ws?|lien[\s-]*waivers?|waivers?)\s+(?:were\s+|was\s+)?(?:sent|requested|pending|outstanding|needed|missing)\b",
        r"\b(?:partial|remaining)\s+(?:l/?ws?|lien[\s-]*waivers?|waivers?)\b")),

    ("backup", "Invoices / backup", "borrower", _any(
        _CUE + _WIN + r"\b" + _BACKUP + r"\b(?!\s+(?:from|by|via|in)\s+(?:the\s+)?(?:lg|land\s+gorilla|inspector|inspection|report)\b)",
        r"(?<!provided\s)(?<!received\s)(?<!recd\s)(?<!rcvd\s)\bpops?\s+for\b(?!\s+plans\s+have)",
        r"\bpops?\s+(?:issues?|are\s+provided|provided)\b|\b(?:once|when)\s+(?:the\s+)?(?:borrower\s+)?(?:provides\s+)?pops?\b",
        r"\b(?:does\s+not|doesn'?t)\s+have\s+(?:the\s+)?(?:backup|invoices?|receipts?|pops?|photos?)\b",
        r"\bconfirmation\s+from\s+(?:the\s+)?borr\w*\s+that\b",
        r"\b(?:flooring|item|line\s+item|scope|material|finish(?:es)?)\s+confirmation\b",
        r"\breach\s+out\s+to\s+(?:the\s+)?borr\w*\s+about\s+(?:the\s+)?scope\b",
        r"\bhad\s+questions\b")),

    ("documents", "Missing documents", "borrower", _any(
        _CUE + _WIN + r"\b" + _DOC_OUT + r"\b",
        r"\bpermits?\b(?!\s+(?:is\s+|are\s+)?(?:(?:received|rcvd|recd)\s+(?:and\s+)?)?(?:on\s+file|not\s+(?:required|needed)))(?![\s-]+trigger)",
        r"(?<!for\s)(?<!business\s)(?<!exit\s)\b(?:stamped\s+|approved\s+)?plans\b(?!\s+(?:to|on|that|have\s+been|(?:and|&)\s+engineering)\b)|\bplan\s+set\b",
        r"\b(?:city|county|municipal)\s+(?:final\s+)?(?:inspections?|sign[\s-]?off)\b",
        r"\bgc\s+(?:lic(?:en[cs]e)?|info|contract)\b|\bcontractor'?s?\s+(?:list|lic(?:en[cs]e)?|info)\b|\blicen[cs]e\s+info\w*",
        r"\bheavy\s+rehab\s+info\b",
        r"\b(?:c\s?of\s?o|co)\s+(?:is\s+)?(?:still\s+)?(?:pending|eta|outstanding|not\s+yet)\b")),

    ("borrower_response", "Borrower response", "borrower", _any(
        r"\bno\s+response\b(?!\s+(?:yet\s+)?from\s+(?:the\s+)?(?:lg|land\s+gorilla|inspector|title|single\s*source|city|county|permit\w*|servicer|fci|onity|appraiser|insurance|agent|bank))",
        r"(?<!lg\s)(?<!inspector\s)\bhas\s+not\s+(?:been\s+responsive|responded)\b",
        r"\b(?:borrower|borr|gc)\s+(?:is\s+|was\s+)?(?:un|non[\s-]?)responsive\b",
        r"\basked\b(?!\s+(?:the\s+)?(?:lg|land\s+gorilla|inspector|title|single\s*source)\b)[^.;\n]{0,40}?\bagain\b",
        r"\bsent\s+(?:a\s+)?reminders?\b(?!\s+to\s+(?:the\s+)?(?:lg|inspector|title))",
        r"\bwill\s+(?:cancel|delete)\b",
        r"\b(?:pending|awaiting|waiting\s+(?:on|for))\s+(?:the\s+)?borr\w*'?s?\s+(?:response|reply|answer|decision|confirmation|on|re|regarding|about)\b",
        r"(?<!lg\s)(?<!inspector\s)\bresponse\s+on\b",
        r"\bfollow(?:ed|ing)?[\s-]*up\s+(?:\w+\s+){0,2}?(?:with|to)\s+(?:the\s+)?borr\w*",
        r"(?<!lg\sto\s)\bfollow(?:ed|ing)?[\s-]*up\s+on\s+(?!(?:the\s+)?(?:report|insp\w*|title|extension|lg|order|site\s+visit|scheduling)\b)\w",
        r"\b(?:waiting|wait)\s+(?:on|for)\s+[^.;\n]{0,40}?\bfrom\s+(?:the\s+)?borr\w*",
        r"\bhave\s+not\s+received\s+(?:an?\s+)?update\b",
        r"\b(?:waiting|wait)\s+(?:on|for)\s+(?:the\s+)?borr\w*\s+to\b",
        r"\b(?:emailed|called|texted|messaged|sent)\b(?:(?!\b(?:lg|land\s+gorilla|inspector)\b)[^.;\n]){0,25}?\bremind\w*\b"
        r"(?!\s+to\s+(?:the\s+)?(?:lg|land\s+gorilla|inspector|title|single\s*source)\b)",
        r"\b(?:requested|requesting|asked\s+for)\s+(?:an?\s+)?(?:update|status)\s+from\s+(?:the\s+)?borr\w*",
        r"\bsent\s+(?:another|an?\s+additional|a\s+second|2nd|(?:a\s+)?follow[\s-]*up)\s+(?:email|e-mail|message|request|vm|voice\s*mail)\b"
        r"(?![^.;\n]{0,25}\b(?:lg|inspector|title|single\s*source)\b)",
        r"\b(?:l/?ws?|lien\s+waivers?|soft[\s-]?costs?|pops?|invoices?|photos?|pics|bloc|w-?9s?|wire\s+info|eoi|insurance|"
        r"back[\s-]?up|self[\s-]*insp\w*|guc\s+info)\b[^.;\n]{0,15}?-\s*\d{1,2}/\d{1,2}\s+(?:[a-z]{2}\s+)?followed\s+up\b")),

    ("self_inspection", "Self-inspection / site access", "borrower", _any(
        r"\bself[\s-]*(?:insp\w*|inps\w*)",
        r"\bremote\s+inspection\b",
        r"\bno\s+access\b(?!\s+to\s+(?:the\s+)?(?:portal|system|box|folder|lg|email))",
        r"\b(?:could\s*n[o']?t|unable\s+to|not\s+able\s+to|can'?t|cannot|did\s*n[o']?t)\s+(?:get\s+(?:in(?:to)?\b|access)|access|enter)",
        r"\b(?:site|property|house|home|unit|gate)\s+(?:was\s+|is\s+)?locked\b",
        r"\b(?:borrower|borr|gc|contractor|poc)\s+(?:has\s+|is\s+|was\s+)?(?:moved|rescheduled|requested|requesting|scheduled|cancell?ed|postponed|delayed|pushed)\b(?:\s+that)?\s+(?:the\s+)?(?:\w+\s+)?inspection",
        r"\b(?:is\s*n[o']?t|is\s+not|was\s*n[o']?t|not)\s+ready\s+for\s+(?:the\s+)?inspection\b",
        r"\b(?:asked|wants|wanted)\s+to\s+(?:hold|delay|push|move|reschedule|postpone)\s+(?:the\s+)?inspection",
        r"\bin\s+lieu\s+of\s+(?:an?\s+)?inspection\b")),

    ("insurance", "Insurance", "borrower", _any(
        r"\b(?:insurance|policy|eoi)\b[^.;\n]{0,20}?\b(?:expired|lapsed|cancell?ed|was\s+extended)\b",
        r"\b(?:updated?|new|renewed)\s+(?:insurance|policy|eoi|dec\s+page)\b",
        r"\b(?:need|needs|pending|waiting\s+(?:on|for)|awaiting|requested|asked|asking|req|reminders?\s+for|expects?|follow(?:ed)?[\s-]*up\s+(?:on|for))\b"
        r"[^.;\n]{0,30}?\b(?:insurance|eoi|policy)\b(?!\s+(?:claim|agent|current|exp))",
        r"\bextension\s+of\s+(?:the\s+)?(?:insurance\s+)?policy\b",
        r"\beoi\b(?!\s+(?:is\s+)?(?:on\s+file|current|good|received|rcvd|recd|in\s+(?:box|sf)))|\bevidence\s+of\s+insurance\b|\bbuilder'?s'?\s+risk\b")),

    ("taxes_payments", "Taxes / payments due", "borrower", _any(
        r"\btax(?:es)?\b[^.;\n/]{0,25}?\b(?<!not\s)(?<!not\syet\s)(?<!no\slonger\s)(?<!never\s)(?:delinq|deliq)\w*\b(?!\s+until)",
        r"\b(?:dq|delinquent)\s+(?:property\s+)?tax",
        r"\b(?:prop(?:erty)?\s+)?tax\s+(?:pmt|payment)",
        r"\bconfirmation\s+that\s+(?:the\s+)?(?:property\s+)?taxes\b",
        r"\btaxes\s+(?:are\s+)?due\b(?!\s+(?:on\s+|in\s+|by\s+)?(?:\d|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec))(?![^.;\n]*\bnot\s+(?:yet\s+)?delinquent)",
        r"\bwaiv\w*\s+(?:the\s+)?(?:property\s+)?taxes\b|\bbring\s+(?:the\s+)?(?:property\s+)?taxes\s+current\b",
        r"\b(?:fci|onity|servicer)\s+(?:\w+\s+)?(?:payment|pmt)\b(?!s?\s+(?:due|made\s+internally))",
        r"\b(?:pending|waiting\s+on|awaiting|need|due\s+for|missed|late)\s+(?:the\s+)?(?:(?:loan|mortgage|monthly|interest|"
        r"january|february|march|april|may|june|july|august|september|october|november|december|sept|oct|nov|dec|aug|"
        r"\d{1,2}/\d{1,2})\s+){1,2}(?:payment|pmt)s?\b",
        r"\b(?:loan|mortgage|monthly|\d{1,2}/\d{1,2})\s+(?:payment|pmt)\s+(?:was\s+|has\s+been\s+)?(?:made|rec\w*)\b",
        r"\b(?:loan|mortgage)\s+(?:payment|pmt)\b(?!\s+(?:is\s+)?(?:due|amount))",
        r"\bdue\s+for\s+(?:the\s+)?\d{1,2}/\d{1,2}\b(?![^.;\n]{0,30}\b(?:has\s+(?:an?\s+)?ir|interest\s+reserve|autopay)\b)",
        r"\b(?:pending|need(?:s|ed)?|still\s+(?:pending|due)|reminders?\s+for|follow(?:ed)?[\s-]*up\s+on|waiting\s+on|awaiting)\b"
        r"[^.;\n]{0,40}?\b(?:\d{1,2}/\d{1,2}/?(?:\d{2,4})?|loan|mortgage|monthly)\s+(?:payment|pmt)s?\b",
        r"\b(?:payment|pmt)s?\s+(?:must|needs?\s+to|has\s+to)\s+be\s+(?:brought\s+)?current\b",
        r"\bpast\s+due\b|\bnow\s+current\b")),

    ("budget_changes", "Budget changes", "borrower", _any(
        r"(?<!no\s)(?<!no\snew\s)(?<!without\s)(?<!without\sa\s)\bchange\s*order(?:s|\s*#?\d+)?\b",
        r"\bre-?alloc\w*",
        r"\bbudget\s+(?:revision|change|modification|increase|reallocation|amendment|shortfall|overrun)s?\b",
        r"\bcontingency\b|\bcost\s+overruns?\b|\bover[\s-]budget\b",
        r"\badditional\s+\$\s?\d",
        r"\bunexpected\s+(?:costs?|expenses?)\b|\bcosts?\s+over\s+\$",
        r"\b(?:changed|changing|change\s+(?:in|to)|revised|revising)\s+(?:the\s+)?(?:business\s+plan|scope(?:\s+of\s+work)?)\b|\bscope\s+change\b")),

    # ------------------------------------------------ admin (loan-level: both sides work on it)
    ("extension_maturity", "Extension / maturity / payoff", "admin", _any(
        r"(?<!bedroom\s)(?<!bedroom-)(?<!room\s)(?<!kitchen\s)(?<!rear\s)(?<!garage\s)(?<!house\s)(?<!home\s)(?<!story\s)(?<!deck\s)(?<!building\s)"
        r"\bextension\b(?!\s+of\s+(?:the\s+)?(?:insurance\s+)?polic)(?!\s+or\s+(?:new\s+)?(?:bath|bed|room|addition|kitchen))",
        r"\bmaturity\b|\boutside\s+(?:advance|complete|completion)\s+date\b",
        r"\bloan\s+mod(?:ification)?\b|\b(?:extension|loan)\s*/\s*mod\w*",
        r"(?<!lien\s)(?<!mortgage\s)(?<!judgment\s)(?<!judgement\s)\bpay[\s-]?offs?\b(?!\s+(?:of\s+)?(?:the\s+)?(?:\w+\s+)?(?:lien|mortgage|judge?ment)s?\b)|\brefinanc\w*|\brefi\b")),

    # ------------------------------------------------ third party
    ("inspection", "Inspection", "third_party", _any(
        r"(?<!self-)(?<!self\s)(?<!remote\s)(?<!no\s)(?<!lieu\sof\s)(?<!permit\s)(?<!city\s)(?<!county\s)"
        r"(?<!self-\s)(?<!without\sa\sfinal\s)(?<!last\s)(?<!previous\s)(?<!prior\s)(?<!city\sfinal\s)(?<!county\sfinal\s)\binsp\w*\b(?!\s+fees?\b)",
        r"\bqc\b|\bqaqc\b",
        r"(?<!title\s)\breport\s+(?:rec\w*|rcvd|in\b|expected|pending|delivered|finally)",
        r"\bin\s+review\b(?=[^.;\n]*\blg\b)|\bsite\s+visit\b")),

    ("inspection_delay", "Inspection delay / revision", "third_party", _any(
        r"\binsp\w*\s+(?:scheduling\s+)?(?:delays?|issues?|problems?|revisions?)\b|\bscheduling\s+delays?\b",
        r"\b(?:lg|land\s+gorilla|inspector)\s+delays?\b|\bdelays?\s+(?:from|with|at|by)\s+(?:the\s+)?(?:lg|land\s+gorilla|inspector)\b",
        r"\bre-?inspect\w*|\bnew\s+inspection\b|\bsecond\s+inspection\b|\bgo\s+back\s+out\b",
        r"\brevised\s+report\b|(?<!title\s)\breport\s+(?:update|revision)s?\b",
        r"\b(?:needs?|needing|pending|wait(?:ing)?\s+(?:on|for))\s+(?:(?:inspection|report|the)\s+)?revisions?\b",
        r"\brevisions?\s+(?:from|by)\s+(?:the\s+)?inspector\b",
        r"\binspector\s+(?:was\s+|is\s+)?(?:non[\s-]?responsive|unresponsive)\b",
        r"\binspector\s+(?:didn'?t|did\s+not|has\s+not|hasn'?t)\s+(?:get|take|include|send|upload)\b",
        r"\b(?:pics|photos|pictures)\b[^.;\n]{0,40}?\bnot\s+included\b|\bnot\s+included\s+in\s+the\s+report\b",
        r"\bno\s+update\s+on\s+(?:the\s+)?(?:inspection|report|order)\b|\bnot\s+(?:yet\s+)?scheduled\b",
        r"\b(?:follow(?:ed|ing)?[\s-]*up|f/u|reached\s+out|inquired|checked\s+in)\s+(?:with|to)\s+(?:the\s+)?(?:lg|land\s+gorilla|inspector)\b",
        r"\b(?:status|update)\s+from\s+(?:lg|land\s+gorilla)\b|\b(?:messaged|emailed|sent)\s+(?:the\s+)?lg\b|\blg\s+followed\s+up\b",
        r"\b(?:sent|emailed|messaged)\s+(?:an?\s+)?reminders?\s+to\s+(?:the\s+)?(?:lg|land\s+gorilla|inspector)\b|"
        r"\breminded\s+(?:the\s+)?(?:lg|land\s+gorilla|inspector)\b",
        r"\bsent\s+(?:an?\s+)?(?:email|message)\b[^.;\n]{0,40}?\b(?:report|inspection)\b|\bfollowing\s+up\s+on\s+(?:the\s+)?report\b",
        r"\breason\s+for\s+(?:the\s+)?delay\b|\bconfirmed\s+(?:inspection|site\s+visit)\s+date\b",
        r"\bholiday\b[^.;\n]{0,60}?\b(?:report|inspection)\b|\binformation\s+from\s+the\s+inspector\b",
        r"\b(?:unable\s+to|could\s*n[o']?t|not)\s+(?:yet\s+)?(?:locate[d]?|find|found|assign(?:ed)?)\s+(?:an?\s+)?inspector\b")),

    # a line that only sets conditions for later fundings ("no more fundings without ... title search") is skipped
    ("title", "Title", "third_party",
        _LINE_START + r"(?![^\n]*\bno\s+(?:more|further)\s+fund\w*\s+without\b)[^\n]*?"
        r"(?:(?<!job\s)\btitle\b(?!\s*:\s*(?:\w+\s+)?(?:officer|manager|director|analyst|processor|closer|coordinator|"
        r"specialist|vp|president)\b)|\bsingle\s*source\b|\babstractor\b)"),

    ("liens", "Liens / title issues", "third_party", _any(
        r"(?<!no\s)(?<!no\sopen\s)(?<!no\sother\s)(?<!no\sknown\s)(?<!no\snew\s)\bliens?\b(?![\s-]*(?:waivers?|free))",
        r"\bmortgagee\b|\bnot\s+(?:yet\s+)?(?:been\s+)?recorded\b|\bbe\s+recorded\b",
        r"\b(?:second|2nd|junior)\s+(?:mortgage|lien|deed)s?\b",
        r"\bjudge?ments?\b(?!\s+call)|\blis\s+pendens\b|\blitigation\b|\blawsuits?\b",
        r"\bcannot\s+complete\s+the\s+title\b|\btitle\s+(?:issue|problem|exception|defect)s?\b")),

    ("capital_partner", "Capital partner", "third_party", _any(
        r"\bcap(?:ital)?\.?\s+partners?\b",
        r"\bcp\s+(?:release|approval|funding)\b|\bpartner\s+release\b",
        r"\brejected\s+by\s+(?:the\s+)?(?:capital\s+)?partner\b")),

    # ------------------------------------------------ lender
    ("internal_approval", "Internal approval", "lender", _any(
        r"\bsecondary\s+review\b|\b(?:mgmt|management|manager|senior|credit\s+committee)\s+(?:review|approval)\b",
        r"\breview(?:ed)?\s+by\s+(?:mgmt|management)\b|\bfunding\s+approval\b",
        r"\b(?:response|decision|review|approval|sign[\s-]?off)\s+(?:by|from)\s+(?:mgmt|management|manager|credit)\b",
        r"\bpending\s+(?!(?-i:LG|LW|CO|IC|QC|GC|SC|PF|DQ|IR|OK|IT|CV|AM|PM|ID|NO|RE|FW|TO|IN|ON|AT|BY|UP|CP)\b)"
        r"(?-i:[A-Z]{2})\s+(?:review|approval)\b",
        r"\bset\s+up\s+a\s+(?:conference\s+)?call\b",
        r"\bhow\s+to\s+proceed\b|(?<!title\s)\bexceptions?\b",
        r"\b(?:approv\w*|fund\w*|process\w*|releas\w*)\b[^.;\n]{0,40}?\b(?:despite|knowing\s+that|even\s+though)\b",
        r"\bwaiv(?:e|ed|ing)\b(?!\s+(?:the\s+)?(?:inspection\s+)?fees?\b)",
        r"\bapproved\b[^.;\n]{0,40}?\bafter\s+talking\s+(?:to|with)\b",
        r"\bescalat\w*\s+to\s+(?:mgmt|management|manager|credit)\b|\bper\s+(?:mgmt|management)\b")),

    ("processing", "Processing / wire timing", "lender", _any(
        r"\bcut[\s-]?off\b|\bmissed\s+the\s+wire\b|\bwire\s+deadline\b|\b(?:after|past)\s+(?:the\s+)?(?:wire\s+)?deadline\b",
        r"\b(?:approval|sign[\s-]?off)\s+from\s+accounting\b|\baccounting\s+(?:approval|review)\b|\bpending\s+accounting\b",
        r"\b(?:cancel+ed|recalled|returned|rejected|bounced)\s+(?:the\s+)?wires?\b|\bwires?\s+(?:was\s+|were\s+)?(?:cancel+ed|recalled|returned|rejected|bounced)\b|"
        r"\bmanual(?:ly)?\b[^.;\n]{0,25}?\b(?:wd|workday)\b",
        r"\bvendor\s+(?:setup|set[\s-]?up|account)\b|\bpayee\s+(?:setup|set[\s-]?up)\b|\bbank\s+account\s+change\b",
        r"\bwiring\s+info\s+approved\b|\bunlock\w*",
        r"\byardi\b|\bpending\s+workday\b",
        r"\bno\s+one\s+can\s+(?:currently\s+)?approve\b|\brecreat\w*",
        r"\b(?:advance|draw|request)\s+(?:\w+\s+)?(?:was\s+)?(?:somehow\s+)?deleted\b",
        r"\bper\s+ticket\b|\b(?-i:IT)\s+(?:has\s+change\w*|issues?|outage|ticket)\b",
        r"\bholiday\b[^.;\n]{0,40}?\b(?:funding|wire)\b|\boffice\s+clos\w+")),

    # ------------------------------------------------ status markers (milestones, not reasons)
    ("sent_for_approval", "Sent for approval", "status_marker", _any(
        r"\bsubmit(?:ted|ting)?\s+(?:[\w/]+\s+){0,3}?for\s+(?:final\s+)?approval\b",
        r"\bpending\s+approval\b|\bsubmitted\s+to\s+ic\b",
        r"\bfull\s+(?:draw\s+)?(?:pkg|package)\s+(?:rec\w*|rcvd|in|complete)\b")),

    ("approved", "Approved", "status_marker", _any(
        _LINE_START + _HDR + r"(?:approved|apv)\b(?!\s+to\s+(?:switch|change|use|order|add)\b)(?!\s+(?:an?\s+)?(?:additional|extra|change\s+order|co\b|budget|realloc\w*)(?![^.;\n]{0,25}\b(?:and|&)\s+(?:the\s+)?draw\b))",
        r"(?<!then\s)\b(?:advance|draw)\s+(?:was\s+|is\s+|has\s+been\s+)?approved\b",
        r"\bapprov(?:ed|ing)\s+(?:this\s+)?(?:wire|draw)\b(?!\s+amount)|\bapproved\s+and\s+funded\b",
        r"(?<!not\s)(?<!be\s)\bapproved\s+(?:for\s+(?:funding|draw|release)|to\s+(?:fund|release|wire)|as)\b",
        r"\bdraw\s+reviewed\s*(?:;|,|and)?\s*approved\b|[.;]\s+approved\s+(?:by|per|for)\b",
        r"\bfunding\s+(?:to\s+be\s+|was\s+|is\s+)?initiated\b|\bsubmitting\s+(?:the\s+)?wire\b|\bpending\s+release\b")),

    ("unsubmitted", "Sent back to Unsubmitted", "status_marker", _any(
        r"(?<!go\sinto\s)(?<!will\sbe\sput\sin\s)(?<!will\sput\sin\s)(?<!will\sput\sit\sin\s)(?<!will\sgo\sto\s)(?<!will\smove\sto\s)(?<!will\sbe\smoved\sto\s)\bun-?s(?:ub|bu)mitted\b|\bunsub\b",
        r"\bnot\s+submitted\b")),

    ("final_draw", "Final draw", "status_marker", _any(
        r"(?<!for\s)(?<!if\sthis\sis\sthe\s)(?<!on\sthe\s)(?<!supposed\sto\sbe\s)(?<!not\s)(?<!not\sa\s)(?<!not\sthe\s)(?<!next\sdraw\swill\sbe\sthe\s)(?<!next\sone\swill\sbe\sthe\s)(?<!next\sdraw\sis\sthe\s)"
        r"\bfinal\s+draw\b(?!\s+with\s+no)",
        r"\bnow\s+final\b|\bdraw\s+is\s+(?:now\s+)?final\b",
        r"\bproject\s+is\s+(?:100%\s+)?complete\b",
        r"(?<!without\sa\s)\bfinal\s+inspection\b",
        r"\bfinal\s+(?:l/?ws?|lien\s+waivers?|waivers?|pay\s*app\w*)\b")),
]

DELAY_REASONS = {key: (label, side, re.compile(rx, re.IGNORECASE)) for key, label, side, rx in _CATEGORIES}
REASON_SIDES = ("borrower", "third_party", "lender", "admin")

# ---------------------------------------------------------------------------------------------------------
# Entry header: initials + date in the orders staff write them. Case-sensitive on purpose: initials AFTER a date
# must be upper case so "10/01 payment due" is not read as initials "pa". Dot dates ("AB 08.13") only after
# initials, so "8.5% complete" is never a date.
#   "AB 09/28 - text"  "Ab 09/28 -"  "CD09/22 approved"  "EF- 9/2- text"  "AB - 09/02 - text"  "GH 08.13 Approved"
#   "9/30/2026 CD: text"  "9/15 EF- text"  "09/22 Approved"  "10/01 payment due"  "GH approved 09/28"
# ---------------------------------------------------------------------------------------------------------
NOTE_ENTRY = re.compile(
    r"^\s*['\"*\-]*\s*(?:"
    r"(?P<ini>[A-Za-z]{2})\s*-?\s*(?P<date>\d{1,2}(?:/\d{1,2}(?:/\d{2,4})?|\.\d{2}))(?![\d/%]|\.\d)"
    r"|(?P<date2>\d{1,2}/\d{1,2}(?:/\d{2,4})?)(?![\d/%]|\.\d)(?:\s*(?P<ini2>[A-Z]{2})\b)?"
    r"|(?P<ini3>[A-Za-z]{2})\s+(?P<verb>[Aa]pproved|APPROVED|[Aa]pv)\s+(?P<date3>\d{1,2}/\d{1,2}(?:/\d{2,4})?)(?![\d/%]|\.\d)"
    r")\s*[-\N{EN DASH}\N{EM DASH}:.]*\s*(?P<text>.*)$")
_NOT_INITIALS = {"LG", "LW", "CO", "IC", "QC", "GC", "SC", "PF", "DQ", "IR", "OK", "IT", "CV", "AM", "PM", "ID",
                 "NO", "RE", "FW", "TO", "IN", "ON", "AT", "BY", "UP"}
_APPROVED_FIRST = re.compile(r"(?:approved|apv)\b", re.IGNORECASE)   # "CO 09/14 - approved": CO = initials
_DUE_LINE = re.compile(r"^\s*\d{1,2}/\d{1,2}(?:/\d{2,4})?\s+(?:payment|pmt)s?\s+(?:now\s+)?due\b", re.IGNORECASE)


def _valid(date):
    parts = re.split(r"[/.]", date)
    try:
        m, d = int(parts[0]), int(parts[1])
    except (ValueError, IndexError):
        return False
    return 1 <= m <= 12 and 1 <= d <= 31


def note_entries(text):
    """Split a note log into [(date or None, entry text), ...] in the order written (newest first, as staff
    write them). The date is the entry header's date as written, with '.' turned into '/' ("09/28",
    "9/30/2026"); the text has the initials/date header removed (a header word such as LG or CO is kept in the text,
    except before "approved", where it is the approver's initials). 'MM/DD payment due' lines give the next loan
    payment date, not an entry date, so they come back as (None, whole line). Lines starting with '/' continue
    the entry above. Excel's '_x000D_' escapes count as line breaks."""
    if text is None or not isinstance(text, str):
        return []
    out = []
    for raw in re.split(r"[\r\n]+", re.sub(r"_x000[dD]_", "\n", text)):
        line = " ".join(raw.split())          # one space between words: long blank runs made some patterns crawl
        if not line:
            continue
        if line.startswith("/") and out:
            d, t = out[-1]
            out[-1] = (d, (t + " " + line.lstrip("/ ").strip()).strip())
            continue
        if _DUE_LINE.match(line):
            out.append((None, line))
            continue
        m = NOTE_ENTRY.match(line)
        if not m:
            out.append((None, line))
            continue
        ini = m.group("ini") or m.group("ini2") or m.group("ini3")
        date = m.group("date") or m.group("date2") or m.group("date3")
        body = m.group("text").strip()
        if m.group("verb"):
            body = (m.group("verb") + " " + body).strip()
        if ini and ini.upper() in _NOT_INITIALS and (m.group("ini3") or not _APPROVED_FIRST.match(body)):
            body = (ini + " " + body).strip()         # "LG 9/14 ..." is not a person: keep the word
        date = date.replace(".", "/") if _valid(date) else None
        out.append((date, body))
    return out


# A self / remote inspection named in the note ("elf inspection" is a typo seen in the notes) ...
_SELF_INSP = re.compile(r"\bs?elf[\s-]*(?:insp|inps)\w*|\bremote\s+inspection\b|\bin\s+lieu\s+of\s+(?:an?\s+)?inspection\b",
                        re.IGNORECASE)
# ... and signs of a third-party inspection: LG, the inspector, QC, a report, a site visit, 3rd party, re-inspection,
# or an inspection ordered / scheduled in an entry that is not about the self-inspection.
_THIRD_INSP = re.compile(r"\b(?:lg|land\s+gorilla|inspectors?|qc|qaqc|report|site\s+visit|on[\s-]?site|third[\s-]*party|"
                         r"3rd[\s-]*party|partner|re-?inspect\w*)\b", re.IGNORECASE)
_ORDERED = re.compile(r"\b(?:ordered|scheduled|sched)\b", re.IGNORECASE)


def _self_inspection_only(entries):
    """True when the note names a self / remote inspection and nothing in it shows a third-party inspection."""
    if not any(_SELF_INSP.search(e) for e in entries):
        return False
    return not any(_THIRD_INSP.search(e) or (_ORDERED.search(e) and not _SELF_INSP.search(e)) for e in entries)


def delay_reasons(text):
    """DELAY_REASONS keys found in a note (or any free text), in DELAY_REASONS order, each once.
    Includes status markers; keep side != 'status_marker' for 'held up by' reasons.
    One note-level rule: when the note names a self / remote inspection and shows no third-party inspection
    (_self_inspection_only), a bare "inspection" there is the borrower's self-inspection, so 'inspection' (third
    party) is dropped and 'self_inspection' (borrower) kept."""
    found = set()
    entries = [entry for _date, entry in note_entries(text)]
    for entry in entries:
        for key, (_label, _side, rx) in DELAY_REASONS.items():
            if key not in found and rx.search(entry):
                found.add(key)
    if "inspection" in found and "self_inspection" in found and _self_inspection_only(entries):
        found.discard("inspection")
    return [key for key in DELAY_REASONS if key in found]


# How the app shows them: progress markers aren't reasons, and a routine inspection (about half of all draws have
# one) is listed on each draw but not counted as an account's usual hold-up.
WAITING_ON = {"borrower": "Borrower", "third_party": "Third parties", "lender": "CoreVest", "admin": "Loan changes"}
ROUTINE_REASONS = ("inspection",)


@st.cache_resource(show_spinner=False)
def _reason_memo(rules_version: str) -> dict:
    """Process-wide note -> reasons memo. It outlives reruns (the original deployment re-runs this whole file each
       time), and a change to the patterns starts a fresh one."""
    return {}


_RULES_VERSION = hashlib.sha1("\n".join(rx for _k, _l, _s, rx in _CATEGORIES).encode()).hexdigest()
_MEMO: dict | None = None                     # the memo above, fetched once: looking it up per note is slow


def delay_keys(note) -> tuple:
    """DELAY_REASONS keys for what a draw's notes say held it up, in display order (no progress markers)."""
    global _MEMO
    t = _txt(note)
    if not t:
        return ()
    if _MEMO is None:
        _MEMO = _reason_memo(_RULES_VERSION)
    memo = _MEMO
    hit = memo.get(t)
    if hit is None:
        hit = tuple(k for k in delay_reasons(t) if DELAY_REASONS[k][1] in REASON_SIDES)
        if len(memo) > 20000:
            memo.clear()
        memo[t] = hit
    return hit


def delay_labels(note) -> list:
    """What a draw's notes say held it up, in plain English."""
    return [DELAY_REASONS[k][0] for k in delay_keys(note)]


def delay_evidence(note, limit: int = 3) -> list:
    """The note entries behind those reasons, in the order written (usually newest first): '09/28: Pending LW'."""
    keys = delay_keys(note)
    if not keys:
        return []
    out = []
    for when, entry in note_entries(_txt(note)):
        if any(DELAY_REASONS[k][2].search(entry) for k in keys):
            out.append((f"{when}: " if when else "") + (entry if len(entry) <= 160 else entry[:157] + "…"))
            if len(out) == limit:
                break
    return out


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


def _md(v) -> str:
    """Salesforce text made safe to drop into markdown: no stray bold, links or math."""
    return re.sub(r"([\\`*_{}\[\]()#+!|~<>$])", r"\\\1", _txt(v))


def _show(v, sftype: str | None = None) -> str | None:
    """A Salesforce / Land Gorilla value as display text: dates m/d/yyyy, money $1,234, percents 45% (None if blank)."""
    if isinstance(v, (list, tuple)):
        return ", ".join(s for s in (_show(x) for x in v) if s) or None
    t = _txt(v)
    if not t:
        return None
    if sftype in ("date", "datetime") or isinstance(v, date):
        d = _d(v)
        return _mdy(d) if d else t
    try:
        if sftype == "currency":
            return f"${float(v):,.0f}"
        if sftype == "percent":
            return f"{float(v):,.1f}".rstrip("0").rstrip(".") + "%"
    except (TypeError, ValueError):
        return t
    if sftype == "boolean":
        return "Yes" if t.lower() in ("true", "1", "yes", "y") else "No"
    return t


def _draw_no(r) -> str:
    """A draw's number within its loan: '11' for advance 34961-11 (whose Salesforce Name is 11)."""
    name, adv = _txt(r.get("Name")), _txt(r.get("Loan_Advance_Number__c"))
    if name.isdigit():
        return name
    m = re.search(r"-(\d+)\s*$", adv)
    return m.group(1) if m else (adv or name or "?")


def _draw_nos(df: pd.DataFrame) -> list:
    return [_draw_no({"Name": n, "Loan_Advance_Number__c": a})
            for n, a in zip(_col(df, "Name"), _col(df, "Loan_Advance_Number__c"))]


def _day(s) -> pd.Series:
    return pd.to_datetime(s, errors="coerce").astype("datetime64[ns]")


def _created_day(df: pd.DataFrame) -> pd.Series:
    """Salesforce's CreatedDate as an Eastern-time date."""
    t = pd.to_datetime(_col(df, "CreatedDate"), errors="coerce", utc=True)
    return t.dt.tz_convert("America/New_York").dt.tz_localize(None).dt.normalize().astype("datetime64[ns]")


def _wire_day(df: pd.DataFrame) -> pd.Series:
    return _day(_col(df, "_wire")).fillna(_day(_col(df, WIRE_FIELD)))


def _request_date(df: pd.DataFrame) -> pd.Series:
    """Each draw's request date: Land Gorilla's; else the earlier of Salesforce's request date (often just the wire
       date copied over) and the day the record was created — unless that came after the wire (a bulk load)."""
    made, wire = _created_day(df), _wire_day(df)
    sf = pd.concat([_day(_col(df, REQ_FIELD)), made.where(wire.isna() | (made <= wire))], axis=1).min(axis=1)
    return _day(_col(df, "_lg_req")).fillna(sf.astype("datetime64[ns]"))


def _order_date(df: pd.DataFrame) -> pd.Series:
    """The date a draw sorts by: its request date, else its wire date (older draws often have nothing else), else the
       day it was created in Salesforce."""
    return _request_date(df).fillna(_wire_day(df)).fillna(_created_day(df))


def _waiting_summary(keys_per_draw) -> tuple:
    """What a set of draws was usually waiting on: (top hold-ups, who they were mostly on, hold-up counts). A draw
       counts once per hold-up and once per side, so one long note can't outvote several draws; tied sides are all
       named; ties between hold-ups go alphabetically. Routine inspections aren't counted."""
    keys = [[k for k in ks if k not in ROUTINE_REASONS] for ks in keys_per_draw]
    top = pd.Series([DELAY_REASONS[k][0] for ks in keys for k in ks], dtype=object).value_counts()
    top = top.sort_index().sort_values(ascending=False, kind="stable")
    sides = pd.Series([s for ks in keys for s in {DELAY_REASONS[k][1] for k in ks}], dtype=object).value_counts()
    who = None
    if len(sides):
        best = set(sides[sides == sides.max()].index)
        who = " & ".join(WAITING_ON[s] for s in WAITING_ON if s in best)
    return list(top.index[:3]), who, top


def _newest_first(df: pd.DataFrame) -> pd.DataFrame:
    """Most recent request on top. A loan's draws stay together, highest draw number first (draws are numbered in
       the order they're requested); loans are ordered by their most recent draw."""
    if df.empty:
        return df
    when = _order_date(df)
    seq = pd.to_numeric(pd.Series(_draw_nos(df), index=df.index), errors="coerce")
    loan = _col(df, "Loan_Number__c").map(_txt)
    latest = when.groupby(loan).transform("max")
    return (df.assign(_when=when.values, _seq=seq.values, _latest=latest.values, _loan=loan.values)
              .sort_values(["_latest", "_loan", "_seq", "_when"], ascending=[False, True, False, False],
                           na_position="last", kind="stable")
              .drop(columns=["_latest", "_loan"]))


def _by_date(df: pd.DataFrame) -> pd.DataFrame:
    """Draws from many loans, most recent request first (ties: higher draw number first)."""
    seq = pd.to_numeric(pd.Series(_draw_nos(df), index=df.index), errors="coerce")
    return (df.assign(_when=_order_date(df).values, _seq=seq.values)
              .sort_values(["_when", "_seq"], ascending=False, na_position="last", kind="stable"))


def latest_entry(note, ref: date | None = None) -> tuple | None:
    """The newest dated entry in a note log, as (date, text). Logs run newest-first or oldest-first, and entry dates
       usually have no year, so each goes in the year that puts it closest to ref (the draw's own date)."""
    ref, best = ref or _today(), None
    for when, text in note_entries(_txt(note)):
        text = text.lstrip(";:,-–— ").strip()
        if not when or not text:
            continue
        p = [int(x) for x in when.split("/")]
        try:
            d = (date(p[2] + (2000 if p[2] < 100 else 0), p[0], p[1]) if len(p) == 3 else
                 min((date(y, p[0], p[1]) for y in (ref.year - 1, ref.year, ref.year + 1)),
                     key=lambda x: abs((x - ref).days)))
        except ValueError:
            continue
        if best is None or d > best[0]:
            best = (d, text)
    return best


def _timeline(row: pd.Series, a) -> pd.DataFrame:
    """What happened to a draw, in date order — only steps that were recorded: Land Gorilla's request and funding,
       the Salesforce milestones that were filled in, the full-package date, and each approval with who gave it."""
    req, wire = _d(row.get("_lg_req")), _d(row.get("_wire")) or _d(row.get(WIRE_FIELD))
    sf_req = _d(row.get(REQ_FIELD))
    if not req and sf_req and (not wire or sf_req < wire):        # Salesforce's own request date, unless it's just
        req = sf_req                                               # the wire date copied over
    ev = [(req, "Requested", _txt(row.get("_lg_by")))]
    sub = _d(row.get("_lg_sub"))
    if sub and sub != req:
        ev.append((sub, "Submitted in Land Gorilla", ""))
    ev += [(_d(row.get(field)), label, "") for label, field in MILESTONES[1:-2]]
    ev.append((_d(row.get("_pkg")), "Full draw package received", ""))
    steps = a.get("_steps") if a is not None else None
    if isinstance(steps, (list, tuple)):
        ev += [(d, step, who) for d, step, who in steps]
    if a is not None and not any(s == "Approved" for _, s, _ in ev) and _d(a.get("Approved on")):
        ev.append((_d(a.get("Approved on")), "Approved", _txt(a.get("Approved by"))))
    ev.append((wire, "Wired", ""))
    t = pd.DataFrame([(pd.Timestamp(d), s, w) for d, s, w in ev if d], columns=["Date", "Step", "By"])
    return t.sort_values("Date", kind="stable").reset_index(drop=True)


def _draw_names(df: pd.DataFrame) -> list:
    """How each draw reads in a picker: 'Loan 12345 · Draw 3 · Maple Row LLC  (requested 9/2/2026, wired 9/12/2026)'."""
    out = []
    for loan, no, deal, req, wired in zip(_col(df, "Loan_Number__c"), _draw_nos(df), _col(df, "Deal__r.Name"),
                                          _day(_col(df, "_lg_req")), _day(_col(df, "_wire"))):
        when = ([f"requested {_mdy(req)}"] if pd.notna(req) else []) + ([f"wired {_mdy(wired)}"] if pd.notna(wired)
                                                                          else [])
        bits = [f"Loan {_txt(loan) or '?'}", f"Draw {no}"] + ([_txt(deal)] if _txt(deal) else [])
        out.append(" · ".join(bits) + (f"  ({', '.join(when)})" if when else ""))
    return out


def _draw_table(d: pd.DataFrame, inst: str, ap: pd.DataFrame | None = None) -> pd.DataFrame:
    """The draws in plain, business-friendly columns (shown in the app and used for the Excel export)."""
    ids = _col(d, "Id")
    t = pd.DataFrame({
        "Loan #": _col(d, "Loan_Number__c"), "Draw #": pd.Series(_draw_nos(d), index=d.index, dtype=object),
        "Deal": _col(d, "Deal__r.Name"), "Borrower": _col(d, "Borrower_Name_Text__c"),
        "Status": _col(d, "Status__c"), "Requested": _dates(_col(d, "_lg_req")),
        "Package received": _dates(_col(d, "_pkg")), "Wired": _dates(_col(d, "_wire")),
        "Turn-time (bd)": pd.to_numeric(_col(d, "_fund_bd"), errors="coerce"),
        "Funded ($)": pd.to_numeric(_col(d, NET_FIELD), errors="coerce"),
    })
    for c in ("Approved by", "Approved on", "Approval comments"):
        t[c] = ids.map(ap[c]) if ap is not None and c in ap else None
    t["Approved on"] = _dates(t["Approved on"])
    t["Held up by"] = [delay_labels(n) for n in _col(d, NOTES_FIELD)]
    t["Coordinator"] = _col(d, "Advance_Coordinator__r.Name")
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
        "Days to a full package (bd)": C.NumberColumn("Days to a full package", format="%d bd",
                                                      help="Business days from the request to a full package"),
        "Request → wire (bd)": C.NumberColumn("Request → wire", format="%d bd"),
        "Funded ($)": C.NumberColumn("Funded", format="dollar", step=1),
        "Approved on": C.DateColumn(format="M/D/YYYY"),
        "Held up by": C.ListColumn(help="What the draw's notes say it was waiting on"),
        "What the notes say": C.TextColumn(width="large", help="The note entries behind those hold-ups"),
        "Approval comments": C.TextColumn(width="large"),
        "Notes": C.TextColumn(width="large"),
        "Salesforce": C.LinkColumn(display_text="Open", width="small"),
    }


def _excel_table(fg: pd.DataFrame, table: pd.DataFrame) -> pd.DataFrame:
    """The turn-time data dump: every column in the table plus the milestone dates and the borrower-side numbers
       (kept out of the overview, but useful in a pivot by account)."""
    x = table.drop(columns=["Salesforce"]).copy()
    if "Held up by" in x:
        x["Held up by"] = x["Held up by"].map(lambda r: ", ".join(r) if r else None)
    pos = x.columns.get_loc("Package received")
    for i, (label, field) in enumerate(MILESTONES[1:-2]):    # inspection ordered … internal review complete
        x.insert(pos + i, label, _dates(_col(fg, field)))
    x.insert(x.columns.get_loc("Turn-time (bd)") + 1, "Days to a full package (bd)",
             pd.to_numeric(_col(fg, "_wait_bd"), errors="coerce"))
    x.insert(x.columns.get_loc("Days to a full package (bd)") + 1, "Request → wire (bd)",
             pd.to_numeric(_col(fg, "_days"), errors="coerce"))
    x.insert(x.columns.get_loc("Borrower") + 1, "Account", _col(fg, "Deal__r.Account.Name"))
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


def _turn_stats(fg: pd.DataFrame, same_day: int) -> dict:
    """The package-received -> wire numbers for a set of draws, plus how many could be measured."""
    turn = fg["_fund_bd"].dropna()
    eligible = int((pd.to_datetime(fg["_wire"], errors="coerce") >= PKG_LAUNCH).sum())
    return {"turn": turn, "n": len(turn), "eligible": eligible,
            "coverage": len(turn) / eligible * 100 if eligible else None,
            "same": (turn == same_day).mean() * 100 if len(turn) else None,
            "fast": (turn <= 3 + same_day).mean() * 100 if len(turn) else None}


def render_pipeline(inst, tok, rt, sel, same_day, links=()):
    st.header("Draw turn-time")
    st.caption("How fast a draw is wired once the borrower's full draw package is in. All figures are business days.")
    if links:                                   # the other two pages, one click away
        lk = st.columns([1, 1, 4])
        lk[0].page_link(links[0], label="Find a draw", icon=":material/search:")
        lk[1].page_link(links[1], label="Funding request", icon=":material/request_quote:")
    c = st.columns([2, 1.3, 1.4])
    gran = c[0].segmented_control("View by", ["Month", "Quarter", "Year", "Date range"], default="Month",
                                  key="ov_gran") or "Month"
    if gran == "Date range":
        rng = c[1].date_input("Wire dates", value=(_today() - timedelta(days=30), _today()), key="ov_rng")
        if len(rng) != 2:
            st.info("Pick a start and an end date."); return
        start, end, gi = rng[0], min(rng[1], _today()), "Monthly"
    else:
        gi = {"Month": "Monthly", "Quarter": "Quarterly", "Year": "Yearly"}[gran]
        opts = {"Month": [6, 12, 24], "Quarter": [4, 8, 12], "Year": [3, 5]}[gran]
        nper = c[1].selectbox("How many", opts, index=0 if gran == "Month" else 1 if gran == "Quarter" else 0,
                              key=f"ov_n_{gran}")
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

    s = _turn_stats(fg, same_day)
    funded = pd.to_numeric(fg.get(AMT), errors="coerce").sum() if AMT else None
    holds = int(flight["Status__c"].astype(str).str.contains("Hold|Pending Borrower|Revision", case=False,
                                                              na=False).sum()) if "Status__c" in flight else 0
    k = st.columns(4)
    if s["eligible"]:
        k[0].metric("Turn-time · full package → wire", _bd(s["turn"].mean()) if s["n"] else "—",
                    f"{s['same']:.0f}% same day" if s["same"] is not None else None, delta_color="off",
                    delta_arrow="off", border=True,
                    help="Average business days from the full draw package being received to the wire.")
        k[1].metric("Funded within 3 days of a full package", f"{s['fast']:.0f}%" if s["fast"] is not None else "—",
                    border=True)
        k[2].metric("Median turn-time", _bd(s["turn"].median()) if s["n"] else "—", border=True)
        k[3].metric("Measured on", f"{s['n']:,} of {s['eligible']:,} draws",
                    f"{s['coverage']:.0f}% have a package date" if s["coverage"] is not None else None,
                    delta_color="off", delta_arrow="off", border=True,
                    help="Draws wired since Aug 3, 2026 that have a full-package date in the right order. "
                         "See the Package date check tab.")
    else:
        st.info("Turn-time starts with draws wired on or after **Aug 3, 2026**, when the *Full Draw Package "
                "Received* date began being recorded. Pick a period from August 2026 on.", icon=":material/info:")
    k2 = st.columns(4)
    k2[0].metric(f"Draws funded · {feat}", f"{len(fg):,}", money(funded) if funded is not None else None,
                 delta_color="off", delta_arrow="off", border=True)
    k2[1].metric("Open draws now", f"{len(flight):,}", f"{holds} on hold or waiting" if holds else None,
                 delta_color="off", delta_arrow="off", border=True)
    if s["n"]:
        st.caption(f"In **{feat}**, once a full draw package was in, draws were wired in **{_bd(s['turn'].mean())}** "
                   f"on average" + (f" — **{s['same']:.0f}% the same day**" if s["same"] else "") + ".")

    _period_funding_file(inst, tok, rt, feat, period_where, p_end)      # the template, right where it's needed

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

    t_sum, t_draws, t_detail, t_check = st.tabs([":material/insights: Summary", f":material/list: Draws ({len(fg):,})",
                                                 ":material/manage_search: Draw & deal details",
                                                 ":material/fact_check: Package date check"])
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
        _holdups_chart(fg, f"What held draws up · {feat}")
        with st.expander("How these numbers work"):
            st.markdown(
                "- **Package received** — the *Full Draw Package Received* date in Salesforce: the day the borrower's "
                "package was complete. It started being recorded on **Aug 3, 2026**, so turn-time begins there.\n"
                "- **Wired** — the draw's funding date in Land Gorilla.\n"
                "- **Turn-time** = package received → wired, in business days (weekends and federal holidays skipped).\n"
                "- How long borrowers take to get to a full package is their responsibility, so it isn't an overview "
                "number — see it by account or deal under **Find a draw**.\n"
                "- **Held up by** — what a draw's notes say it was waiting on: lien waivers, invoices, title, the "
                "inspection and so on.\n"
                "- Draws without a package date, or with dates out of order, are left out of turn-time and listed on "
                "the Package date check tab.")
    with t_draws:
        st.dataframe(table, hide_index=True, width="stretch", height=460, column_config=_table_config())
        _period_workbook(feat, _excel_table(fg, table),
                         roll.drop(columns="_raw").rename(columns={"Draws": "Draws funded"}))
    with t_detail:
        _draw_details(inst, tok, rt, sel, fg, ap, feat)
    with t_check:
        _package_check(fg, table, same_day)


DEAL_FACTS = (("C", "Deal"), ("D", "Borrower"), ("E", "Sponsor"), ("J", "Project type"), ("AJ", "RM / LO"),
              ("O", "Total loan commitment"), ("M", "Loan funded"), ("N", "Maturity"), ("AC", "Next payment due"),
              ("AD", "Interest reserve left"), ("AH", "Warehouse line"))
PROPERTY_FACTS = (("F", "Street"), ("G", "City"), ("H", "State"), ("I", "Property type"), ("K", "As-is value"),
                  ("L", "As-complete value"), ("Q", "Initial funding"), ("AB", "Current UPB"))


@sf_fragment
def _draw_details(inst, tok, rt, sel, fg: pd.DataFrame, ap: pd.DataFrame, feat: str):
    """Everything about one draw and the deal behind it, picked from the period's draws (most recent request first).
       A fragment, so picking a draw doesn't reload the rest of the page."""
    d = _by_date(fg)
    d = d[_col(d, "Id").map(_txt) != ""]
    names = dict(zip(d["Id"], _draw_names(d)))
    pick = st.selectbox("Draw", list(names), index=None, format_func=lambda i: names.get(i, str(i)),
                        key=f"detail_{feat}", placeholder="Pick a draw — type a loan #, deal or draw # to search",
                        help="This period's draws, most recent request first.")
    if pick is None:
        st.info("Pick a draw to see its timeline, money, approvals and notes, and the deal behind it: terms, "
                "property and construction progress.", icon=":material/manage_search:")
        return
    row = d[d["Id"] == pick].iloc[0]
    with st.container(border=True):
        _render_draw(row, ap, inst)
    st.subheader("Deal", divider="gray")
    cfg, loan = lg_config(), _txt(row.get("Loan_Number__c"))
    with st.spinner("Loading the deal…"):
        ov = _deal_details(inst, tok, rt, row)
        cols = [f for f in ("Id", "Name", "Loan_Advance_Number__c", CONTAINER_FIELD) if f in sel.split(",")]
        hist = run_soql(inst, tok, f"SELECT {','.join(cols)} FROM Advance__c WHERE RecordTypeId='{rt}' "
                                   f"AND Loan_Number__c='{soql_escape(loan)}'") if cfg and loan else None
    if cfg and loan:
        st.subheader("Construction progress", divider="gray")
        if not ov:
            st.caption(f"Land Gorilla has no loan record for loan {_md(loan)}, so there's no project progress to "
                       "show — only the loan's draws.")
        _render_lg_for_loan(cfg, ov or {}, hist if hist is not None and not hist.empty else row.to_frame().T)


def _check_label(col: str) -> str:
    """'chk:Completion LTC % — deal Max LTC' -> 'Max LTC' (the Salesforce field's own name)."""
    return re.sub(r"^(?:deal|property)\s+", "", col[4:].split(" — ")[-1])


def _deal_details(inst, tok, rt, row: pd.Series) -> dict | None:
    """The deal behind a draw — its terms and people, then each property — with Land Gorilla first where it has the
       value, as in the funding-request file. Returns Land Gorilla's overview of the loan, if any."""
    rid = _txt(row.get("Id"))
    rows, ftypes, _ = funding_request_rows(inst, tok, rt, f"Id='{soql_escape(rid)}'") if rid else (pd.DataFrame(),
                                                                                                    {}, [])
    if rows.empty:
        st.caption("Salesforce has no deal details for this draw.")
        return None
    merged, mtypes, lg = lg_overlay(rows, ftypes, lg_config())
    first = merged.iloc[0]
    facts = [(label, _show(first.get(col), mtypes.get(col))) for col, label in DEAL_FACTS]
    facts.insert(1, ("Account", _show(row.get("Deal__r.Account.Name"))))
    facts += [("Draw coordinator", _show(row.get("Advance_Coordinator__r.Name"))),
              ("Draw analyst", _show(row.get("Advance_Analyst__r.Name")))]
    facts += [(_check_label(c), _show(first.get(c), mtypes.get(c))) for c in merged.columns
              if str(c).startswith("chk:") and c[4:] in DEAL_CHECKS]
    facts = [(k, v) for k, v in facts if v]
    if facts:
        half = (len(facts) + 1) // 2
        for col, part in zip(st.columns(2), (facts[:half], facts[half:])):
            col.dataframe(pd.DataFrame(part, columns=["Detail", "Value"]), hide_index=True, width="stretch")
    props = pd.DataFrame({label: [_show(v, ftypes.get(col)) for v in _col(rows, col)] for col, label in PROPERTY_FACTS})
    if len(rows) > 1:
        props["This draw"] = [_show(v, "currency") for v in _col(rows, "_pa_amount")]
    for c in [c for c in rows.columns if str(c).startswith("chk:") and c[4:] in PROPERTY_CHECKS]:
        props[_check_label(c)] = [_show(v, ftypes.get(c)) for v in rows[c]]
    props = props.loc[:, props.notna().any()]
    if len(props.columns):
        st.markdown("**Properties**" if len(rows) > 1 else "**Property**")
        st.dataframe(props, hide_index=True, width="stretch")
    if any(x is True or x is np.True_ for x in _col(rows, "_multi_property")):
        st.caption("Portfolio loan — this draw couldn't be matched to one property.")
    return lg["loans"].get(_txt(first.get("B")) or _txt(row.get("Loan_Number__c")))


def _package_check(fg: pd.DataFrame, table: pd.DataFrame, same_day: int):
    """Is the full-package date trustworthy for this period? Checked live, in plain English."""
    s = _turn_stats(fg, same_day)
    st.markdown("The *Full Draw Package Received* date is new (Aug 3, 2026), so these checks show whether it's "
                "being recorded well enough to measure turn-time.")
    if not s["eligible"]:
        st.info("No draws in this period were wired after Aug 3, 2026."); return
    wire = pd.to_datetime(fg["_wire"], errors="coerce")
    pkg = pd.to_datetime(fg["_pkg"], errors="coerce")
    req = pd.to_datetime(fg["_lg_req"], errors="coerce")
    after = wire >= PKG_LAUNCH
    measured = pkg.notna()
    on_wire_day = (pkg[measured] == wire[measured]).mean() * 100 if measured.any() else None
    total = pd.to_numeric(fg["_days"], errors="coerce")
    med_with = total[after & measured].median()
    med_without = total[after & ~measured].median()
    n_without = int((after & ~measured & total.notna()).sum())
    excluded = int(fg["_check"].isin(CHECK_EXCLUDED).sum())
    old = int(fg["_check"].astype(str).str.startswith("Package date predates").sum())

    rows = []
    cov = s["coverage"]
    rows.append(("Draws with a package date", f"{cov:.0f}%  ({s['n']} of {s['eligible']})",
                 "Too few draws to judge" if s["eligible"] < 30 else
                 "Good — nearly every draw has one" if cov >= 90 else
                 "Fair — some draws are missing it, so treat turn-time as indicative" if cov >= 75 else
                 "Low — too many draws are missing it for turn-time to speak for the period"))
    if on_wire_day is not None:
        rows.append(("Package date on the wire day", f"{on_wire_day:.0f}%",
                     "Looks filled in on funding day — the date may be entered when the draw is wired"
                     if on_wire_day >= 90 else "Watch — close to being entered on funding day" if on_wire_day >= 85
                     else "Normal — complete packages are often wired the same day (September notes: 74%)"))
    if n_without >= 10 and pd.notna(med_with) and pd.notna(med_without):
        gap = med_without - med_with
        rows.append(("Do missing dates hide slow draws?", f"request → wire median {med_with:.0f} bd with a date, "
                     f"{med_without:.0f} bd without",
                     "No — draws without a package date look like the rest" if abs(gap) <= 1 else
                     f"Yes — the draws missing a date took {gap:.0f} more business days end to end, so turn-time "
                     "may look better than it is" if gap > 1 else
                     "The draws missing a date were the quicker ones, so turn-time may look slower than it is"))
    share_bad = excluded / s["eligible"] * 100
    rows.append(("Dates out of order", f"{excluded}  ({share_bad:.0f}%)",
                 "Fine — listed below" if share_bad <= 2 else "Data-entry problem — review the list below"))
    if old:
        rows.append(("Old values ignored", str(old), "Dates from before Aug 3, 2026 belong to the field's old use"))
    st.dataframe(pd.DataFrame(rows, columns=["Check", "Result", "What it means"]), hide_index=True, width="stretch",
                 column_config={"What it means": st.column_config.TextColumn(width="large")})

    months = fg.assign(_m=wire.dt.to_period("M").astype(str))[after]
    if months["_m"].nunique() > 1:
        by = []
        for month, g in months.groupby("_m"):
            t = g["_fund_bd"].dropna()
            by.append({"Month": month, "Draws": len(g), "With a package date": f"{g['_pkg'].notna().mean() * 100:.0f}%",
                       "Same day": f"{(t == same_day).mean() * 100:.0f}%" if len(t) else "—",
                       "Median turn-time": _bd(t.median()) if len(t) else "—"})
        st.markdown("**By month**")
        st.dataframe(pd.DataFrame(by), hide_index=True, width="stretch")
    flagged = table.assign(Issue=fg["_check"].values)
    flagged = flagged[flagged["Issue"].isin(CHECK_EXCLUDED) | (flagged["Issue"] == "No package-received date")]
    if len(flagged):
        with st.expander(f"Draws to check ({len(flagged)})", icon=":material/rule:"):
            st.dataframe(flagged[["Loan #", "Draw #", "Deal", "Requested", "Package received", "Wired", "Issue",
                                  "Salesforce"]], hide_index=True, width="stretch", column_config=_table_config())


@sf_fragment
def _period_funding_file(inst, tok, rt, feat, period_where, p_end):
    """Melanie's funding-request workbook for the period's wired draws. A fragment, so building it doesn't
       reload the whole page."""
    tpl = _find_template()
    with st.container(border=True):
        st.markdown(f"**:material/description: Funding-request file** · Melanie's workbook, one row per draw wired "
                    f"in {feat}")
        if not tpl:
            st.error("The funding-request template is missing from the app."); return
        e = st.columns([1.3, 1, 1.1])
        who = e[0].text_input("Submitted by (optional)", key=f"pby_{feat}")
        fdate = e[1].date_input("Scheduled funding date", value=p_end, key=f"pdate_{feat}")
        e[2].markdown("&nbsp;")
        if e[2].button("Build funding-request file", type="primary", icon=":material/build:", key=f"mel_{feat}",
                       width="stretch"):
            with st.spinner("Filling the template from Land Gorilla and Salesforce…"):
                rows, ftypes, _ = funding_request_rows(inst, tok, rt, period_where)
                xlsx, n = build_funding_file(inst, tok, rt, tpl, rows, ftypes, who, fdate)
            st.download_button(f"Download funding-request file ({n:,} draws)", xlsx, mime=XLSX_MIME, type="primary",
                               file_name=f"Construction-Renovation Funding Request {feat}.xlsx",
                               icon=":material/download:", key=f"meldl_{feat}", on_click="ignore")
            st.toast("Funding-request file is ready.", icon="✅")


@sf_fragment
def _period_workbook(feat, xtable, roll):
    if st.button("Build turn-time workbook (Excel)", icon=":material/table_view:", key=f"xls_{feat}"):
        with st.spinner("Building the workbook…"):
            xlsx = build_excel({"Summary": roll, "Draws": xtable})
        st.download_button("Download turn-time workbook", xlsx, mime=XLSX_MIME, icon=":material/download:",
                           file_name=f"Draw turn-time {feat}.xlsx", key=f"dl_{feat}", on_click="ignore")


def _holdups_chart(d: pd.DataFrame, title: str):
    """How often each hold-up shows up in these draws' notes (share of draws), coloured by who it was on."""
    keys = [k for n in _col(d, NOTES_FIELD) for k in delay_keys(n) if k not in ROUTINE_REASONS]
    st.markdown(f"**{title}**")
    if not keys:
        st.caption("Nothing in these draws' notes points to a hold-up.")
        return
    n = pd.Series(keys, dtype=object).value_counts()
    data = pd.DataFrame({"Hold-up": [DELAY_REASONS[k][0] for k in n.index], "Draws": n.values,
                         "Share": n.values / max(len(d), 1),
                         "Waiting on": [WAITING_ON.get(DELAY_REASONS[k][1], "") for k in n.index]})
    if alt is None:
        st.bar_chart(data.set_index("Hold-up")["Share"])
    else:
        st.altair_chart(alt.Chart(data).mark_bar().encode(
            x=alt.X("Share:Q", title="Share of draws", axis=alt.Axis(format="%")),
            y=alt.Y("Hold-up:N", sort="-x", title=None),
            color=alt.Color("Waiting on:N", legend=alt.Legend(orient="bottom", title="Waiting on")),
            tooltip=["Hold-up", "Draws", alt.Tooltip("Share:Q", format=".0%"), "Waiting on"]),
            width="stretch")
    st.caption(f"Read from the notes of {len(d):,} draws; a draw can have more than one. Routine inspections are "
               "left out unless they ran late.")


def _render_charts(roll: pd.DataFrame, fg: pd.DataFrame, feat: str, same_day: int):
    """Turn-time distribution for the period, plus dollars per period."""
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
        st.markdown(f"**Full package → wire, {feat}**")
        st.altair_chart(alt.Chart(dist).mark_bar(color="#4C78A8").encode(
            x=alt.X("Business days:N", sort=list(dist["Business days"]), axis=alt.Axis(labelAngle=0)),
            y=alt.Y("Draws:Q"), tooltip=["Business days", "Draws"]), width="stretch")
    r = roll.drop(columns="Period").rename(columns={"_raw": "Period"})
    st.markdown("**Funded per period**")
    st.altair_chart(alt.Chart(r).mark_bar(color="#54A24B").encode(
        x=alt.X("Period:N", sort=list(r["Period"]), title=None, axis=alt.Axis(labelAngle=-45)),
        y=alt.Y("Funded ($):Q", title=None, axis=alt.Axis(format="$.2~s")),   # $18M, $2.5M — no repeats
        tooltip=["Period", alt.Tooltip("Funded ($):Q", format="$,.0f"), "Draws"]), width="stretch")


# ───────────────────────────── Find a draw ──────────────────────────────────
def lg_filenumber(loan_no) -> str | None:
    """LG file number is 'rb0'+the SF loan number, but only for the clean numeric RB0 book."""
    s = str(loan_no).strip()
    return f"rb0{s}" if s.isdigit() else None


def render_loan_detail(inst, tok, rt, sel, same_day):
    st.header("Find a draw")
    t_search, t_open, t_accounts = st.tabs([":material/search: Search", ":material/pending_actions: Not wired yet",
                                            ":material/groups: Who's falling short"])
    with t_search:
        _draw_search(inst, tok, rt, sel, same_day)
    with t_open:
        _open_draws(inst, tok, rt, sel, same_day)
    with t_accounts:
        _accounts_view(inst, tok, rt, sel, same_day)


@sf_fragment
def _open_draws(inst, tok, rt, sel, same_day):
    """Every draw that isn't wired yet — its status, how long it's been open, and what its notes say it's waiting on
       — most recent request first. A fragment, so loading it doesn't redraw the search results."""
    st.caption("Draws that haven't been wired yet: where each one stands and what its notes say it's waiting on.")
    data = st.session_state.get("open_data")                # built on request, then kept for the session
    if st.button("Refresh" if data else "Show draws not wired yet", icon=":material/pending_actions:",
                 key="open_build"):
        with st.spinner("Loading open draws and Land Gorilla dates…"):
            df = run_soql(inst, tok, f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND {OPEN_WHERE}")
            df = _by_date(apply_lg_correction(add_intervals(df), lg_config(), same_day)) if not df.empty else df
        st.session_state["open_data"] = data = (df, pd.Timestamp.now(tz="America/New_York"))
    if not data:
        return
    df, asof = data
    if df.empty:
        st.info("Every draw has been wired."); return
    notes = _col(df, NOTES_FIELD)
    req = _request_date(df)
    latest = [latest_entry(n, w.date() if pd.notna(w) else None) for n, w in zip(notes, req)]
    t = pd.DataFrame({
        "Loan #": _col(df, "Loan_Number__c").values, "Draw #": _draw_nos(df), "Deal": _col(df, "Deal__r.Name").values,
        "Status": _col(df, "Status__c").values, "Requested": req.values,
        "Open (bd)": bdays(req, pd.Series(pd.Timestamp(_today()), index=df.index), same_day).values,
        "Waiting on": [delay_labels(n) for n in notes],
        "Latest note": [f"{_mdy(x[0])}: {x[1]}" if x else None for x in latest],
        "Coordinator": _col(df, "Advance_Coordinator__r.Name").values,
        "Salesforce": _col(df, "Id").map(lambda i: f"{inst.rstrip('/')}/{i}" if isinstance(i, str) and i
                                         else None).values,
    })
    by_status = _col(df, "Status__c").fillna("(none)").value_counts()
    st.markdown(f"**{len(df):,} draws not wired yet** (as of {asof.strftime('%I:%M %p').lstrip('0')}) · "
                + " · ".join(f"{_md(s)} {n}" for s, n in by_status.items()))
    st.dataframe(t, hide_index=True, width="stretch", height=420, column_config={
        **_table_config(),
        "Open (bd)": st.column_config.NumberColumn("Open", format="%d bd",
                                                   help="Business days since the draw was requested"),
        "Waiting on": st.column_config.ListColumn(help="What the draw's notes say it's waiting on"),
        "Latest note": st.column_config.TextColumn(width="large", help="The newest dated entry in the notes")})
    _holdups_chart(df, "What open draws are waiting on")


def _draw_search(inst, tok, rt, sel, same_day):
    text = st.text_input("Search", key="ld_q", placeholder="Loan #, property or deal name, borrower, or draw #",
                         help="Type any part of it.")
    if not text:
        st.info("Search by loan number, property / deal name, borrower or draw number to see each draw's full "
                "cycle and the borrower's track record.", icon=":material/search:")
        return
    q = soql_escape(text)
    searchable = [f for f in ["Loan_Number__c", "Deal__r.Name", "Borrower_Name_Text__c", "Name",
                              "Loan_Advance_Number__c", "Deal__r.Account.Name"] if f in sel.split(",")]
    ors = " OR ".join(f"{f} LIKE '%{q}%'" for f in searchable) or "Id != null"
    order = "CreatedDate DESC" if "CreatedDate" in sel.split(",") else f"Loan_Number__c, {REQ_FIELD}"
    with st.spinner("Searching…"):
        df = add_intervals(run_soql(inst, tok, f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' "
                                               f"AND ({ors}) ORDER BY {order} LIMIT {SEARCH_LIMIT + 1}"))
    if df.empty:
        st.warning("No matching draws."); return
    if len(df) > SEARCH_LIMIT:
        df = df.head(SEARCH_LIMIT)
        st.warning(f"Showing the {SEARCH_LIMIT} most recent matching draws — narrow the search to see older ones.")
    cfg = lg_config()
    with st.spinner("Loading draw history…"):
        df = _newest_first(apply_lg_correction(df, cfg, same_day))
        ap = approvals_for(inst, tok, tuple(_col(df, "Id").dropna()))
    nloans = df["Loan_Number__c"].nunique(dropna=False) if "Loan_Number__c" in df else len(df)
    view = "By loan"
    if nloans > 1:
        view = st.segmented_control(
            "Show", ["Newest first", "By loan"], default="Newest first", key="ld_view",
            help="Newest first: every draw in request order, the most recent at the top. By loan: each loan's draws "
                 "together, with the loan's track record and construction progress.") or "Newest first"
    if view == "Newest first":
        st.caption(f"{len(df)} draw(s) on {nloans} loans · most recent request at the top, earliest at the bottom")
        for _, row in _by_date(df).iterrows():
            with st.container(border=True):
                who = [x for x in (_txt(row.get("Deal__r.Name")), _txt(row.get("Deal__r.Account.Name"))) if x]
                st.caption(f"**Loan {_md(row.get('Loan_Number__c'))}**" + "".join(f" · {_md(x)}" for x in who))
                _render_draw(row, ap, inst)
        return

    st.caption(f"{len(df)} draw(s) on {nloans} loan(s) · loans with the most recent request first, each loan's "
               "latest draw on top")
    groups = list(df.groupby("Loan_Number__c", dropna=False, sort=False))       # already newest first
    if len(groups) > MAX_LOAN_SECTIONS:
        st.warning(f"Showing the {MAX_LOAN_SECTIONS} loans with the most recent draws — narrow the search to see the "
                   "rest.")
        groups = groups[:MAX_LOAN_SECTIONS]
    tmpl = str((cfg or {}).get("template_id", LG_TEMPLATE_ID_DEFAULT))
    loans = {}
    if cfg:
        try:
            loans = lg_loans(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), tmpl,
                             [loan for loan, _ in groups])
        except Exception:
            loans = {}

    for loan_no, g in groups:
        def first(col):
            return _txt(g[col].dropna().iloc[0]) if col in g and g[col].notna().any() else ""
        with st.container(border=True):
            st.subheader(f"Loan {_md(loan_no)} · {_md(first('Deal__r.Name'))}")
            sub = [_md(x) for x in (first("Deal__r.Account.Name"), first("Borrower_Name_Text__c")) if x]
            st.caption("  ·  ".join(sub + [f"{len(g)} draw(s)"]))
            _track_record(g)
            for _, row in g.iterrows():
                _render_draw(row, ap, inst)
            ov = loans.get(_txt(loan_no))
            if cfg and ov:
                with st.expander("Construction progress", icon=":material/construction:"):
                    _render_lg_for_loan(cfg, ov, g)


def _track_record(g: pd.DataFrame):
    """How this loan's borrower does at getting to a full package — and what holds them up."""
    wait = pd.to_numeric(g["_wait_bd"], errors="coerce").dropna()
    top, _who, _n = _waiting_summary([delay_keys(n) for n in _col(g, NOTES_FIELD)])
    m = st.columns(3)
    m[0].metric("Days to a full package", _bd(wait.mean()) if len(wait) else "—", border=True,
                help="Average business days from the borrower's request to a complete package, for this loan's draws "
                     "with a package date.")
    m[1].metric("Draws measured", f"{len(wait)} of {len(g)}", border=True)
    m[2].metric("Most common hold-up", top[0] if top else "None noted", border=True,
                help="What this loan's draw notes most often say it was waiting on (routine inspections aside).")


def _render_draw(row: pd.Series, ap: pd.DataFrame, inst: str):
    """One draw: what it funded and how long it took, what held it up, its latest note, a timeline of what actually
       happened (with each approver), the construction dollars, and the full notes."""
    rid = _txt(row.get("Id"))
    head = st.columns([4, 1])
    status = _txt(row.get("Status__c"))
    head[0].markdown(f"**Draw {_md(_draw_no(row))}**" + (f" &nbsp; :blue-badge[{_md(status)}]" if status else ""))
    if rid:
        head[1].link_button("Open in Salesforce", f"{inst.rstrip('/')}/{rid}", icon=":material/open_in_new:",
                            width="stretch")
    a = ap.loc[rid] if rid and rid in ap.index else None
    reasons = delay_labels(row.get(NOTES_FIELD))
    m = st.columns(4)
    m[0].metric("Funded", money(row.get(NET_FIELD)) if pd.notna(row.get(NET_FIELD)) else "—", border=True)
    m[1].metric("Turn-time", _bd(row.get("_fund_bd")), border=True,
                help="Business days from full draw package received to wire")
    m[2].metric("Days to a full package", _bd(row.get("_wait_bd")), border=True,
                help="Business days from the request to a full package — the borrower's part")
    m[3].metric("Request → wire", _bd(row.get("_days")), border=True)
    if reasons:
        st.markdown("**Held up by:** " + " ".join(f":orange-badge[{r}]" for r in reasons))
    latest = latest_entry(row.get(NOTES_FIELD), _d(row.get("_wire")) or _d(row.get("_lg_req")))
    if latest:
        st.caption(f":material/history: **Latest note, {_mdy(latest[0])}:** {_md(latest[1])}")
    lg = [f"{lbl} {money(row.get(col)).replace('$', chr(92) + '$')}" for lbl, col in
          (("requested", "_lg_requested"), ("approved", "_lg_amount"))
          if pd.notna(pd.to_numeric(row.get(col), errors="coerce"))]
    lg += [f"status {_md(row.get('_lg_status'))}"] if _txt(row.get("_lg_status")) else []
    payees = row.get("_lg_payees")
    lg += [f"paid to {', '.join(_md(p) for p in payees)}"] if isinstance(payees, (list, tuple)) and payees else []
    facts = (["**Land Gorilla:** " + " · ".join(lg)] if lg else []) + (
        [f"**Inspection:** {_md(row.get('Inspection_Method__c'))}"] if _txt(row.get("Inspection_Method__c")) else [])
    if facts:
        st.caption("  ·  ".join(facts))

    tl = _timeline(row, a)
    if len(tl):
        st.dataframe(tl, hide_index=True, width="stretch",
                     column_config={"Date": st.column_config.DateColumn(format="M/D/YYYY", width="small"),
                                    "Step": st.column_config.TextColumn(width="medium"),
                                    "By": st.column_config.TextColumn(help="Who did it, where the record says")})
    if _txt(row.get("_check")):
        st.caption(f":material/error: {row['_check']}")
    if a is not None and _txt(a.get("Approval comments")):
        st.caption("**Approval comments:**  \n" + _md(a["Approval comments"]).replace("\n", "  \n"))
    elif (a is None or not _txt(a.get("Approved by"))) and _txt(row.get("LastModifiedBy.Name")):
        when = _d(row.get("LastModifiedDate"))
        st.caption(f"Last updated by {_md(row['LastModifiedBy.Name'])}" + (f" on {_mdy(when)}" if when else ""))
    vals = pd.to_numeric(pd.Series([row.get(f) for f, _ in CONSTRUCTION_FIELDS]), errors="coerce")
    money_bits = [f"{lbl}: {money(v)}" for (_, lbl), v in zip(CONSTRUCTION_FIELDS, vals) if pd.notna(v) and v]
    if money_bits:                                  # "$" escaped: two of them would turn the text between into math
        st.caption("  ·  ".join(money_bits).replace("$", "\\$"))
    note = _txt(row.get(NOTES_FIELD))
    if note:
        with st.expander("Notes", icon=":material/sticky_note_2:"):
            st.text(note)


@sf_fragment
def _accounts_view(inst, tok, rt, sel, same_day):
    """Which borrowers keep taking longest to get to a full package, and what their draws were waiting on — a
       relationship view, not a KPI. A fragment, so picking an account doesn't redraw the search results."""
    st.caption("How long each account's borrowers take to get to a full draw package, and what their draws were "
               "waiting on. Built from draws wired in the window you pick.")
    c = st.columns([1.2, 1, 2])
    days = c[0].selectbox("Window", [30, 60, 90, 180], index=2, format_func=lambda d: f"Last {d} days",
                          key="acct_window")
    min_draws = c[1].number_input("At least this many draws", 1, 20, 3, key="acct_min")
    params = (int(days), int(same_day))
    if c[2].button("Build the list", icon=":material/groups:", key="acct_build"):
        start = max(_today() - timedelta(days=int(days)), PKG_LAUNCH.date())
        with st.spinner("Loading draws and Land Gorilla dates…"):
            df = apply_lg_correction(add_intervals(run_soql(
                inst, tok, f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' "
                           f"AND {WIRE_FIELD}>={start:%Y-%m-%d} AND {WIRE_FIELD}<={_today():%Y-%m-%d}")),
                lg_config(), same_day)
            out = []
            if not df.empty:
                df["_acct"] = _col(df, "Deal__r.Account.Name").map(_txt).replace("", "(no account)")
                df["_keys"] = [delay_keys(n) for n in _col(df, NOTES_FIELD)]
                for acct, g in df.groupby("_acct"):
                    wait = pd.to_numeric(g["_wait_bd"], errors="coerce").dropna()
                    total = pd.to_numeric(g["_days"], errors="coerce").dropna()
                    top, who, _n = _waiting_summary(g["_keys"])
                    out.append({"Account": acct, "Draws": len(g), "Loans": g["Loan_Number__c"].nunique(),
                                "Days to a full package (avg)": round(wait.mean(), 1) if len(wait) else None,
                                "Days to a full package (median)": wait.median() if len(wait) else None,
                                "Over 5 days": (wait > 5).mean() if len(wait) else None,
                                "Request → wire (median)": total.median() if len(total) else None,
                                "Usual hold-ups": top, "Mostly waiting on": who})
        st.session_state["acct_data"] = (params, start, df, pd.DataFrame(out))    # kept while you look around
    data = st.session_state.get("acct_data")
    if not data:
        return
    if data[0] != params:
        st.info("The window changed — click **Build the list** to rebuild it.", icon=":material/refresh:"); return
    _, start, df, everyone = data
    if df.empty:
        st.info("No draws were wired in that window."); return
    res = everyone[everyone["Draws"] >= min_draws] if len(everyone) else everyone
    if res.empty:
        st.info("No account has enough draws in that window."); return
    res = res.sort_values(["Days to a full package (avg)", "Account"], ascending=[False, True], na_position="last")
    st.caption(f"Since {_mdy(start)} · days to a full package count draws with a package date only (it began "
               "Aug 3, 2026) · hold-ups come from the draw notes.")
    st.dataframe(res, hide_index=True, width="stretch", height=420, column_config={
        "Days to a full package (avg)": st.column_config.NumberColumn(format="%.1f bd"),
        "Days to a full package (median)": st.column_config.NumberColumn(format="%d bd"),
        "Over 5 days": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1,
                                                       help="Share of the account's draws that took more than 5 "
                                                            "business days to reach a full package"),
        "Request → wire (median)": st.column_config.NumberColumn(format="%d bd",
                                                                 help="Business days from the request to the wire"),
        "Usual hold-ups": st.column_config.ListColumn(help="What the account's draw notes most often say they were "
                                                           "waiting on (routine inspections aside)"),
        "Mostly waiting on": st.column_config.TextColumn(help="Who most of those hold-ups were on: the borrower, "
                                                              "third parties (inspector, title), CoreVest, or "
                                                              "loan changes (extension, payoff)")})
    acct = st.selectbox("See an account's draws", list(res["Account"]), index=None,
                        key=f"acct_pick_{days}_{min_draws}", placeholder="Pick an account to see what each draw "
                                                                         "was waiting on")
    if acct:
        _account_draws(df[df["_acct"] == acct], acct, inst)


def _account_draws(d: pd.DataFrame, acct: str, inst: str):
    """One account's draws, most recent request first: how long each took and what it was waiting on."""
    d = _by_date(d)
    notes = _col(d, NOTES_FIELD)
    t = pd.DataFrame({
        "Loan #": _col(d, "Loan_Number__c").values, "Draw #": _draw_nos(d),
        "Deal": _col(d, "Deal__r.Name").values, "Requested": _dates(_col(d, "_lg_req")).values,
        "Package received": _dates(_col(d, "_pkg")).values, "Wired": _dates(_col(d, "_wire")).values,
        "Days to a full package (bd)": pd.to_numeric(_col(d, "_wait_bd"), errors="coerce").values,
        "Request → wire (bd)": pd.to_numeric(_col(d, "_days"), errors="coerce").values,
        "Held up by": [delay_labels(n) for n in notes],
        "What the notes say": [" · ".join(delay_evidence(n)) or None for n in notes],
        "Salesforce": _col(d, "Id").map(lambda i: f"{inst.rstrip('/')}/{i}" if isinstance(i, str) and i
                                        else None).values,
    })
    _top, who, counts = _waiting_summary([delay_keys(n) for n in notes])
    line = f"**{_md(acct)}** · {len(d)} draw(s)"
    if who:
        line += (f" · mostly waiting on **{who}** · usual hold-ups: "
                 + ", ".join(f"{_md(r).lower()} ({n})" for r, n in counts.head(3).items()))
    st.markdown(line)
    st.dataframe(t, hide_index=True, width="stretch", column_config=_table_config())


def _render_lg_for_loan(cfg: dict, ov: dict, g: pd.DataFrame):
    has_metrics = any(ov.get(k) is not None for k in ("project_pct", "to_finish", "balance", "last_draw"))
    m = st.columns(4) if has_metrics else [None] * 4          # (no metric is drawn when there are none)
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
    bits = [f"**{lbl}:** {_mdy(ov[key]) if isinstance(ov[key], date) else _md(ov[key])}"
            for lbl, key in [("Status", "status"), ("Program", "program"), ("Funded", "funded_date"),
                             ("Due", "due_date")] if ov.get(key)]
    loc = " ".join(_txt(x) for x in [ov.get("city"), ov.get("state")] if x)
    if loc:
        bits.append(f"**Location:** {_md(loc)}")
    if bits:
        st.caption("  ·  ".join(bits))
    if ov.get("risk"):
        st.caption("**Risk:** " + " · ".join(_md(r) for r in ov["risk"]))
    try:
        draws = lg_draws(cfg["user"], cfg["password"], bool(cfg.get("verify", True)),
                         _col(g, CONTAINER_FIELD).tolist())
    except Exception as exc:          # no exception text: a malformed secret would be echoed to every viewer
        st.caption(f"Land Gorilla's draw details are unavailable right now ({type(exc).__name__}).")
        draws = {}
    rows_, payees = [], []
    for _, row in g.iterrows():
        det = draws.get(row.get(CONTAINER_FIELD)) if isinstance(row.get(CONTAINER_FIELD), str) else None
        if not det:
            continue
        sf_draw = _txt(row.get("Loan_Advance_Number__c")) or _txt(row.get("Name"))
        rows_.append({"Draw": sf_draw, "Type": det.get("type"), "Status": det.get("status"),
                      "Requested": det.get("created"), "Submitted": det.get("submitted"),
                      "Funded": det.get("effective") or det.get("approved"), "Amount": det.get("amount")})
        payees += [{"Draw": sf_draw, "Paid to": p} for p in det.get("payees") or []]
    if rows_:
        st.markdown("**Draws**")
        ddf = pd.DataFrame(rows_)
        for col in ("Requested", "Submitted", "Funded"):
            ddf[col] = _dates(ddf[col])
        ddf = ddf.sort_values("Requested", ascending=False, na_position="last", kind="stable")   # newest first
        ddf["Amount"] = pd.to_numeric(ddf["Amount"], errors="coerce")
        dcol = st.column_config.DateColumn(format="M/D/YYYY")
        st.dataframe(ddf, hide_index=True, width="stretch",
                     column_config={"Requested": dcol, "Submitted": dcol, "Funded": dcol,
                                    "Amount": st.column_config.NumberColumn(format="dollar", step=1)})
        if payees:
            with st.expander(f"Paid to ({len(payees)})"):
                st.dataframe(pd.DataFrame(payees), hide_index=True, width="stretch")


# ───────────────────────────── Funding request (weekly file) ────────────────
RANGE_PRESETS = {"Last 7 days": 7, "Last 14 days": 14, "Month to date": None, "Custom": None}


def render_funding_request(inst, tok, rt):
    st.header("Funding request")
    st.caption("Melanie's weekly *Construction/Renovation Funding Request* workbook — the draws wired since the last "
               "request, filled from Land Gorilla first and Salesforce for the rest.")
    tpl = _find_template()
    if not tpl:
        st.error("The funding-request template is missing from the app."); return
    with st.container(border=True):
        c = st.columns([1.2, 1, 1.8])
        who = c[0].text_input("Submitted by (optional)", key="fr_by")
        fdate = c[1].date_input("Scheduled funding date", value=_today(), key="fr_date")
        mode = c[2].segmented_control("Draws", ["Wired", "Open (not yet wired)"], default="Wired",
                                      key="fr_mode") or "Wired"
        if mode == "Wired":
            r = st.columns([2, 1.4])
            preset = r[0].pills("Wire dates", list(RANGE_PRESETS), default="Last 14 days", key="fr_preset") \
                or "Last 14 days"
            if preset == "Custom":
                rng = r[1].date_input("From – to", value=(_today() - timedelta(days=14), _today()), key="fr_rng")
                if len(rng) != 2:
                    st.info("Pick a start and an end date."); return
                lo, hi = rng
            elif preset == "Month to date":
                lo, hi = _today().replace(day=1), _today()
            else:
                lo, hi = _today() - timedelta(days=RANGE_PRESETS[preset]), _today()
            where = f"{WIRE_FIELD}>={lo:%Y-%m-%d} AND {WIRE_FIELD}<={hi:%Y-%m-%d}"
            scope = f"wired {_mdy(lo)} – {_mdy(hi)}"
        else:
            where, scope = OPEN_WHERE, "not yet wired"
    with st.spinner("Loading draws…"):
        rows, ftypes, _ = funding_request_rows(inst, tok, rt, where)
    if rows.empty:
        st.info(f"No draws {scope}."); return
    rows, ftypes, _ = lg_overlay(rows, ftypes, lg_config(), loans=False)      # Land Gorilla's draw amounts

    f = st.columns(2)
    people = sorted({p for col in PERSON_FIELDS if col in rows for p in rows[col].dropna()})
    if people:
        person = f[0].selectbox("Coordinator / analyst", ["Everyone"] + people, key="fr_person",
                                help="Everyone gives the full list. Pick a person to match the file Salesforce "
                                     "generates for them.")
        if person != "Everyone":
            rows = rows[np.logical_or.reduce([rows[col].eq(person) for col in PERSON_FIELDS if col in rows])]
    if "Status__c" in rows:                                # every status in full, one click to drop one
        counts = rows["Status__c"].fillna("(none)").value_counts()
        picked = st.pills("Statuses", list(counts.index), selection_mode="multi", default=list(counts.index),
                          format_func=lambda s: f"{s} ({counts[s]})", key=f"fr_status_{mode}")
        rows = rows[rows["Status__c"].fillna("(none)").isin(picked)]
    total = pd.to_numeric(rows["U"], errors="coerce").sum() if "U" in rows else None
    m = st.columns(3)
    m[0].metric("Draws", f"{len(rows):,}", border=True)
    m[1].metric("Loans", f"{rows['B'].nunique() if 'B' in rows else 0:,}", border=True)
    m[2].metric("Current draw amount", money(total) if total is not None else "—", border=True)
    _funding_download(inst, tok, rt, tpl, rows, ftypes, who, fdate, scope)
    preview = pd.DataFrame({"Loan #": _col(rows, "B"), "Property": _col(rows, "F"), "Deal": _col(rows, "C"),
                            "Borrower": _col(rows, "D"), "Status": _col(rows, "Status__c"),
                            "Current draw": pd.to_numeric(_col(rows, "U"), errors="coerce"),
                            "Paid to": _col(rows, "AG"), "Coordinator": _col(rows, "Advance_Coordinator__r.Name"),
                            "Held up by": [delay_labels(n) for n in _col(rows, "AF")]})
    st.dataframe(preview, hide_index=True, width="stretch", height=360,
                 column_config={"Current draw": st.column_config.NumberColumn(format="dollar", step=1),
                                "Held up by": st.column_config.ListColumn(help="What the draw's notes say it was "
                                                                                "waiting on")})


@sf_fragment
def _funding_download(inst, tok, rt, tpl, rows, ftypes, who, fdate, scope):
    if st.button(f"Build funding-request file ({len(rows):,} draws {scope})", type="primary",
                 icon=":material/build:", disabled=rows.empty, key="fr_build"):
        with st.spinner("Filling the template from Land Gorilla and Salesforce…"):
            xlsx, n = build_funding_file(inst, tok, rt, tpl, rows, ftypes, who, fdate)
        st.download_button(f"Download funding-request file ({n:,} draws)", xlsx, mime=XLSX_MIME, type="primary",
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
        find = st.Page(lambda: render_loan_detail(inst, tok, rt, sel, same_day), title="Find a draw",
                       icon=":material/search:", url_path="find", default=start == "find")
        funding = st.Page(lambda: render_funding_request(inst, tok, rt), title="Funding request",
                          icon=":material/request_quote:", url_path="funding-request",
                          default=start == "funding-request")
        turn = st.Page(lambda: render_pipeline(inst, tok, rt, sel, same_day, links=(find, funding)),
                       title="Turn-time", icon=":material/timer:", url_path="turn-time", default=start == "turn-time")
        pg = st.navigation([turn, find, funding], position="top")      # pages along the top, always visible
        pg.run()
        st.session_state.pop("_sf_retried", None)
    except Exception as exc:
        if _session_expired(exc):
            _restart_session()
        raise


if __name__ == "__main__":
    main()
