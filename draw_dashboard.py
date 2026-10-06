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

MILESTONES = [
    ("Requested",                  REQ_FIELD),
    ("Inspection ordered",         "Date_Inspection_Ordered__c"),
    ("Inspection",                 "Date_Of_Inspection__c"),
    ("Inspection report received", "Date_Inspection_Report_Received__c"),
    ("Submitted for review",       "Date_Submitted_For_Approval__c"),
    ("Internal review complete",   "Date_Internal_Review_Complete__c"),
    ("Full draw package received", PKG_FIELD),
    ("Manager approval",           "Manager_Approval_Date__c"),
    ("Wired",                      WIRE_FIELD),
]
WISH = (["Id", "Name", "Loan_Number__c", "Loan_Advance_Number__c", "Deal__r.Name",
         "Borrower_Name_Text__c", "Lender__c", "Status__c", "Inspection_Method__c",
         "Advance_Coordinator__r.Name", NOTES_FIELD, CONTAINER_FIELD,
         NET_FIELD, GROSS_FIELD] + [m[1] for m in MILESTONES])
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
    """Pull the per-draw timeline + amount out of GET /api/clm/draw/{id}."""
    d = payload.get("data", payload) if isinstance(payload, dict) else {}
    total = (((d.get("lineItems") or {}).get("total")) or {})
    return {
        "name": d.get("name"),
        "status": d.get("status"),
        "created": _lg_date(d.get("createdDate")),
        "submitted": _lg_date(d.get("submittedDate")),
        "approved": _lg_date(d.get("approvedDate")),
        "funded": _lg_date(d.get("effectiveDate")),
        "amount": total.get("totalLessRetainage"),
    }


@st.cache_resource(show_spinner=False)
def get_lg_client(user: str, password: str, verify: bool):
    return LGClient(user, password, verify=verify)


@st.cache_data(ttl=600, show_spinner=False)
def lg_draw_detail(user: str, password: str, verify: bool, draw_id: str) -> dict | None:
    """Fetch + parse one draw's detail; returns None if gated/unavailable (degrades gracefully)."""
    try:
        cli = get_lg_client(user, password, verify)
        r = cli.get(f"/api/clm/draw/{draw_id}")
        if not r.ok:
            return {"_error": f"HTTP {r.status_code}"}
        return parse_draw_detail(r.json())
    except Exception as exc:
        return {"_error": str(exc)[:80]}


# ───────────────────────────── Pipeline (macro) ─────────────────────────────
_PRETTY = {
    "Loan_Number__c": "Loan #", "Deal__r.Name": "Property", "Borrower_Name_Text__c": "Borrower",
    "Status__c": "Status", NOTES_FIELD: "Notes", REQ_FIELD: "Requested",
    PKG_FIELD: "Package received", WIRE_FIELD: "Wired", "turn_bd": "Turn-time (bd)",
    "prepkg_bd": "Pre-package (bd)", NET_FIELD: "Net $", GROSS_FIELD: "Gross $",
    "Advance_Coordinator__r.Name": "Coordinator", "Loan_Advance_Number__c": "Draw #",
}


def render_pipeline(inst, tok, rt, sel, same_day):
    st.subheader("Pipeline — construction draws")
    c = st.columns([1, 1, 2])
    period = c[0].selectbox("Completed window",
                            ["This month", "Last month", "This quarter", "This year", "Last 90 days"], index=0)
    start, end = period_bounds(period)
    c[1].caption(f"{start:%m/%d/%Y} → {end:%m/%d/%Y}")

    done = add_intervals(run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' "
        f"AND {WIRE_FIELD}>={start:%Y-%m-%d} AND {WIRE_FIELD}<={end:%Y-%m-%d}"), same_day)
    flight = run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND {WIRE_FIELD}=null "
        f"AND Status__c NOT IN ({','.join(chr(39)+s+chr(39) for s in TERMINAL)})")

    # KPIs
    k = st.columns(5)
    k[0].metric("Completed (window)", f"{len(done):,}")
    if "turn_bd" in done and done["turn_bd"].notna().any():
        k[1].metric("Median turn-time", f"{done['turn_bd'].median():.0f} bd",
                    f"{(done['turn_bd']<=3).mean()*100:.0f}% ≤3 bd")
    else:
        k[1].metric("Median turn-time", "—")
    k[2].metric("Net funded", money(pd.to_numeric(done.get(NET_FIELD), errors="coerce").sum()) if NET_FIELD in done else "—")
    k[3].metric("Gross funded", money(pd.to_numeric(done.get(GROSS_FIELD), errors="coerce").sum()) if GROSS_FIELD in done else "—")
    k[4].metric("In flight", f"{len(flight):,}")

    # pre-package context + on hold
    cc = st.columns(2)
    if "prepkg_bd" in done and done["prepkg_bd"].notna().any():
        pp = done["prepkg_bd"].dropna()
        cc[0].caption(f"Pre-package (borrower/inspection/title): median {pp.median():.0f} bd · "
                      f"90th pct {pp.quantile(.9):.0f} bd — this is where the long timelines live, not funding.")
    if not flight.empty and "Status__c" in flight:
        holds = flight[flight["Status__c"].astype(str).str.contains("Hold|Pending Borrower|Revision", case=False, na=False)]
        cc[1].caption(f"On hold / needs attention: **{len(holds)}** of {len(flight)} in flight.")

    # by-month rollup
    if not done.empty and "turn_bd" in done:
        d = done.copy(); d["Month"] = pd.to_datetime(d[WIRE_FIELD]).dt.to_period("M").astype(str)
        roll = (d.groupby("Month")["turn_bd"]
                .agg(draws="size", median_bd="median", pct_le3=lambda s: round((s <= 3).mean()*100))
                .reset_index())
        st.markdown("**By month**")
        st.dataframe(roll, hide_index=True, use_container_width=True)

    # detail table
    st.markdown("**Completed draws (window)**")
    cols = [c for c in ["Loan_Number__c", "Loan_Advance_Number__c", "Deal__r.Name", "Borrower_Name_Text__c",
                        "Status__c", REQ_FIELD, PKG_FIELD, WIRE_FIELD, "turn_bd", "prepkg_bd",
                        NET_FIELD, GROSS_FIELD, NOTES_FIELD] if c in done]
    if not done.empty:
        st.dataframe(done[cols].rename(columns=_PRETTY).sort_values("Wired", ascending=False),
                     use_container_width=True, height=380)
        st.download_button("Download (CSV)", done[cols].rename(columns=_PRETTY).to_csv(index=False).encode(),
                           f"draws_{start}_{end}.csv")
    else:
        st.info("No completed draws in this window.")
    st.caption("Queried from the Advance object (not the pipeline report), so paid-off draws are retained. "
               "Turn-time = business days from full draw package received to wire; only draws with a package "
               "date are measured.")


# ───────────────────────────── Loan detail (micro) ──────────────────────────
def render_loan_detail(inst, tok, rt, sel, same_day):
    st.subheader("Loan detail")
    c = st.columns([3, 1])
    text = c[0].text_input("Search borrower, property, or loan #", placeholder="e.g. 63390  ·  116 South Street")
    mode = c[1].radio("Match on", ["Loan #", "Property", "Borrower"], label_visibility="collapsed")
    if not text:
        st.info("Search a loan to see every draw's full cycle — Salesforce milestones + the Land Gorilla draw timeline.")
        return

    esc = soql_escape(text)
    field = {"Loan #": "Loan_Number__c", "Property": "Deal__r.Name", "Borrower": "Borrower_Name_Text__c"}[mode]
    op = "=" if mode == "Loan #" and text.isdigit() else "LIKE"
    val = f"'{esc}'" if op == "=" else f"'%{esc}%'"
    df = add_intervals(run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' AND {field} {op} {val} "
        f"ORDER BY Loan_Number__c, {REQ_FIELD}"), same_day)
    if df.empty:
        st.warning("No matching construction draws."); return

    cfg = lg_config()
    for loan_no, g in df.groupby("Loan_Number__c", dropna=False):
        prop = g["Deal__r.Name"].dropna().iloc[0] if "Deal__r.Name" in g and g["Deal__r.Name"].notna().any() else ""
        borrower = g["Borrower_Name_Text__c"].dropna().iloc[0] if "Borrower_Name_Text__c" in g and g["Borrower_Name_Text__c"].notna().any() else ""
        st.markdown(f"### Loan {loan_no} — {prop}")
        if borrower:
            st.caption(f"Borrower: {borrower}  ·  Land Gorilla file: rb0{loan_no}")
        for _, row in g.iterrows():
            _render_draw(row, cfg)
        st.divider()


def _render_draw(row: pd.Series, cfg: dict | None):
    draw_no = row.get("Loan_Advance_Number__c") or row.get("Name") or "draw"
    top = st.columns([2, 1, 1, 1])
    top[0].markdown(f"**Draw {draw_no}** — {row.get('Status__c','')}")
    if NET_FIELD in row and pd.notna(row.get(NET_FIELD)):
        top[1].metric("Net", money(row[NET_FIELD]))
    if GROSS_FIELD in row and pd.notna(row.get(GROSS_FIELD)):
        top[2].metric("Gross", money(row[GROSS_FIELD]))
    if "turn_bd" in row and pd.notna(row.get("turn_bd")):
        top[3].metric("Turn-time", f"{row['turn_bd']:.0f} bd")

    # Salesforce milestone timeline
    steps = [{"Milestone": lbl, "Date": (pd.to_datetime(row.get(f)).date() if pd.notna(row.get(f)) else None),
              "": "✅" if pd.notna(row.get(f)) else "⬜"} for lbl, f in MILESTONES]
    tdf = pd.DataFrame(steps)
    recorded = tdf["Date"].notna().sum()
    st.progress(recorded / len(MILESTONES), text=f"Salesforce milestones {recorded}/{len(MILESTONES)}")
    cc = st.columns([1, 1])
    cc[0].dataframe(tdf, hide_index=True, use_container_width=True,
                    column_config={"": st.column_config.TextColumn(width="small")})

    # Land Gorilla per-draw timeline via DrawContainerId__c
    with cc[1]:
        st.markdown("**Land Gorilla draw**")
        container = row.get(CONTAINER_FIELD)
        if not cfg:
            st.caption("Add [landgorilla] secrets to show the LG draw timeline.")
        elif pd.isna(container) or not container:
            st.caption("No Land Gorilla draw container on this advance yet.")
        else:
            det = lg_draw_detail(cfg["user"], cfg["password"], bool(cfg.get("verify", True)), str(container))
            if not det or det.get("_error"):
                st.caption(f"LG draw detail unavailable ({(det or {}).get('_error','no data')}).")
            else:
                lg_steps = [("Created", det["created"]), ("Submitted", det["submitted"]),
                            ("Approved", det["approved"]), ("Funded", det["funded"])]
                st.dataframe(pd.DataFrame([{"Stage": s, "Date": d} for s, d in lg_steps]),
                             hide_index=True, use_container_width=True)
                line = [f"Status: {det.get('status','—')}"]
                if det.get("amount") is not None:
                    line.append(f"Amount (less retainage): {money(det['amount'])}")
                st.caption(" · ".join(line))

    note = row.get(NOTES_FIELD)
    if pd.notna(note) and str(note).strip():
        st.caption(f"📝 {note}")


# ───────────────────────────────── main ─────────────────────────────────────
def main():
    st.set_page_config(page_title="Construction Draw Tracker", page_icon="🏗️", layout="wide")
    st.title("🏗️ Construction Draw Tracker")

    cfg = None; err = None
    try:
        cfg = load_sf_oauth(); finish_oauth(cfg)
    except Exception as exc:
        err = str(exc)
    sf = None if err else sf_from_session()

    with st.sidebar:
        st.header("Salesforce")
        if err:
            st.error(err)
        elif sf is None:
            st.info("Not connected")
        else:
            st.success("Connected")
            st.caption(st.session_state.get("salesforce_auth", {}).get("instance_url", ""))
            if st.button("Log out", use_container_width=True):
                clear_sf_session(); st.rerun()
        st.divider()
        same_day = 1 if st.radio("Same-day convention", ["0 business days", "1 business day"],
                                 help="Package in & wired same day counts as this. Confirm with Melanie — "
                                      "it moves the median.").startswith("1") else 0
        st.divider()
        st.caption("Land Gorilla: " + ("configured ✅" if lg_config() else "not set"))

    st.subheader("Step 1 — Log in to Salesforce")
    if err:
        st.error(err); st.stop()
    if sf is None:
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
            clear_sf_session(); st.warning("Your Salesforce session expired. Log in again."); st.stop()
        raise


if __name__ == "__main__":
    main()
