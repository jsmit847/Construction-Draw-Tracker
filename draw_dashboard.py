"""
Construction Draw Dashboard  (Streamlit)
========================================
Same Salesforce OAuth login as the AM slide app, then two views:

  • Pipeline Pulse  — what's happening now: open draws by stage, on-hold draws,
                      completed this period ($ and count), recent wires, aging,
                      and the median business-day turn-time (package -> wire).
  • Draw Lookup     — type a property/deal or advance # and see the full draw
                      cycle for each matching advance: milestone timeline,
                      status, amounts, and the two turn-time intervals.

Built on the Advance__c model we mapped:
  scope   = Record Type "Construction Advance"
  anchor  = Date_Submitted_to_Capital_Partner__c  ("Date Full Draw Package Received")
  wire    = Wire_Date__c
The SELECTs are built from describe(), so a field that doesn't exist in the org
is skipped instead of breaking the query.

Secrets (same as the reference app), in .streamlit/secrets.toml:
  [salesforce]
  client_id     = "..."
  client_secret = "..."
  redirect_uri  = "https://<your-app-host>/"
  auth_host     = "https://login.salesforce.com"   # or https://test.salesforce.com
  scope         = "api refresh_token"
  prompt        = "login"

Run:  streamlit run draw_dashboard.py
"""
from __future__ import annotations

import base64
import hashlib
import json
import secrets
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
PKG_FIELD  = "Date_Submitted_to_Capital_Partner__c"     # "Date Full Draw Package Received"
WIRE_FIELD = "Wire_Date__c"
REQ_FIELD  = "Date_Advance_Requested__c"

# Milestone chain, in order, for the "full draw cycle" timeline.
MILESTONES = [
    ("Requested",                 REQ_FIELD),
    ("Inspection ordered",        "Date_Inspection_Ordered__c"),
    ("Inspection",                "Date_Of_Inspection__c"),
    ("Inspection report received","Date_Inspection_Report_Received__c"),
    ("Submitted for review",      "Date_Submitted_For_Approval__c"),
    ("Internal review complete",  "Date_Internal_Review_Complete__c"),
    ("Full draw package received",PKG_FIELD),
    ("Manager approval",          "Manager_Approval_Date__c"),
    ("Wired",                     WIRE_FIELD),
]
# Fields we'd like if the org has them (intersected with describe()).
WISH_TEXT = ["Name", "Deal__r.Name", "Lender__c", "Status__c", "IC_Approval_Status__c",
             "Exception__c", "Cancellation_Reason__c", "Inspection_Method__c",
             "Advance_Coordinator__r.Name", "Advance_Analyst__r.Name",
             "Underwriter__r.Name", "Advance_Requestor__r.Name"]
WISH_AMOUNT = ["Net_Funded_Amount__c", "Current_Draw_Amount__c", "Draw_Amount__c",
               "Advance_Amount__c", "Amount__c"]
OPEN_EXCLUDE_STATUS = ["Completed", "Cancelled", "Rescinded", "Rejected by Capital Partner"]

_HOLS = USFederalHolidayCalendar().holidays("2018-01-01", "2032-12-31").values.astype("datetime64[D]")


# ───────────────────────────── Salesforce OAuth ─────────────────────────────
# (Faithful to the AM slide app so it uses the same secrets + redirect URI.)
def install_truststore() -> None:
    try:
        import truststore
        truststore.inject_into_ssl()
    except Exception:
        pass


def load_salesforce_oauth_config() -> dict[str, str]:
    section = dict(st.secrets.get("salesforce", {}))
    required = ["client_id", "client_secret", "redirect_uri", "auth_host"]
    missing = [k for k in required if not section.get(k)]
    if missing:
        raise RuntimeError(
            "Missing Salesforce OAuth secrets: " + ", ".join(missing)
            + ". Add them under [salesforce] in Streamlit secrets."
        )
    section.setdefault("scope", "api refresh_token")
    section.setdefault("prompt", "login")
    return section


@st.cache_resource
def _pkce_store() -> dict:
    return {}


def generate_pkce_pair() -> tuple[str, str]:
    verifier = base64.urlsafe_b64encode(secrets.token_bytes(64)).rstrip(b"=").decode("ascii")
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")
    return verifier, challenge


def build_salesforce_login_url(cfg: dict[str, str]) -> str:
    auth_host = str(cfg["auth_host"]).rstrip("/")
    state = secrets.token_urlsafe(24)
    verifier, challenge = generate_pkce_pair()
    store = _pkce_store()
    store[state] = verifier
    if len(store) > 50:
        for old in list(store.keys())[:-50]:
            store.pop(old, None)
    query = urlencode({
        "response_type": "code",
        "client_id": cfg["client_id"],
        "redirect_uri": cfg["redirect_uri"],
        "scope": cfg.get("scope", "api refresh_token"),
        "prompt": cfg.get("prompt", "login"),
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
    })
    return f"{auth_host}/services/oauth2/authorize?{query}"


def exchange_code_for_token(cfg: dict[str, str], code: str, verifier: str | None) -> dict[str, Any]:
    install_truststore()
    token_url = f"{str(cfg['auth_host']).rstrip('/')}/services/oauth2/token"
    fields = {
        "grant_type": "authorization_code",
        "client_id": cfg["client_id"],
        "client_secret": cfg["client_secret"],
        "redirect_uri": cfg["redirect_uri"],
        "code": code,
    }
    if verifier:
        fields["code_verifier"] = verifier
    req = Request(token_url, data=urlencode(fields).encode(),
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


def _qp(name: str) -> str | None:
    v = st.query_params.get(name)
    return v[0] if isinstance(v, list) else v


def clear_salesforce_session() -> None:
    for k in ["salesforce_auth", "_last_sf_code"]:
        st.session_state.pop(k, None)


def maybe_finish_oauth(cfg: dict[str, str]) -> None:
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
    verifier = _pkce_store().pop(state, None) if state else None
    if state and not verifier:
        st.query_params.clear()
        raise RuntimeError("Login could not be completed (PKCE verifier missing — the app likely "
                           "restarted between steps). Click 'Log in to Salesforce' and try again.")
    payload = exchange_code_for_token(cfg, code, verifier)
    access_token = payload.get("access_token")
    instance_url = payload.get("instance_url")
    if not access_token or not instance_url:
        raise RuntimeError("Login succeeded but no access token / instance URL was returned.")
    st.session_state["salesforce_auth"] = {"access_token": access_token, "instance_url": instance_url}
    st.session_state["_last_sf_code"] = code
    st.query_params.clear()
    st.rerun()


def get_sf_from_session() -> Salesforce | None:
    install_truststore()
    auth = st.session_state.get("salesforce_auth", {})
    if not auth.get("instance_url") or not auth.get("access_token"):
        return None
    return Salesforce(instance_url=auth["instance_url"], session_id=auth["access_token"])


# ───────────────────────────── query helpers ────────────────────────────────
def soql_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("'", "\\'")


def flatten(rec: dict) -> dict:
    """Flatten one/two levels of relationship dicts to dotted keys; drop 'attributes'."""
    out: dict[str, Any] = {}
    for k, v in rec.items():
        if k == "attributes":
            continue
        if isinstance(v, dict):
            for k2, v2 in v.items():
                if k2 == "attributes":
                    continue
                out[f"{k}.{k2}"] = v2.get("Name") if isinstance(v2, dict) else v2
        else:
            out[k] = v
    return out


@st.cache_data(ttl=900, show_spinner=False)
def describe_fields(instance_url: str, token: str) -> list[str]:
    sf = Salesforce(instance_url=instance_url, session_id=token)
    return [f["name"] for f in sf.Advance__c.describe()["fields"]]


@st.cache_data(ttl=900, show_spinner=False)
def construction_rt_id(instance_url: str, token: str) -> str | None:
    sf = Salesforce(instance_url=instance_url, session_id=token)
    for r in sf.query("SELECT Id,DeveloperName FROM RecordType "
                      "WHERE SobjectType='Advance__c'")["records"]:
        if r["DeveloperName"] == CONSTRUCTION_DEV_NAME:
            return r["Id"]
    return None


@st.cache_data(ttl=300, show_spinner=False)
def run_soql(instance_url: str, token: str, soql: str) -> pd.DataFrame:
    sf = Salesforce(instance_url=instance_url, session_id=token)
    recs = sf.query_all(soql)["records"]
    return pd.DataFrame([flatten(r) for r in recs])


def available(instance_url: str, token: str) -> list[str]:
    present = set(describe_fields(instance_url, token))
    fields = ["Id"]
    for f in WISH_TEXT + [m[1] for m in MILESTONES] + ["Target_Advance_Date__c"]:
        base = f.split(".")[0]
        if base in present and f not in fields:
            fields.append(f)
    amt = next((a for a in WISH_AMOUNT if a in present), None)
    if amt:
        fields.append(amt)
    return fields, amt


# ───────────────────────────── turn-time math ───────────────────────────────
def bdays(a: pd.Series, b: pd.Series, same_day_as: int) -> pd.Series:
    a, b = pd.to_datetime(a, errors="coerce"), pd.to_datetime(b, errors="coerce")
    m = a.notna() & b.notna()
    out = pd.Series(np.nan, index=a.index)
    if m.any():
        out[m] = np.busday_count(a[m].values.astype("datetime64[D]"),
                                 b[m].values.astype("datetime64[D]"), holidays=_HOLS) + same_day_as
    out[out < 0] = np.nan
    return out


def add_intervals(df: pd.DataFrame, same_day_as: int) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    if PKG_FIELD in df and WIRE_FIELD in df:
        df["turn_bd"] = bdays(df[PKG_FIELD], df[WIRE_FIELD], same_day_as)
    if REQ_FIELD in df and PKG_FIELD in df:
        df["prepkg_bd"] = bdays(df[REQ_FIELD], df[PKG_FIELD], same_day_as)
    if REQ_FIELD in df and WIRE_FIELD in df:
        df["total_bd"] = bdays(df[REQ_FIELD], df[WIRE_FIELD], same_day_as)
    return df


def period_bounds(choice: str) -> tuple[date, date]:
    today = date.today()
    if choice == "This week":
        start = today - pd.Timedelta(days=today.weekday()); return start, today
    if choice == "This month":
        return today.replace(day=1), today
    if choice == "This quarter":
        q = (today.month - 1) // 3
        return date(today.year, q * 3 + 1, 1), today
    if choice == "This year":
        return date(today.year, 1, 1), today
    if choice == "Last 90 days":
        return today - pd.Timedelta(days=90), today
    return date(today.year, 1, 1), today


def money(x) -> str:
    try:
        return f"${x:,.0f}"
    except Exception:
        return "—"


# ───────────────────────────── UI: Pipeline Pulse ───────────────────────────
def render_pulse(inst: str, tok: str, rt: str, fields: list[str], amt: str | None, same_day: int):
    st.subheader("Pipeline Pulse")
    colf = st.columns([1, 1, 2])
    period = colf[0].selectbox("Completed window",
                               ["This week", "This month", "This quarter", "This year", "Last 90 days"],
                               index=1)
    start, end = period_bounds(period)
    colf[1].caption(f"{start:%m/%d/%Y} → {end:%m/%d/%Y}")

    sel = ",".join(fields)

    # --- open (in-flight) draws: not wired, not terminal ---
    open_status = "(" + ",".join(f"'{s}'" for s in OPEN_EXCLUDE_STATUS) + ")"
    open_df = run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' "
        f"AND {WIRE_FIELD}=null AND Status__c NOT IN {open_status}")
    # --- completed in window ---
    done_df = run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' "
        f"AND {WIRE_FIELD}>={start:%Y-%m-%d} AND {WIRE_FIELD}<={end:%Y-%m-%d}")
    done_df = add_intervals(done_df, same_day)

    # --- KPI row ---
    k = st.columns(4)
    k[0].metric("Open draws (in flight)", f"{len(open_df):,}")
    k[1].metric(f"Completed ({period.lower()})", f"{len(done_df):,}")
    if amt and amt in done_df:
        k[2].metric("Funded in window", money(pd.to_numeric(done_df[amt], errors="coerce").sum()))
    else:
        k[2].metric("Funded in window", "—")
    if "turn_bd" in done_df and done_df["turn_bd"].notna().any():
        med = done_df["turn_bd"].median()
        within3 = (done_df["turn_bd"] <= 3).mean() * 100
        k[3].metric("Median turn-time", f"{med:.0f} bd", f"{within3:.0f}% ≤3 bd")
    else:
        k[3].metric("Median turn-time", "—", "no package dates in window")

    st.divider()

    # --- open by stage + on hold ---
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Open draws by stage**")
        if not open_df.empty and "Status__c" in open_df:
            by = open_df["Status__c"].value_counts().rename_axis("Status").reset_index(name="Draws")
            st.bar_chart(by.set_index("Status"))
            if amt and amt in open_df:
                pend = pd.to_numeric(open_df[amt], errors="coerce").sum()
                st.caption(f"Estimated open exposure: {money(pend)}")
        else:
            st.info("No open draws.")
    with c2:
        st.markdown("**On hold / needs attention**")
        holds = open_df[open_df.get("Status__c", pd.Series(dtype=str)).astype(str)
                        .str.contains("Hold|Pending Borrower|Revision", case=False, na=False)] \
                if not open_df.empty else open_df
        if not holds.empty:
            show = [c for c in ["Name", "Deal__r.Name", "Status__c", "Advance_Coordinator__r.Name"] if c in holds]
            st.dataframe(holds[show].rename(columns=_pretty), use_container_width=True, height=240)
        else:
            st.success("Nothing sitting on hold.")

    st.divider()

    # --- aging of open draws ---
    st.markdown("**Aging — oldest open draws (calendar days since requested)**")
    if not open_df.empty and REQ_FIELD in open_df:
        aged = open_df.copy()
        aged["Days open"] = (pd.Timestamp(date.today()) - pd.to_datetime(aged[REQ_FIELD], errors="coerce")).dt.days
        aged = aged.sort_values("Days open", ascending=False)
        show = [c for c in ["Name", "Deal__r.Name", "Status__c", "Days open"] if c in aged]
        st.dataframe(aged[show].head(15).rename(columns=_pretty), use_container_width=True, height=300)
    else:
        st.info("No requested-date data on open draws.")

    st.divider()

    # --- recent wires ---
    st.markdown("**Recently completed (last 15 wires in window)**")
    if not done_df.empty:
        rc = done_df.sort_values(WIRE_FIELD, ascending=False).head(15)
        show = [c for c in ["Name", "Deal__r.Name", WIRE_FIELD, amt, "turn_bd"] if c and c in rc]
        st.dataframe(rc[show].rename(columns=_pretty), use_container_width=True, height=300)
        st.download_button("Download completed (window) as CSV",
                           done_df.to_csv(index=False).encode(), f"completed_{start}_{end}.csv")
    else:
        st.info("No completed draws in this window.")

    if "turn_bd" in done_df:
        st.caption("Turn-time = business days from full draw package received to wire. Only draws "
                   "with a recorded package date are measured — coverage in Salesforce is partial; "
                   "Land Gorilla (IHD-109768) fills the rest.")


# ───────────────────────────── UI: Draw Lookup ──────────────────────────────
_pretty = {
    "Name": "Advance #", "Deal__r.Name": "Property / Deal", "Status__c": "Status",
    "Lender__c": "Lender", "Advance_Coordinator__r.Name": "Coordinator",
    "Advance_Analyst__r.Name": "Analyst", "Underwriter__r.Name": "Underwriter",
    "Advance_Requestor__r.Name": "Requestor", WIRE_FIELD: "Wire date",
    REQ_FIELD: "Requested", PKG_FIELD: "Full package received",
    "turn_bd": "Turn-time (bd)", "prepkg_bd": "Pre-package (bd)", "total_bd": "Total (bd)",
}


def render_lookup(inst: str, tok: str, rt: str, fields: list[str], amt: str | None, same_day: int):
    st.subheader("Draw Lookup")
    c = st.columns([3, 1])
    text = c[0].text_input("Search by property / deal name or advance #", placeholder="e.g. 745 South 9th Street")
    mode = c[1].radio("Match on", ["Property / Deal", "Advance #"], label_visibility="collapsed")
    if not text:
        st.info("Type a property/deal name or an advance number to see its full draw cycle.")
        return

    esc = soql_escape(text)
    field = "Deal__r.Name" if mode.startswith("Property") else "Name"
    sel = ",".join(fields)
    df = run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' "
        f"AND {field} LIKE '%{esc}%' ORDER BY {REQ_FIELD} DESC NULLS LAST")
    df = add_intervals(df, same_day)
    if df.empty:
        st.warning("No matching construction advances.")
        return

    st.caption(f"{len(df)} matching advance(s).")
    # group by deal so multiple draws on one property read as a cycle
    deal_col = "Deal__r.Name" if "Deal__r.Name" in df else "Name"
    for deal, g in df.groupby(deal_col, dropna=False):
        with st.expander(f"{deal}  ·  {len(g)} draw(s)", expanded=(len(df) <= 5)):
            for _, row in g.iterrows():
                _render_one_draw(row, amt)


def _render_one_draw(row: pd.Series, amt: str | None):
    top = st.columns([2, 1, 1])
    top[0].markdown(f"**{row.get('Name','(advance)')}** — {row.get('Status__c','')}")
    if amt and amt in row and pd.notna(row[amt]):
        top[1].metric("Amount", money(row[amt]))
    if "turn_bd" in row and pd.notna(row.get("turn_bd")):
        top[2].metric("Turn-time", f"{row['turn_bd']:.0f} bd")

    # milestone timeline
    steps = []
    for label, f in MILESTONES:
        val = row.get(f)
        steps.append({"Milestone": label,
                      "Date": pd.to_datetime(val).date() if pd.notna(val) else None,
                      "": "✅" if pd.notna(val) else "⬜"})
    tdf = pd.DataFrame(steps)
    done = tdf["Date"].notna().sum()
    st.progress(done / len(MILESTONES), text=f"{done}/{len(MILESTONES)} milestones recorded")
    st.dataframe(tdf, hide_index=True, use_container_width=True,
                 column_config={"": st.column_config.TextColumn(width="small")})

    meta = []
    for f in ["Lender__c", "Advance_Coordinator__r.Name", "Advance_Analyst__r.Name",
              "Underwriter__r.Name", "Advance_Requestor__r.Name"]:
        if f in row and pd.notna(row[f]):
            meta.append(f"**{_pretty.get(f, f)}:** {row[f]}")
    if "prepkg_bd" in row and pd.notna(row.get("prepkg_bd")):
        meta.append(f"**Pre-package:** {row['prepkg_bd']:.0f} bd (borrower/inspection/title)")
    if meta:
        st.caption("  ·  ".join(meta))
    st.divider()


# ───────────────────────────────── main ─────────────────────────────────────
def main():
    st.set_page_config(page_title="Construction Draw Dashboard", page_icon="🏗️", layout="wide")
    st.title("🏗️ Construction Draw Dashboard")

    cfg = None
    setup_error = None
    try:
        cfg = load_salesforce_oauth_config()
        maybe_finish_oauth(cfg)
    except Exception as exc:
        setup_error = str(exc)

    sf = None if setup_error else get_sf_from_session()

    with st.sidebar:
        st.header("Salesforce")
        if setup_error:
            st.error(setup_error)
        elif sf is None:
            st.info("Not connected")
        else:
            st.success("Connected")
            st.caption(st.session_state.get("salesforce_auth", {}).get("instance_url", ""))
            if st.button("Log out", use_container_width=True):
                clear_salesforce_session(); st.rerun()
        st.divider()
        same_day = 1 if st.radio(
            "Same-day convention",
            ["0 business days", "1 business day"],
            help="Package in and wired the same day counts as this. Confirm with Melanie — it moves the median."
        ).startswith("1") else 0

    # login gate
    st.subheader("Step 1 — Log in to Salesforce")
    if setup_error:
        st.error(setup_error); st.stop()
    if sf is None:
        st.info("Log in to load the construction draw pipeline.")
        st.link_button("Log in to Salesforce", build_salesforce_login_url(cfg))
        st.caption(f"Callback URL: {cfg['redirect_uri']}")
        st.stop()

    inst = st.session_state["salesforce_auth"]["instance_url"]
    tok = st.session_state["salesforce_auth"]["access_token"]

    rt = construction_rt_id(inst, tok)
    if not rt:
        st.error("Could not find the 'Construction Advance' record type on Advance__c."); st.stop()
    fields, amt = available(inst, tok)

    page = st.sidebar.radio("View", ["Pipeline Pulse", "Draw Lookup"])
    try:
        if page == "Pipeline Pulse":
            render_pulse(inst, tok, rt, fields, amt, same_day)
        else:
            render_lookup(inst, tok, rt, fields, amt, same_day)
    except Exception as exc:
        msg = str(exc)
        if "INVALID_SESSION_ID" in msg or "Session expired" in msg:
            clear_salesforce_session()
            st.warning("Your Salesforce session expired. Log in again.")
            st.stop()
        raise


if __name__ == "__main__":
    main()
