"""
Construction Draw Dashboard  (Streamlit)
========================================
Same Salesforce OAuth login as the AM slide app, then three views:

  • Pipeline Pulse    — what's happening now: open draws by stage, on-hold draws,
                        completed this period ($ and count), recent wires, aging,
                        and the median business-day turn-time (package -> wire).
                        Open pipeline downloads to Excel for the Thursday call.
  • Turn-Time Report  — on-demand, any period (last month / quarter / year or a
                        custom range): every draw wired in the window with its
                        milestone dates, wire date, construction manager, notes,
                        stage-by-stage business days, and the official turn-time
                        (full draw package received -> wire). One-click Excel
                        workbook: Summary, Draw Detail, By Month, By Manager,
                        Definitions.
  • Draw Lookup       — type a property/deal or advance # and see the full draw
                        cycle for each matching advance: milestone timeline,
                        status, amounts, and the two turn-time intervals.

Built on the Advance__c model we mapped:
  scope   = Record Type "Construction Advance"
  anchor  = Date_Submitted_to_Capital_Partner__c  ("Date Full Draw Package Received")
  wire    = Wire_Date__c
The SELECTs are built from describe(), so a field that doesn't exist in the org
is skipped instead of breaking the query. Deal and Property API names come from
the org's column glossary (Column_objects_cleaned.docx); Advance__c isn't in it,
so construction manager, hold reason and notes are still found by label there.

Draws are selected straight off Advance__c by wire date, with no filter on the
loan's status — so a draw funded one day and paid off the next still shows up
(the gap in the existing Salesforce pipeline report).

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
import io
import json
import secrets
from datetime import date, datetime, timedelta, timezone
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
PKG_FIELD  = "Date_Submitted_to_Capital_Partner__c"     # "Date full draw package received"
WIRE_FIELD = "Wire_Date__c"
REQ_FIELD  = "Date_Advance_Requested__c"
INSP_FIELD = "Date_Of_Inspection__c"
RPT_FIELD  = "Date_Inspection_Report_Received__c"
BUSINESS_TZ = "America/Los_Angeles"   # datetime fields are converted to this before taking the date

# Milestone chain, in order, for the "full draw cycle" timeline.
MILESTONES = [
    ("Requested",                 REQ_FIELD),
    ("Inspection ordered",        "Date_Inspection_Ordered__c"),
    ("Inspection",                INSP_FIELD),
    ("Inspection report received",RPT_FIELD),
    ("Submitted for review",      "Date_Submitted_For_Approval__c"),
    ("Internal review complete",  "Date_Internal_Review_Complete__c"),
    ("Full draw package received",PKG_FIELD),
    ("Manager approval",          "Manager_Approval_Date__c"),
    ("Wired",                     WIRE_FIELD),
]
# Stage-by-stage business days (plain business-day counts, so they add up) —
# shows where the time went before the package was complete.
STAGES = [
    ("stage_req_insp", "Requested → Inspection (bd)",           REQ_FIELD, INSP_FIELD),
    ("stage_insp_rpt", "Inspection → Report received (bd)",     INSP_FIELD, RPT_FIELD),
    ("stage_rpt_pkg",  "Report received → Full package (bd)",   RPT_FIELD, PKG_FIELD),
]
# Fields we'd like if the org has them (intersected with describe()).
WISH_TEXT = ["Name", "Deal__r.Name", "Lender__c", "Status__c", "IC_Approval_Status__c",
             "Exception__c", "Cancellation_Reason__c", "Inspection_Method__c",
             "Advance_Coordinator__r.Name", "Advance_Analyst__r.Name",
             "Underwriter__r.Name", "Advance_Requestor__r.Name"]
WISH_AMOUNT = ["Net_Funded_Amount__c", "Current_Draw_Amount__c", "Draw_Amount__c",
               "Advance_Amount__c", "Amount__c"]
# Deal (Opportunity) fields, API names from the org's column glossary
# (Column_objects_cleaned.docx). role -> (candidate paths, first that exists wins; heading)
DEAL_WISH = {
    "loan_number":    (["Deal__r.Deal_Loan_Number__c"], "Loan #"),
    "loan_manager":   (["Deal__r.Loan_Coordinator__r.Name", "Deal__r.Loan_Coordinator__c"], "Loan manager"),
    "product_type":   (["Deal__r.LOC_Loan_Type__c"], "Product type"),
    "product_sub":    (["Deal__r.Product_Sub_Type__c"], "Product sub-type"),
    "servicer_status":(["Deal__r.Servicer_Status__c"], "Servicer loan status"),
    "payoff_date":    (["Deal__r.Payoff_Date__c"], "Loan payoff date"),
    "deal_comments":  (["Deal__r.Construction_Comments__c"], "Construction comments (deal)"),
}
# Property__c fields rolled up per deal (a deal can have several properties).
PROPERTY_WISH = {
    "ConstructionManagementLoanId__c": "Land Gorilla loan ID",
    "Servicer_Loan_Number__c":         "Servicer loan #",
}
# Advance__c isn't in the glossary, so these are still found by label on Advance__c.
#   key -> (label keywords, allowed field types, max fields, column heading)
_TEXTY = ("string", "textarea", "picklist", "multipicklist")
DISCOVER = {
    "construction_manager": (["construction manager"], _TEXTY + ("reference",), 1, "Construction manager"),
    "hold_reason":          (["hold reason", "on hold", "delay reason", "pending reason"], _TEXTY, 2, None),
    "notes":                (["note", "comment"], ("string", "textarea"), 3, None),
}
OPEN_EXCLUDE_STATUS = ["Completed", "Cancelled", "Rescinded", "Rejected by Capital Partner"]

PERIODS = ["Last month", "Last quarter", "Last year", "Month to date", "Quarter to date",
           "Year to date", "Last 90 days", "Custom range"]

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


def flatten(rec: dict, prefix: str = "") -> dict:
    """Flatten relationship dicts (any depth) to dotted keys; drop 'attributes'."""
    out: dict[str, Any] = {}
    for k, v in rec.items():
        if k == "attributes":
            continue
        if isinstance(v, dict):
            out.update(flatten(v, f"{prefix}{k}."))
        else:
            out[f"{prefix}{k}"] = v
    return out


@st.cache_data(ttl=900, show_spinner=False)
def describe_object(instance_url: str, token: str, sobject: str) -> list[dict]:
    sf = Salesforce(instance_url=instance_url, session_id=token)
    return [{"name": f["name"], "label": f["label"], "type": f["type"],
             "rel": f.get("relationshipName"), "ref": (f.get("referenceTo") or [None])[0]}
            for f in getattr(sf, sobject).describe()["fields"]]


def _describe(inst: str, tok: str, sobject: str) -> list[dict]:
    try:
        return describe_object(inst, tok, sobject)
    except Exception as exc:
        if "INVALID_SESSION_ID" in str(exc):
            raise
        return []


def field_exists(inst: str, tok: str, path: str, sobject: str = "Advance__c") -> bool:
    """True if a (possibly dotted, e.g. Deal__r.Name) field path exists in the org."""
    head, _, rest = path.partition(".")
    fields = _describe(inst, tok, sobject)
    if not rest:
        return any(f["name"] == head for f in fields)
    rel = next((f for f in fields if f["rel"] == head and f["ref"]), None)
    return bool(rel) and field_exists(inst, tok, rest, rel["ref"])


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


class Schema:
    """What this org actually has: SELECT list, amount field, discovered fields."""

    def __init__(self, fields: list[str], amt: str | None,
                 roles: dict[str, list[str]], labels: dict[str, str]):
        self.fields, self.amt, self.roles, self.labels = fields, amt, roles, labels

    @property
    def select(self) -> str:
        return ",".join(self.fields)

    def role(self, key: str) -> str | None:
        hits = self.roles.get(key) or []
        return hits[0] if hits else None

    def label(self, col: str) -> str:
        return _pretty.get(col) or self.labels.get(col) or col


def _discover(inst: str, tok: str) -> tuple[dict[str, list[str]], dict[str, str]]:
    roles: dict[str, list[str]] = {}
    labels: dict[str, str] = {}
    for key, (words, types, limit, heading) in DISCOVER.items():
        hits: list[str] = []
        for f in _describe(inst, tok, "Advance__c"):
            if f["type"] in types and any(w in f["label"].lower() for w in words):
                path = f"{f['rel']}.Name" if f["type"] == "reference" and f["rel"] else f["name"]
                if path not in hits:
                    hits.append(path)
                    labels[path] = heading or f["label"]
        roles[key] = hits[:limit]
    for key, (candidates, heading) in DEAL_WISH.items():
        path = next((p for p in candidates if field_exists(inst, tok, p)), None)
        roles[key] = [path] if path else []
        if path:
            labels[path] = heading
    return roles, labels


def attach_property_ids(inst: str, tok: str, df: pd.DataFrame) -> pd.DataFrame:
    """Add Land Gorilla loan ID / servicer loan # from Property__c, joined per deal."""
    if df.empty or "Deal__c" not in df:
        return df
    fields = [f for f in PROPERTY_WISH if field_exists(inst, tok, f, "Property__c")]
    if not fields or not field_exists(inst, tok, "Deal__c", "Property__c"):
        return df
    deal_ids = sorted(df["Deal__c"].dropna().astype(str).unique())
    frames = []
    for i in range(0, len(deal_ids), 200):
        ids = ",".join(f"'{soql_escape(x)}'" for x in deal_ids[i:i + 200])
        frames.append(run_soql(inst, tok, f"SELECT Deal__c,{','.join(fields)} FROM Property__c "
                                          f"WHERE Deal__c IN ({ids})"))
    props = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if props.empty:
        return df
    joined = props.groupby("Deal__c")[[f for f in fields if f in props]].agg(
        lambda s: "; ".join(dict.fromkeys(s.dropna().astype(str))) or None)
    joined.columns = [f"Property.{c}" for c in joined.columns]
    return df.merge(joined, left_on="Deal__c", right_index=True, how="left")


def load_schema(inst: str, tok: str) -> Schema:
    present = {f["name"] for f in _describe(inst, tok, "Advance__c")}
    labels = {f["name"]: f["label"] for f in _describe(inst, tok, "Advance__c")}
    labels.update({f"Property.{k}": v for k, v in PROPERTY_WISH.items()})
    fields = ["Id"]
    for f in ["Deal__c"] + WISH_TEXT + [m[1] for m in MILESTONES] + ["Target_Advance_Date__c"]:
        if f not in fields and field_exists(inst, tok, f):
            fields.append(f)
    amt = next((a for a in WISH_AMOUNT if a in present), None)
    if amt:
        fields.append(amt)
    roles, found_labels = _discover(inst, tok)
    labels.update(found_labels)
    for paths in roles.values():
        fields += [p for p in paths if p not in fields]
    return Schema(fields, amt, roles, labels)


def id_cols(schema: Schema) -> list[str | None]:
    """Columns that identify a draw: advance, property, loan numbers, product."""
    return ["Name", "Deal__r.Name", schema.role("loan_number"),
            *[f"Property.{k}" for k in PROPERTY_WISH],
            schema.role("product_type"), schema.role("product_sub"), "Lender__c", "Status__c"]


def people_cols(schema: Schema) -> list[str | None]:
    return [schema.role("construction_manager"), schema.role("loan_manager"),
            "Advance_Coordinator__r.Name", "Advance_Analyst__r.Name"]


def note_cols(schema: Schema) -> list[str | None]:
    return [*schema.roles.get("hold_reason", []), "Exception__c", *schema.roles.get("notes", []),
            schema.role("deal_comments")]


# ───────────────────────────── turn-time math ───────────────────────────────
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
    for f in [m[1] for m in MILESTONES] + [DEAL_WISH["payoff_date"][0][0]]:
        if f in df:
            df[f] = to_date(df[f])
    if PKG_FIELD in df and WIRE_FIELD in df:
        df["turn_bd"] = bdays(df[PKG_FIELD], df[WIRE_FIELD], same_day_as)
    if REQ_FIELD in df and PKG_FIELD in df:
        df["prepkg_bd"] = bdays(df[REQ_FIELD], df[PKG_FIELD], same_day_as)
    if REQ_FIELD in df and WIRE_FIELD in df:
        df["total_bd"] = bdays(df[REQ_FIELD], df[WIRE_FIELD], same_day_as)
    for key, _, a, b in STAGES:
        if a in df and b in df:
            df[key] = bdays(df[a], df[b], 0)
    return df


def period_bounds(choice: str, today: date | None = None) -> tuple[date, date]:
    today = today or date.today()
    month_start = today.replace(day=1)
    q_start = date(today.year, 3 * ((today.month - 1) // 3) + 1, 1)
    if choice == "Week to date":
        return today - timedelta(days=today.weekday()), today
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
    raise ValueError(f"Unknown period: {choice}")


def period_picker(col, key: str, options: list[str], index: int = 0) -> tuple[date, date, str]:
    choice = col.selectbox("Wire date window", options, index=index, key=f"{key}_period")
    if choice == "Custom range":
        d = st.columns(2)
        start = d[0].date_input("From", date.today().replace(day=1), key=f"{key}_from")
        end = d[1].date_input("To", date.today(), key=f"{key}_to")
        if start > end:
            st.error("'From' is after 'To'."); st.stop()
        return start, end, f"{start:%m/%d/%Y}–{end:%m/%d/%Y}"
    start, end = period_bounds(choice)
    return start, end, choice


def money(x) -> str:
    try:
        return f"${x:,.0f}"
    except Exception:
        return "—"


# ───────────────────────────── Excel export ─────────────────────────────────
def build_workbook(sheets: dict[str, pd.DataFrame], money_cols: tuple[str, ...] = ()) -> bytes:
    """Write DataFrames to an .xlsx with bold frozen headers, filters, widths, date formats."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
            ws = xw.sheets[name[:31]]
            ws.freeze_panes = "A2"
            if len(df):
                ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            for i, col in enumerate(df.columns, start=1):
                letter = get_column_letter(i)
                series = df[col]
                is_date = pd.api.types.is_datetime64_any_dtype(series)
                longest = max([len(str(col))] + [len(str(v)) for v in series.head(500) if pd.notna(v)])
                width = 12 if is_date else min(max(10, longest + 2), 60)
                ws.column_dimensions[letter].width = width
                fmt = ("mm/dd/yyyy" if is_date else
                       "$#,##0" if col in money_cols else
                       "0.0%" if str(col).startswith("%") else None)
                wrap = longest > 60
                if fmt or wrap:
                    for cell in ws[letter][1:]:
                        if fmt:
                            cell.number_format = fmt
                        if wrap:
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
    return buf.getvalue()


# ───────────────────────────── UI: Pipeline Pulse ───────────────────────────
def render_pulse(inst: str, tok: str, rt: str, schema: Schema, same_day: int):
    st.subheader("Pipeline Pulse")
    amt = schema.amt
    colf = st.columns([1, 1, 2])
    start, end, period = period_picker(
        colf[0], "pulse",
        ["Week to date", "Month to date", "Quarter to date", "Year to date", "Last 90 days",
         "Last month", "Last quarter", "Custom range"], index=1)
    colf[1].caption(f"{start:%m/%d/%Y} → {end:%m/%d/%Y}")

    sel = schema.select

    # --- open (in-flight) draws: not wired, not terminal ---
    open_status = "(" + ",".join(f"'{s}'" for s in OPEN_EXCLUDE_STATUS) + ")"
    open_df = run_soql(inst, tok,
        f"SELECT {sel} FROM Advance__c WHERE RecordTypeId='{rt}' "
        f"AND {WIRE_FIELD}=null AND Status__c NOT IN {open_status}")
    open_df = add_intervals(open_df, same_day)
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
        within3 = (done_df["turn_bd"].dropna() <= 3).mean() * 100
        k[3].metric("Median turn-time", f"{med:.0f} bd", f"{within3:.0f}% ≤3 bd")
    else:
        k[3].metric("Median turn-time", "—", "no package dates in window")

    st.divider()

    # --- open by stage + on hold ---
    hold_cols = schema.roles.get("hold_reason", [])
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
            show = [c for c in ["Name", "Deal__r.Name", "Status__c", *hold_cols,
                                "Advance_Coordinator__r.Name"] if c in holds]
            st.dataframe(holds[show].rename(columns=schema.label), width="stretch", height=240)
        else:
            st.success("Nothing sitting on hold.")

    st.divider()

    # --- aging of open draws ---
    st.markdown("**Aging — oldest open draws (calendar days since requested)**")
    if not open_df.empty and REQ_FIELD in open_df:
        aged = open_df.copy()
        aged["Days open"] = (pd.Timestamp(date.today()) - aged[REQ_FIELD]).dt.days
        aged = aged.sort_values("Days open", ascending=False)
        show = [c for c in ["Name", "Deal__r.Name", "Status__c", "Days open"] if c in aged]
        st.dataframe(aged[show].head(15).rename(columns=schema.label), width="stretch", height=300)
        aged = attach_property_ids(inst, tok, aged)
        cols = [*id_cols(schema), *people_cols(schema), amt, *[m[1] for m in MILESTONES],
                "Days open", *note_cols(schema)]
        cols = list(dict.fromkeys(c for c in cols if c and c in aged))
        st.download_button("Download open pipeline (Excel)",
                           build_workbook({"Open draws": aged[cols].rename(columns=schema.label)},
                                          money_cols=(schema.label(amt),) if amt else ()),
                           f"open_draw_pipeline_{date.today():%Y-%m-%d}.xlsx")
    else:
        st.info("No requested-date data on open draws.")

    st.divider()

    # --- recent wires ---
    st.markdown("**Recently completed (last 15 wires in window)**")
    if not done_df.empty:
        rc = done_df.sort_values(WIRE_FIELD, ascending=False).head(15)
        show = [c for c in ["Name", "Deal__r.Name", WIRE_FIELD, amt, "turn_bd"] if c and c in rc]
        st.dataframe(rc[show].rename(columns=schema.label), width="stretch", height=300)
        st.caption("Full detail and Excel export for any period: **Turn-Time Report** in the sidebar.")
    else:
        st.info("No completed draws in this window.")

    if "turn_bd" in done_df:
        st.caption("Turn-time = business days from full draw package received to wire. Only draws "
                   "with a recorded package date are measured — coverage in Salesforce is partial; "
                   "Land Gorilla (IHD-109768) fills the rest.")


# ───────────────────────────── UI: Turn-Time Report ─────────────────────────
TURN_BUCKETS = [(-1, 0, "0 bd"), (0, 1, "1 bd"), (1, 2, "2 bd"), (2, 3, "3 bd"),
                (3, 5, "4–5 bd"), (5, 10, "6–10 bd"), (10, np.inf, "11+ bd")]


def _pct(s: pd.Series, limit: int) -> float | None:
    s = s.dropna()
    return float((s <= limit).mean()) if len(s) else None


def _med(s: pd.Series) -> float | None:
    s = s.dropna()
    return float(s.median()) if len(s) else None


def build_detail(df: pd.DataFrame, schema: Schema, inst: str) -> pd.DataFrame:
    """One row per draw, in the column order Melanie asked for, with readable headings."""
    df = df.copy()
    stage_cols = [k for k, *_ in STAGES if k in df]
    if stage_cols:
        stage_names = {k: lbl.replace(" (bd)", "") for k, lbl, *_ in STAGES}
        has = df[stage_cols].notna().any(axis=1)
        df["longest_stage"] = None
        df.loc[has, "longest_stage"] = df.loc[has, stage_cols].idxmax(axis=1).map(stage_names)
    if "turn_bd" in df:
        df["within_3"] = np.where(df["turn_bd"].isna(), None,
                                  np.where(df["turn_bd"] <= 3, "Yes", "No"))
    if "Id" in df:
        df["sf_link"] = inst.rstrip("/") + "/" + df["Id"].astype(str)
    cols = [*id_cols(schema), *people_cols(schema), schema.amt,
            *[m[1] for m in MILESTONES],
            "turn_bd", "within_3", "prepkg_bd", "total_bd", *stage_cols, "longest_stage",
            *note_cols(schema), schema.role("payoff_date"), schema.role("servicer_status"), "sf_link"]
    cols = list(dict.fromkeys(c for c in cols if c and c in df))
    out = df[cols].rename(columns=schema.label)
    seen: dict[str, int] = {}
    heads = []
    for h in out.columns:       # two fields can share a label; keep headings unique
        seen[h] = seen.get(h, 0) + 1
        heads.append(h if seen[h] == 1 else f"{h} ({seen[h]})")
    out.columns = heads
    return out


def build_summary(df: pd.DataFrame, schema: Schema, start: date, end: date,
                  period: str, same_day: int) -> pd.DataFrame:
    t = df.get("turn_bd", pd.Series(dtype=float))
    rows = [
        ("Report period", period),
        ("Wire date from", start.strftime("%m/%d/%Y")),
        ("Wire date to", end.strftime("%m/%d/%Y")),
        ("Generated", datetime.now().strftime("%m/%d/%Y %I:%M %p")),
        ("Same-day convention", f"Package in and wired same day = {same_day} business day(s)"),
        ("", ""),
        ("Draws wired", len(df)),
    ]
    if schema.amt and schema.amt in df:
        rows.append(("Total funded", money(pd.to_numeric(df[schema.amt], errors="coerce").sum())))
    measured = int(t.notna().sum())
    rows += [
        ("Draws with a full-package date (measured)", measured),
        ("Measured coverage", f"{measured / len(df):.0%}" if len(df) else "—"),
        ("", ""),
        ("OFFICIAL TURN-TIME: full package received → wire", ""),
        ("  Median (business days)", _fmt_bd(_med(t))),
        ("  Average (business days)", _fmt_bd(float(t.mean()) if measured else None)),
        ("  Funded within 1 business day", _fmt_pct(_pct(t, 1))),
        ("  Funded within 3 business days", _fmt_pct(_pct(t, 3))),
        ("  Funded within 5 business days", _fmt_pct(_pct(t, 5))),
        ("", ""),
        ("BORROWER / THIRD-PARTY TIME: request → full package", ""),
        ("  Median (business days)", _fmt_bd(_med(df.get("prepkg_bd", pd.Series(dtype=float))))),
    ]
    for key, lbl, *_ in STAGES:
        if key in df:
            rows.append((f"    {lbl.replace(' (bd)', '')} — median bd", _fmt_bd(_med(df[key]))))
    rows += [
        ("", ""),
        ("END TO END: request → wire", ""),
        ("  Median (business days)", _fmt_bd(_med(df.get("total_bd", pd.Series(dtype=float))))),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def _fmt_bd(x: float | None) -> str:
    return "—" if x is None or pd.isna(x) else f"{x:.1f}"


def _fmt_pct(x: float | None) -> str:
    return "—" if x is None else f"{x:.0%}"


def group_summary(df: pd.DataFrame, by: pd.Series, name: str, amt: str | None) -> pd.DataFrame:
    g = df.groupby(by, dropna=False)
    out = pd.DataFrame({"Draws wired": g.size()})
    if amt and amt in df:
        out["Funded $"] = g[amt].apply(lambda s: pd.to_numeric(s, errors="coerce").sum())
    if "turn_bd" in df:
        out["Measured"] = g["turn_bd"].count()
        out["Median turn-time (bd)"] = g["turn_bd"].median().round(1)
        out["Avg turn-time (bd)"] = g["turn_bd"].mean().round(1)
        out["% within 3 bd"] = g["turn_bd"].apply(lambda s: _pct(s, 3))
    if "prepkg_bd" in df:
        out["Median request → package (bd)"] = g["prepkg_bd"].median().round(1)
    if "total_bd" in df:
        out["Median request → wire (bd)"] = g["total_bd"].median().round(1)
    out.index.name = name
    return out.reset_index()


DEFINITIONS = pd.DataFrame([
    ("Scope", "Salesforce Advance__c records with record type 'Construction Advance', "
              "wire date inside the report window. No filter on loan status, so draws on loans "
              "paid off after funding are still included ('Loan payoff date' shows which)."),
    ("Land Gorilla loan ID", "From the deal's Property records (ConstructionManagementLoanId__c). "
                             "Small-balance RTL / fix-and-flip loans start with RB0."),
    ("Official turn-time", "Business days from 'Date full draw package received' "
                           "(Date_Submitted_to_Capital_Partner__c) to wire date. Weekends and US "
                           "federal holidays excluded. This is the internal funding time."),
    ("Full draw package", "Package is complete only when everything needed to fund is in hand — "
                          "e.g. inspection report, lien waivers, title cleared. A draw with an "
                          "open lien is not a complete package until the lien is resolved."),
    ("Request → full package", "Business days from draw requested to full package received — time "
                               "waiting on the borrower, inspection, title or other conditions."),
    ("Stage columns", "Plain business-day counts between consecutive milestones, to show where "
                      "pre-package time went. 'Longest stage' is the biggest of those."),
    ("Request → wire", "End-to-end business days, borrower request to funding."),
    ("Blank values", "A blank interval means one of its two dates is not recorded in Salesforce. "
                     "Those draws are excluded from medians and percentages."),
    ("Same-day convention", "Whether a package received and wired the same day counts as 0 or 1 "
                            "business day (sidebar setting; shown on the Summary sheet)."),
], columns=["Term", "Definition"])


def render_report(inst: str, tok: str, rt: str, schema: Schema, same_day: int):
    st.subheader("Turn-Time Report")
    st.caption("Every construction draw wired in the window, with milestone dates and business-day "
               "turn-times. Download as Excel for pivots or management requests.")
    amt = schema.amt
    top = st.columns([1, 1, 2])
    start, end, period = period_picker(top[0], "report", PERIODS, index=0)
    top[1].caption(f"{start:%m/%d/%Y} → {end:%m/%d/%Y}")

    raw = run_soql(inst, tok,
        f"SELECT {schema.select} FROM Advance__c WHERE RecordTypeId='{rt}' "
        f"AND {WIRE_FIELD}>={start:%Y-%m-%d} AND {WIRE_FIELD}<={end:%Y-%m-%d} "
        f"ORDER BY {WIRE_FIELD}")
    df = attach_property_ids(inst, tok, add_intervals(raw, same_day))
    if df.empty:
        st.info("No construction draws were wired in this window.")
        return

    # --- optional filters ---
    mgr_col = schema.role("construction_manager") or "Advance_Coordinator__r.Name"
    with st.expander("Filters"):
        f = st.columns(4)
        prod = schema.role("product_type")
        if prod and prod in df:
            types = sorted(df[prod].dropna().astype(str).unique())
            pick = f[3].multiselect("Product type", types,
                                    help="e.g. keep small-balance RTL / fix-and-flip, drop build-to-rent.")
            if pick:
                df = df[df[prod].astype(str).isin(pick)]
        if "Lender__c" in df:
            lenders = sorted(df["Lender__c"].dropna().astype(str).unique())
            pick = f[0].multiselect("Lender", lenders)
            if pick:
                df = df[df["Lender__c"].astype(str).isin(pick)]
        if mgr_col in df:
            mgrs = sorted(df[mgr_col].dropna().astype(str).unique())
            pick = f[1].multiselect(schema.label(mgr_col), mgrs)
            if pick:
                df = df[df[mgr_col].astype(str).isin(pick)]
        if f[2].checkbox("Only draws with a full-package date", value=False):
            df = df[df.get("turn_bd", pd.Series(np.nan, index=df.index)).notna()]
    if df.empty:
        st.info("No draws match these filters.")
        return

    # --- KPIs ---
    t = df.get("turn_bd", pd.Series(np.nan, index=df.index))
    k = st.columns(5)
    k[0].metric("Draws wired", f"{len(df):,}")
    k[1].metric("Funded", money(pd.to_numeric(df[amt], errors="coerce").sum()) if amt and amt in df else "—")
    k[2].metric("Median turn-time", f"{_med(t):.1f} bd" if t.notna().any() else "—",
                f"{_pct(t, 3):.0%} within 3 bd" if t.notna().any() else None, delta_color="off")
    k[3].metric("Median request → package", _fmt_bd(_med(df.get("prepkg_bd", pd.Series(dtype=float)))) + " bd",
                help="Borrower / inspection / title time before the package is complete.")
    k[4].metric("Measured", f"{int(t.notna().sum())} of {len(df)}",
                help="Draws with both a full-package date and a wire date.")

    # --- charts ---
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Turn-time distribution** (package → wire)")
        if t.notna().any():
            labels = [lbl for *_, lbl in TURN_BUCKETS]
            buckets = pd.cut(t.dropna(), [lo for lo, *_ in TURN_BUCKETS] + [np.inf],
                             labels=labels, right=True)
            dist = buckets.value_counts().reindex(labels, fill_value=0).rename("Draws")
            st.bar_chart(dist)
        else:
            st.info("No full-package dates recorded in this window.")
    with c2:
        st.markdown("**Where the time goes** (median business days)")
        parts = {lbl.replace(" (bd)", ""): _med(df[key]) for key, lbl, *_ in STAGES if key in df}
        parts["Full package → Wire"] = _med(t)
        parts = {k2: v for k2, v in parts.items() if v is not None}
        if parts:
            st.bar_chart(pd.Series(parts, name="Median bd"), horizontal=True)
        else:
            st.info("Not enough milestone dates to break this down.")

    # --- tables ---
    wire_month = df[WIRE_FIELD].dt.to_period("M").astype(str)
    by_month = group_summary(df, wire_month, "Wire month", amt)
    mgr_series = df[mgr_col].fillna("(none)") if mgr_col in df else pd.Series("(none)", index=df.index)
    by_mgr = group_summary(df, mgr_series, schema.label(mgr_col), amt)
    detail = build_detail(df, schema, inst)

    tabs = st.tabs(["Draw detail", "By month", f"By {schema.label(mgr_col).lower()}"])
    tabs[0].dataframe(detail, width="stretch", hide_index=True, height=420,
                      column_config={"Salesforce": st.column_config.LinkColumn(display_text="open")})
    pct = {"% within 3 bd": st.column_config.NumberColumn(format="percent")}
    tabs[1].dataframe(by_month, width="stretch", hide_index=True, column_config=pct)
    tabs[2].dataframe(by_mgr, width="stretch", hide_index=True, column_config=pct)

    summary = build_summary(df, schema, start, end, period, same_day)
    money_cols = ("Funded $",) + ((schema.label(amt),) if amt else ())
    xlsx = build_workbook({"Summary": summary, "Draw Detail": detail, "By Month": by_month,
                           f"By {schema.label(mgr_col)}"[:31]: by_mgr, "Definitions": DEFINITIONS},
                          money_cols=money_cols)
    st.download_button("⬇️ Download Excel report", xlsx,
                       f"draw_turn_time_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.xlsx",
                       type="primary")

    missing = [lbl for key, lbl in [("construction_manager", "construction manager"),
                                    ("hold_reason", "hold / delay reason"),
                                    ("loan_number", "loan #"),
                                    ("product_type", "product type")] if not schema.roles.get(key)]
    if missing:
        st.caption("Not found in Salesforce by label (so not in the export yet): "
                   + ", ".join(missing) + ". Share the field names and they can be added.")


# ───────────────────────────── UI: Draw Lookup ──────────────────────────────
_pretty = {
    "Name": "Advance #", "Deal__r.Name": "Property / Deal", "Status__c": "Status",
    "Lender__c": "Lender", "Advance_Coordinator__r.Name": "Coordinator",
    "Advance_Analyst__r.Name": "Analyst", "Underwriter__r.Name": "Underwriter",
    "Advance_Requestor__r.Name": "Requestor", "Exception__c": "Exception",
    **{f: lbl for lbl, f in MILESTONES},
    PKG_FIELD: "Full package received", WIRE_FIELD: "Wire date",
    "turn_bd": "Turn-time (bd)", "prepkg_bd": "Request → package (bd)",
    "total_bd": "Request → wire (bd)", "within_3": "Within 3 bd?",
    "longest_stage": "Longest pre-package stage", "sf_link": "Salesforce",
    **{k: lbl for k, lbl, *_ in STAGES},
}


def render_lookup(inst: str, tok: str, rt: str, schema: Schema, same_day: int):
    st.subheader("Draw Lookup")
    c = st.columns([3, 1])
    text = c[0].text_input("Search by property / deal name or advance #", placeholder="e.g. 745 South 9th Street")
    modes = {"Property / Deal": "Deal__r.Name", "Advance #": "Name"}
    if schema.role("loan_number"):
        modes["Loan #"] = schema.role("loan_number")
    mode = c[1].radio("Match on", list(modes), label_visibility="collapsed")
    if not text:
        st.info("Type a property/deal name or an advance number to see its full draw cycle.")
        return

    esc = soql_escape(text)
    field = modes[mode]
    df = run_soql(inst, tok,
        f"SELECT {schema.select} FROM Advance__c WHERE RecordTypeId='{rt}' "
        f"AND {field} LIKE '%{esc}%' ORDER BY {REQ_FIELD} DESC NULLS LAST")
    df = attach_property_ids(inst, tok, add_intervals(df, same_day))
    if df.empty:
        st.warning("No matching construction advances.")
        return

    st.caption(f"{len(df)} matching advance(s).")
    # group by deal so multiple draws on one property read as a cycle
    deal_col = "Deal__r.Name" if "Deal__r.Name" in df else "Name"
    for deal, g in df.groupby(deal_col, dropna=False):
        with st.expander(f"{deal}  ·  {len(g)} draw(s)", expanded=(len(df) <= 5)):
            for _, row in g.iterrows():
                _render_one_draw(row, schema)


def _render_one_draw(row: pd.Series, schema: Schema):
    amt = schema.amt
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
    st.dataframe(tdf, hide_index=True, width="stretch",
                 column_config={"": st.column_config.TextColumn(width="small")})

    meta = []
    for f in [schema.role("loan_number"), *[f"Property.{k}" for k in PROPERTY_WISH],
              schema.role("product_type"), "Lender__c", schema.role("construction_manager"),
              schema.role("loan_manager"), "Advance_Coordinator__r.Name", "Advance_Analyst__r.Name",
              "Underwriter__r.Name", "Advance_Requestor__r.Name",
              *schema.roles.get("hold_reason", [])]:
        if f and f in row and pd.notna(row[f]):
            meta.append(f"**{schema.label(f)}:** {row[f]}")
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
            if st.button("Log out", width="stretch"):
                clear_salesforce_session(); st.rerun()
        st.divider()
        same_day = 1 if st.radio(
            "Same-day convention",
            ["0 business days", "1 business day"],
            help="Package in and wired the same day counts as this. Confirm with Melanie — it moves the median."
        ).startswith("1") else 0

    # login gate
    if setup_error:
        st.error(setup_error); st.stop()
    if sf is None:
        st.subheader("Log in to Salesforce")
        st.info("Log in to load the construction draw pipeline.")
        st.link_button("Log in to Salesforce", build_salesforce_login_url(cfg))
        st.caption(f"Callback URL: {cfg['redirect_uri']}")
        st.stop()

    inst = st.session_state["salesforce_auth"]["instance_url"]
    tok = st.session_state["salesforce_auth"]["access_token"]

    page = st.sidebar.radio("View", ["Pipeline Pulse", "Turn-Time Report", "Draw Lookup"])
    try:
        rt = construction_rt_id(inst, tok)
        if not rt:
            st.error("Could not find the 'Construction Advance' record type on Advance__c."); st.stop()
        schema = load_schema(inst, tok)
        if page == "Pipeline Pulse":
            render_pulse(inst, tok, rt, schema, same_day)
        elif page == "Turn-Time Report":
            render_report(inst, tok, rt, schema, same_day)
        else:
            render_lookup(inst, tok, rt, schema, same_day)
    except Exception as exc:
        msg = str(exc)
        if "INVALID_SESSION_ID" in msg or "Session expired" in msg:
            clear_salesforce_session()
            st.warning("Your Salesforce session expired. Log in again.")
            st.stop()
        raise


if __name__ == "__main__":
    main()
