"""
Construction Draw Dashboard  (Streamlit)
========================================
Same Salesforce OAuth login as the AM slide app, then three views:

  • Turn-Time Report  — (default view) on-demand, any period (last month / quarter /
                        year, to-date, all time or a custom range): every construction
                        draw wired in the window, official turn-time (full draw package
                        received -> wire, business days), InspRpt->Wire cross-check,
                        pre-package (borrower/inspection/title) and total elapsed time.
                        Excel workbook in the same shape as the notebook export: Read Me,
                        Turn-Time Summary (headline, By Year, By Quarter, % <=3 bd chart),
                        Draw Detail (pivot source with Wire Year/Quarter/Month), By Month,
                        By Coordinator.
  • Pipeline Pulse    — what's happening now: open draws by stage, on-hold draws,
                        completed this period ($ and count), recent wires, aging.
                        Open pipeline downloads to Excel for the Thursday call.
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
           "Year to date", "Last 90 days", "All time", "Custom range"]

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


def bdays(a: pd.Series, b: pd.Series, same_day_as: int, drop_negative: bool = True) -> pd.Series:
    a, b = pd.to_datetime(a, errors="coerce"), pd.to_datetime(b, errors="coerce")
    m = a.notna() & b.notna()
    out = pd.Series(np.nan, index=a.index)
    if m.any():
        out[m] = np.busday_count(a[m].values.astype("datetime64[D]"),
                                 b[m].values.astype("datetime64[D]"), holidays=_HOLS) + same_day_as
    if drop_negative:
        out[out < 0] = np.nan
    return out


def add_intervals(df: pd.DataFrame, same_day_as: int) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    for f in [m[1] for m in MILESTONES] + [DEAL_WISH["payoff_date"][0][0]]:
        if f in df:
            df[f] = to_date(df[f])
    # official + secondary intervals; a negative one means reversed dates -> blank it, flag the row
    df["reversed_dates"] = False
    for key, a, b in [("turn_bd", PKG_FIELD, WIRE_FIELD),       # OFFICIAL: complete package -> wire
                      ("insp_wire_bd", RPT_FIELD, WIRE_FIELD),  # broader-coverage cross-check
                      ("prepkg_bd", REQ_FIELD, PKG_FIELD),      # borrower / inspection / title side
                      ("total_bd", REQ_FIELD, WIRE_FIELD)]:     # total elapsed
        if a in df and b in df:
            v = bdays(df[a], df[b], same_day_as, drop_negative=False)
            df["reversed_dates"] |= (v < 0).fillna(False)
            df[key] = v.where(v >= 0)
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
    if choice == "All time":
        return date(2018, 1, 1), today
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
INTERVALS = ["turn_bd", "insp_wire_bd", "prepkg_bd", "total_bd"]


def _pct(s: pd.Series, limit: int) -> float | None:
    s = s.dropna()
    return float((s <= limit).mean()) if len(s) else None


def _med(s: pd.Series) -> float | None:
    s = s.dropna()
    return float(s.median()) if len(s) else None


def _q90(s: pd.Series) -> float | None:
    s = s.dropna()
    return float(s.quantile(0.9)) if len(s) else None


def _fmt_bd(x: float | None) -> str:
    return "—" if x is None or pd.isna(x) else f"{x:.0f}"


def _fmt_pct(x: float | None) -> str:
    return "—" if x is None else f"{x:.0%}"


def _col(df: pd.DataFrame, name: str) -> pd.Series:
    return df[name] if name in df else pd.Series(np.nan, index=df.index)


def build_detail(df: pd.DataFrame, schema: Schema, inst: str) -> pd.DataFrame:
    """One row per draw — the pivot source. Notebook column order, plus deal/property IDs."""
    df = df.copy()
    if WIRE_FIELD in df:
        w = df[WIRE_FIELD]
        df["wire_year"] = w.dt.year.astype("Int64")
        df["wire_quarter"] = w.dt.to_period("Q").astype(str).replace("NaT", None)
        df["wire_month"] = w.dt.to_period("M").astype(str).replace("NaT", None)
    if "Id" in df:
        df["sf_link"] = inst.rstrip("/") + "/" + df["Id"].astype(str)
    cols = ["Name", "Deal__r.Name", schema.role("loan_number"), *[f"Property.{k}" for k in PROPERTY_WISH],
            schema.role("product_type"), "Lender__c", "Status__c", "IC_Approval_Status__c",
            schema.role("construction_manager"), "Advance_Coordinator__r.Name", "Advance_Analyst__r.Name",
            "Underwriter__r.Name", schema.role("loan_manager"), "Exception__c", "Cancellation_Reason__c",
            "Inspection_Method__c", schema.amt,
            *[m[1] for m in MILESTONES], *INTERVALS,
            *schema.roles.get("hold_reason", []), *schema.roles.get("notes", []),
            schema.role("deal_comments"), schema.role("payoff_date"), schema.role("servicer_status"),
            "wire_year", "wire_quarter", "wire_month", "sf_link"]
    cols = list(dict.fromkeys(c for c in cols if c and c in df))
    out = df.sort_values(WIRE_FIELD, ascending=False)[cols].rename(columns=schema.label)
    seen: dict[str, int] = {}
    heads = []
    for h in out.columns:       # two fields can share a label; keep headings unique
        seen[h] = seen.get(h, 0) + 1
        heads.append(h if seen[h] == 1 else f"{h} ({seen[h]})")
    out.columns = heads
    return out


def rollup(df: pd.DataFrame, by: pd.Series, name: str, amt: str | None) -> pd.DataFrame:
    """Per-group turn-time stats (official metric over draws that have a package date)."""
    g = df.groupby(by, dropna=False)
    t = _col(df, "turn_bd")
    out = pd.DataFrame({"Draws wired": g.size(), "Measured": t.groupby(by, dropna=False).count()})
    if amt and amt in df:
        out["Funded $"] = pd.to_numeric(df[amt], errors="coerce").groupby(by, dropna=False).sum()
    tg = t.groupby(by, dropna=False)
    out["Median bd"] = tg.median()
    out["Mean bd"] = tg.mean().round(1)
    out["% ≤3 bd"] = tg.apply(lambda s: _pct(s, 3))
    out["Median Pre-Package bd"] = _col(df, "prepkg_bd").groupby(by, dropna=False).median()
    out.index.name = name
    return out.reset_index()


def headline_rows(df: pd.DataFrame, amt: str | None) -> list[tuple[str, str]]:
    t, n = _col(df, "turn_bd"), len(df)
    no = int(t.notna().sum())
    ins = _col(df, "insp_wire_bd"); pre = _col(df, "prepkg_bd")
    rows = [("Construction advances (wired)", f"{n:,}")]
    if amt and amt in df:
        rows.append(("Total funded", money(pd.to_numeric(df[amt], errors="coerce").sum())))
    rows += [
        ("…with a Full Draw Package Received date", f"{no:,}  ({no / n:.0%} coverage)" if n else "0"),
        ("Median turn-time (business days)", _fmt_bd(_med(t))),
        ("Mean turn-time (business days)", "—" if not no else f"{t.mean():.1f}"),
        ("% funded within 1 business day", _fmt_pct(_pct(t, 1))),
        ("% funded within 3 business days", _fmt_pct(_pct(t, 3))),
        ("% funded within 5 business days", _fmt_pct(_pct(t, 5))),
        ("Median PRE-package (borrower/inspection/title)",
         f"{_fmt_bd(_med(pre))} bd  (90th pct {_fmt_bd(_q90(pre))})"),
        ("Median total elapsed (request → wire)", f"{_fmt_bd(_med(_col(df, 'total_bd')))} bd"),
        ("Secondary view — Inspection Report→Wire",
         f"median {_fmt_bd(_med(ins))} bd, {_fmt_pct(_pct(ins, 3))} ≤3  (n={int(ins.notna().sum()):,})"),
    ]
    return rows


def readme_rows(*, period: str, start: date, end: date, same_day: int, coverage: float,
                reversed_rows: int, filters: list[str]) -> list[tuple[str, str]]:
    return [
        ("Generated", datetime.now().strftime("%m/%d/%Y %H:%M")),
        ("Report period", f"{period}  ·  wire date {start:%m/%d/%Y} – {end:%m/%d/%Y}"),
        ("Filters", "; ".join(filters) if filters else "None"),
        ("Source", "Salesforce Advance__c · Record Type: Construction Advance · wired advances only. "
                   "No filter on loan status, so draws on loans paid off after funding are included."),
        ("Official metric", "Business days from FULL DRAW PACKAGE RECEIVED to WIRE DATE"),
        ("  field used", "Date_Submitted_to_Capital_Partner__c  (label: 'Date Full Draw Package Received')"),
        ("Business days", "Excludes weekends and US federal bank holidays"),
        ("Same-day convention", f"package-in / wire same-day = {same_day} business day(s)"),
        ("Coverage caveat", f"Only {coverage:.0%} of these draws carry a package-received date; the metric "
                            "covers that subset. Fuller coverage needs the Land Gorilla feed / better SF data "
                            "entry (ref ticket IHD-109768)."),
        ("Complete package", "A package is complete only when everything needed to fund is in hand — a draw "
                             "with an open lien or title issue is not complete until it is resolved."),
        ("Delay story", "For completed draws Status shows 'Completed' only (hold history not retained). Use "
                        "the Pre-Package interval (Req→Package) as the borrower/inspection/title delay measure."),
        ("Secondary view", "Inspection Report Received→Wire is included for broader coverage as a cross-check."),
        ("Reversed-date rows", f"{reversed_rows} rows had an end date before the start date; those intervals "
                               "were blanked and excluded from metrics."),
        ("Land Gorilla loan ID", "From the deal's Property records (ConstructionManagementLoanId__c). "
                                 "Small-balance RTL / fix-and-flip loans start with RB0."),
        ("How to filter by period", "Use 'Draw Detail' → filter on Wire Year / Wire Quarter / Wire Month, "
                                    "or build a PivotTable off that sheet."),
    ]


def render_report(inst: str, tok: str, rt: str, schema: Schema, same_day: int):
    st.subheader("Turn-Time Report")
    st.caption("Business days from full draw package received to wire, for every construction draw "
               "wired in the window. Download the Excel workbook for pivots or management requests.")
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
    applied: list[str] = []
    with st.expander("Filters"):
        f = st.columns(4)
        for slot, col, label, help_ in [
            (f[0], "Lender__c", "Lender", None),
            (f[1], schema.role("product_type"), "Product type",
             "e.g. keep small-balance RTL / fix-and-flip, drop build-to-rent."),
            (f[2], "Inspection_Method__c", "Inspection method", "Land Gorilla, Trinity or TrustPoint."),
            (f[3], mgr_col, schema.label(mgr_col), None),
        ]:
            if col and col in df:
                opts = sorted(df[col].dropna().astype(str).unique())
                pick = slot.multiselect(label, opts, help=help_, key=f"report_f_{col}")
                if pick:
                    df = df[df[col].astype(str).isin(pick)]
                    applied.append(f"{label}: {', '.join(pick)}")
        if st.checkbox("Only draws with a full-package date", value=False):
            df = df[_col(df, "turn_bd").notna()]
            applied.append("Only draws with a full-package date")
    if df.empty:
        st.info("No draws match these filters.")
        return

    # --- KPIs ---
    t = _col(df, "turn_bd")
    no = int(t.notna().sum())
    k = st.columns(5)
    k[0].metric("Wired draws", f"{len(df):,}",
                money(pd.to_numeric(df[amt], errors="coerce").sum()) if amt and amt in df else None,
                delta_color="off", delta_arrow="off")
    k[1].metric("With package date", f"{no:,}", f"{no / len(df):.0%} coverage", delta_color="off")
    k[2].metric("Median turn-time", f"{_fmt_bd(_med(t))} bd", help="Full draw package received → wire.")
    k[3].metric("Funded within 3 bd", _fmt_pct(_pct(t, 3)))
    k[4].metric("Median pre-package", f"{_fmt_bd(_med(_col(df, 'prepkg_bd')))} bd",
                help="Advance requested → full package: borrower / inspection / title time.")
    reversed_rows = int(_col(df, "reversed_dates").fillna(False).astype(bool).sum())
    if reversed_rows:
        st.caption(f"⚠ {reversed_rows} draw(s) have an end date before the start date — those intervals "
                   "are blanked and excluded from the metrics.")

    # --- charts ---
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Turn-time distribution** (package → wire)")
        if t.notna().any():
            labels = [lbl for *_, lbl in TURN_BUCKETS]
            buckets = pd.cut(t.dropna(), [lo for lo, *_ in TURN_BUCKETS] + [np.inf],
                             labels=labels, right=True)
            st.bar_chart(buckets.value_counts().reindex(labels, fill_value=0).rename("Draws"))
        else:
            st.info("No full-package dates recorded in this window.")
    with c2:
        st.markdown("**Before vs. after the package is complete** (median business days)")
        parts = {lbl.replace(" (bd)", ""): _med(df[key]) for key, lbl, *_ in STAGES if key in df}
        parts["Full package → Wire"] = _med(t)
        parts = {k2: v for k2, v in parts.items() if v is not None}
        if parts:
            st.bar_chart(pd.Series(parts, name="Median bd"), horizontal=True)
        else:
            st.info("Not enough milestone dates to break this down.")

    # --- rollups + detail ---
    w = df[WIRE_FIELD]
    rollups = {
        "By Year": rollup(df, w.dt.year.astype("Int64").astype(str), "Wire Year", amt),
        "By Quarter": rollup(df, w.dt.to_period("Q").astype(str), "Wire Quarter", amt),
        "By Month": rollup(df, w.dt.to_period("M").astype(str), "Wire Month", amt),
    }
    mgr_label = schema.label(mgr_col)
    if mgr_col in df:
        rollups[f"By {mgr_label}"[:31]] = rollup(df, df[mgr_col].fillna("(none)"), mgr_label, amt)
    detail = build_detail(df, schema, inst)

    names = list(rollups)
    tabs = st.tabs(names + ["Draw detail"])
    pct = {"% ≤3 bd": st.column_config.NumberColumn(format="percent"),
           "Funded $": st.column_config.NumberColumn(format="dollar")}
    for tab, name in zip(tabs, names):
        tab.dataframe(rollups[name], width="stretch", hide_index=True, column_config=pct)
    tabs[-1].dataframe(detail, width="stretch", hide_index=True, height=420,
                       column_config={"Salesforce": st.column_config.LinkColumn(display_text="open")})

    # --- workbook ---
    money_cols = ("Funded $",) + ((schema.label(amt),) if amt else ())
    wrap_cols = tuple(schema.label(c) for c in [*schema.roles.get("notes", []), schema.role("deal_comments")] if c)
    xlsx = build_turn_time_workbook(
        readme=readme_rows(period=period, start=start, end=end, same_day=same_day,
                           coverage=no / len(df), reversed_rows=reversed_rows, filters=applied),
        headline=headline_rows(df, amt),
        subtitle=f"{period} ({start:%m/%d/%Y} – {end:%m/%d/%Y}) · median {_fmt_bd(_med(t))} business "
                 f"day(s) · {_fmt_pct(_pct(t, 3))} funded within 3 · n={no:,}",
        rollups=rollups, detail=detail, money_cols=money_cols, wrap_cols=wrap_cols)
    st.download_button("⬇️ Download Excel report", xlsx,
                       f"Construction_Draw_TurnTime_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.xlsx",
                       type="primary")

    missing = [lbl for key, lbl in [("construction_manager", "construction manager"),
                                    ("loan_number", "loan #"),
                                    ("product_type", "product type")] if not schema.roles.get(key)]
    if missing:
        st.caption("Not found in Salesforce (so not in the export yet): " + ", ".join(missing) + ".")


# ───────────────────────────── UI: Draw Lookup ──────────────────────────────
# Column headings — same names as the notebook export.
_pretty = {
    "Name": "Advance #", "Deal__r.Name": "Property", "Status__c": "Status", "Lender__c": "Lender",
    "IC_Approval_Status__c": "IC Approval", "Exception__c": "Exception",
    "Cancellation_Reason__c": "Cancellation Reason", "Inspection_Method__c": "Inspection Method",
    "Advance_Coordinator__r.Name": "Advance Coordinator", "Advance_Analyst__r.Name": "Advance Analyst",
    "Underwriter__r.Name": "Underwriter", "Advance_Requestor__r.Name": "Advance Requestor",
    "Net_Funded_Amount__c": "Net Funded ($)",
    REQ_FIELD: "Advance Requested", "Date_Inspection_Ordered__c": "Inspection Ordered",
    INSP_FIELD: "Inspection Date", RPT_FIELD: "Inspection Report Received",
    "Date_Submitted_For_Approval__c": "Submitted For Review",
    "Date_Internal_Review_Complete__c": "Internal Review Complete",
    PKG_FIELD: "Full Draw Package Received", "Manager_Approval_Date__c": "Manager Approval",
    WIRE_FIELD: "Wire Date",
    "turn_bd": "Turn-Time bd (Package→Wire)", "insp_wire_bd": "Turn-Time bd (InspRpt→Wire)",
    "prepkg_bd": "Pre-Package bd (Req→Package)", "total_bd": "Total Elapsed bd (Req→Wire)",
    "wire_year": "Wire Year", "wire_quarter": "Wire Quarter", "wire_month": "Wire Month",
    "sf_link": "Salesforce",
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

    page = st.sidebar.radio("View", ["Turn-Time Report", "Pipeline Pulse", "Draw Lookup"])
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
