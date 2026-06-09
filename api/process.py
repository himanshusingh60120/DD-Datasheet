import os
import io
import json
import re
import calendar
import datetime as dt
from urllib.parse import urlencode, urlparse, parse_qs, unquote
from collections import defaultdict
from http.server import BaseHTTPRequestHandler

import requests
import pandas as pd
from google.oauth2.credentials import Credentials
import gspread
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import (
    RunReportRequest, DateRange, Dimension, Metric,
)

# ============================================================
# Configuration (unchanged from your script)
# ============================================================
GA4_PROPERTY = {
    "MarTech360": "382134659",
    "ITDigest":   "383424934",
    "AITech365":  "398862446",
    "ReadMag":    "382256344",
}
SHEET_ID = {
    "MarTech360": "1X4FssrUdcjPSdeJYpRVbbEPnBTC7kCpD49-QSbZ_GEM",
    "ReadMag":    "16egTzcMB6L6jGeKc6eWYkmDXrpyTiQflpU9gTscN7EE",
    "ITDigest":   "1TJ8fwplZNxaPJ5f6z1E31F2VibbGoQsr2t2VMhQUg1A",
    "AITech365":  "1lLbnvKYZDsBXC_F2xKOAuAWsCz-ETwOWZfOb1B0rs8I",
}
DOMAIN = {
    "MarTech360": "https://martech360.com",
    "ITDigest":   "https://itdigest.com",
    "AITech365":  "https://aitech365.com",
    "ReadMag":    "https://readmagazine.com",
}
GA_METRIC    = "activeUsers"
GA_DIMENSION = "pagePath"
TAB_MASTER       = "MasterSheet"
TAB_ACTIVE_USERS = "Active Users"
TAB_FOLLOWERS    = "Followers"
XLS_INHOUSE = "DD Inhouse"
XLS_NEWS    = "DD_News"
XLS_DAILY   = "Daily Report"

SCOPES = [
    "https://www.googleapis.com/auth/analytics.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ============================================================
# Week / URL / segment helpers (unchanged logic)
# ============================================================
def build_week_map(year, month):
    ndays = calendar.monthrange(year, month)[1]
    days  = [dt.date(year, month, d) for d in range(1, ndays + 1)]
    raw, widx = {}, 1
    for i, d in enumerate(days):
        if d.weekday() == 0 and i != 0:
            widx += 1
        raw[d] = widx
    for d in raw:
        if raw[d] > 4:
            raw[d] = 4
    return {d: f"Week {w}" for d, w in raw.items()}

def normalise_url(value, domain):
    if value is None:
        return None
    low = unquote(str(value)).strip().lower()
    if low in ("", "(not set)"):
        return domain + "/na"
    if ".html" in low:
        return None
    if "://" in low:
        path = urlparse(low).path or "/"
    else:
        path = low.split("?")[0].split("#")[0]
        if not path.startswith("/"):
            path = "/" + path
    if len(path) > 1:
        path = path.rstrip("/")
    return domain + path

SEGMENT_CANON = {
    "quickbyte": "QuickByte", "news based article": "News Based Article",
    "news recreation": "News Recreation", "guest post": "Guest Article",
    "guest article": "Guest Article", "press release": "Press Release",
    "pr": "Press Release", "news": "News", "article": "Article",
    "roundup": "Roundup", "interview": "Interview", "revamp article": "Revamp Article",
}
def canon_segment(seg, normalise=True):
    if seg is None:
        return None
    if not normalise:
        return str(seg).strip()
    return SEGMENT_CANON.get(str(seg).strip().lower(), str(seg).strip())

PUB_ALIASES = {
    "martech360": "MarTech360", "itdigest": "ITDigest",
    "aitech365": "AITech365", "read magazine": "ReadMag",
    "readmag": "ReadMag", "readmagazine": "ReadMag",
}
def canon_pub(name):
    if name is None:
        return None
    return PUB_ALIASES.get(str(name).strip().lower())

MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}
MONTHS.update({m.lower(): i for i, m in enumerate(calendar.month_abbr) if m})

def parse_month_year(fname):
    base = fname.rsplit("/", 1)[-1]
    name = re.sub(r"\.(xlsx|xls|csv)$", "", base, flags=re.I)
    mon = None
    for tok in re.split(r"[ _\-.]+", name):
        if tok.lower() in MONTHS:
            mon = MONTHS[tok.lower()]; break
    yr = None
    yrs = re.findall(r"(20\d{2})", name)
    if yrs:
        yr = int(yrs[0])
    else:
        two = re.findall(r"(?<!\d)(\d{2})(?!\d)", name)
        if two:
            yr = 2000 + int(two[-1])
    return mon, yr

def _to_date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date):     return v
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try: return dt.datetime.strptime(str(v).strip(), fmt).date()
        except Exception: pass
    return None

def safe_int(val):
    if pd.isna(val) or val is None or str(val).strip() == '':
        return 0
    try:
        return int(float(str(val).replace(',', '').strip()))
    except (ValueError, TypeError):
        return 0

# ============================================================
# Core pipeline — accepts an in-memory xlsx + creds, runs everything
# ============================================================
def run_pipeline(creds, xls_bytes, xls_filename, year=None, month=None,
                 normalise_seg=True, write_mode="overwrite_month", dry_run=False):
    log = []
    def emit(msg):
        log.append(str(msg))

    ga_client = BetaAnalyticsDataClient(credentials=creds)
    gc = gspread.authorize(creds)

    det_month, det_year = parse_month_year(xls_filename)
    MONTH = month if month else det_month
    YEAR  = year if year else det_year
    if not MONTH or not YEAR:
        raise ValueError("Could not detect month/year from filename — set them manually.")
    MONTH_NAME = calendar.month_name[MONTH]
    WEEK_MAP = build_week_map(YEAR, MONTH)
    emit(f"Detected period: {MONTH_NAME} {YEAR}")

    xls = io.BytesIO(xls_bytes)
    rows_by_pub = defaultdict(list)

    # ---- DD Inhouse ----
    inh = pd.read_excel(xls, sheet_name=XLS_INHOUSE, dtype=object)
    inh.columns = [str(c).strip() for c in inh.columns]
    for _, r in inh.iterrows():
        pub = canon_pub(r.get("Media Publication"))
        if pub is None: continue
        pdate = _to_date(r.get("Published Date") or r.get("Date"))
        if pdate is None or pdate.year != YEAR or pdate.month != MONTH: continue
        link = r.get("Published Link")
        if not link or str(link).strip().lower() in ("", "none"): continue
        wk = WEEK_MAP.get(pdate)
        if wk is None: continue
        rows_by_pub[pub].append({
            "date": pdate, "month": MONTH_NAME, "week": wk,
            "segment": canon_segment(r.get("Segment"), normalise_seg),
            "link": str(link).strip(),
            "url_key": normalise_url(link, DOMAIN[pub]),
        })

    # ---- DD_News ----
    xls.seek(0)
    news = pd.read_excel(xls, sheet_name=XLS_NEWS, dtype=object)
    news.columns = [str(c).strip() for c in news.columns]
    for _, r in news.iterrows():
        pub = canon_pub(r.get("Properties of News")) or canon_pub(r.get("Publisher Name"))
        if pub is None: continue
        pdate = _to_date(r.get("Published Date") or r.get("Date"))
        if pdate is None or pdate.year != YEAR or pdate.month != MONTH: continue
        link = r.get("Published Link")
        if not link or str(link).strip().lower() in ("", "none"): continue
        wk = WEEK_MAP.get(pdate)
        if wk is None: continue
        rows_by_pub[pub].append({
            "date": pdate, "month": MONTH_NAME, "week": wk,
            "segment": canon_segment(r.get("News/PR"), normalise_seg) or "News",
            "link": str(link).strip(),
            "url_key": normalise_url(link, DOMAIN[pub]),
        })

    # ---- Daily Report (Followers) ----
    daily_stats = defaultdict(lambda: defaultdict(dict))
    weekly_follower_gains = defaultdict(
        lambda: defaultdict(lambda: {"LinkedIn": 0, "Twitter": 0, "Facebook": 0}))
    try:
        xls.seek(0)
        daily = pd.read_excel(xls, sheet_name=XLS_DAILY, dtype=object)
        daily.columns = [str(c).strip() for c in daily.columns]
        li_col = next((c for c in daily.columns if "linkedin" in c.lower() and "follower" in c.lower()), None)
        tw_col = next((c for c in daily.columns if ("twitter" in c.lower() or "x" in c.lower()) and "follower" in c.lower()), None)
        fb_col = next((c for c in daily.columns if ("facebook" in c.lower() or "fb" in c.lower()) and "follower" in c.lower()), None)
        for _, r in daily.iterrows():
            pub = canon_pub(r.get("Website") or r.get("Media Publication") or r.get("Publication"))
            if pub is None: continue
            pdate = _to_date(r.get("Date"))
            if pdate is None: continue
            if li_col: daily_stats[pub][pdate]["LinkedIn"] = safe_int(r[li_col])
            if tw_col: daily_stats[pub][pdate]["Twitter"] = safe_int(r[tw_col])
            if fb_col: daily_stats[pub][pdate]["Facebook"] = safe_int(r[fb_col])
        for pub in GA4_PROPERTY:
            pub_dates = sorted(daily_stats[pub].keys())
            if not pub_dates: continue

            # True [start, end] boundary date for each week, from the week map.
            wk_bounds = {}  # "Week N" -> [first_day, last_day]
            for d, wk in WEEK_MAP.items():
                if wk not in wk_bounds:
                    wk_bounds[wk] = [d, d]
                else:
                    if d < wk_bounds[wk][0]: wk_bounds[wk][0] = d
                    if d > wk_bounds[wk][1]: wk_bounds[wk][1] = d

            def nearest(target, candidates):
                return min(candidates, key=lambda d: abs((d - target).days))

            # CARRY-OVER model:
            #   each week's delta = (this week's last reading) - (previous week's last reading)
            #   Week 1's baseline  = last logged reading BEFORE this month begins
            #                        (i.e. how the previous month ended).
            # This chains continuously so Week 1 is never a within-week 0.
            month_start = min(wk_bounds[w][0] for w in wk_bounds)
            prior_dates = [d for d in pub_dates if d < month_start]
            prev_end_date = prior_dates[-1] if prior_dates else None  # baseline for Week 1

            for wk_num in [1, 2, 3, 4]:
                wk_str = f"Week {wk_num}"
                if wk_str not in wk_bounds: continue
                bound_end = wk_bounds[wk_str][1]

                in_week = [d for d in pub_dates
                           if d.year == YEAR and d.month == MONTH and WEEK_MAP.get(d) == wk_str]
                if not in_week:
                    # no readings this week: delta is 0, and the baseline rolls forward unchanged
                    continue

                # this week's "end" = reading nearest the week's last day
                end_date = nearest(bound_end, in_week)

                for plat in ["LinkedIn", "Twitter", "Facebook"]:
                    end_val = daily_stats[pub][end_date].get(plat, 0)
                    if prev_end_date is not None:
                        base_val = daily_stats[pub][prev_end_date].get(plat, 0)
                    else:
                        # no prior reading at all (first month of data):
                        # fall back to this week's first reading so delta is within-week
                        first_in_week = nearest(wk_bounds[wk_str][0], in_week)
                        base_val = daily_stats[pub][first_in_week].get(plat, 0)
                    weekly_follower_gains[pub][wk_str][plat] = end_val - base_val

                # roll the baseline forward to this week's end for the next week
                prev_end_date = end_date
        emit("Follower deltas computed.")
    except Exception as e:
        emit(f"Followers skipped: {e}")

    # ---- GA4 fetch ----
    def ga_active_users(property_id, start, end):
        req = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[DateRange(start_date=start, end_date=end)],
            dimensions=[Dimension(name=GA_DIMENSION), Dimension(name="date")],
            metrics=[Metric(name=GA_METRIC)],
            limit=250000,
        )
        resp = ga_client.run_report(req)
        return [(row.dimension_values[0].value,
                 row.dimension_values[1].value,
                 float(row.metric_values[0].value or 0)) for row in resp.rows]

    def gdate_to_date(s): return dt.date(int(s[:4]), int(s[4:6]), int(s[6:8]))

    ndays = calendar.monthrange(YEAR, MONTH)[1]
    START = f"{YEAR}-{MONTH:02d}-01"
    END   = f"{YEAR}-{MONTH:02d}-{ndays:02d}"

    ga_week = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    site_wide_weekly_traffic = defaultdict(lambda: defaultdict(float))

    for pub, pid in GA4_PROPERTY.items():
        dom = DOMAIN[pub]
        try:
            raw = ga_active_users(pid, START, END)
            for path, gdate, val in raw:
                d = gdate_to_date(gdate)
                wk = WEEK_MAP.get(d)
                if wk is None: continue
                site_wide_weekly_traffic[pub][wk] += val
                key = dom + "/na" if str(path).strip().lower() == "(not set)" else normalise_url(path, dom)
                if key is None: continue
                ga_week[pub][key][wk] += val
            emit(f"GA4 {pub}: {len(raw)} pages | "
                 f"W1={int(round(site_wide_weekly_traffic[pub]['Week 1']))} "
                 f"W2={int(round(site_wide_weekly_traffic[pub]['Week 2']))} "
                 f"W3={int(round(site_wide_weekly_traffic[pub]['Week 3']))} "
                 f"W4={int(round(site_wide_weekly_traffic[pub]['Week 4']))}")
        except Exception as e:
            emit(f"GA4 error {pub}: {e}")

    # ---- Join individual links ----
    for pub, rows in rows_by_pub.items():
        gw = ga_week.get(pub, {})
        for row in rows:
            au = gw.get(row["url_key"], {}).get(row["week"])
            row["active_users"] = int(round(au)) if (au and au > 0) else 0

    # ---- Writers ----
    def master_df(pub):
        rows = sorted(rows_by_pub.get(pub, []), key=lambda r: (r["date"], r["week"]))
        return pd.DataFrame([{
            "Date": r["date"].strftime("%d-%m-%Y"), "Month": r["month"], "Week": r["week"],
            "Segment": r["segment"], "Published Link": r["link"], "Active Users": r["active_users"],
        } for r in rows])

    def write_master(pub):
        df = master_df(pub)
        if df.empty: return
        sh = gc.open_by_key(SHEET_ID[pub])
        ws = sh.worksheet(TAB_MASTER)
        values = df.values.tolist()
        existing = ws.get_all_values()
        header, body = existing[0], existing[1:]
        present = any(len(r) >= 2 and r[1].strip().lower() == MONTH_NAME.lower() for r in body)
        if write_mode == "overwrite_month":
            keep = [r for r in body if len(r) < 2 or r[1].strip().lower() != MONTH_NAME.lower()]
            if not dry_run:
                ws.clear()
                ws.update([header] + keep + values, value_input_option="USER_ENTERED")
            emit(f"{pub}: MasterSheet overwrote {MONTH_NAME}.")
        else:
            if present:
                emit(f"{pub}: MasterSheet skipped — {MONTH_NAME} exists.")
                return
            if not dry_run: ws.append_rows(values, value_input_option="USER_ENTERED")
            emit(f"{pub}: MasterSheet appended.")

    def write_active_users(pub):
        sh = gc.open_by_key(SHEET_ID[pub])
        try: ws = sh.worksheet(TAB_ACTIVE_USERS)
        except Exception: return
        rowvals = [MONTH_NAME] + [int(round(site_wide_weekly_traffic[pub].get(f"Week {i}", 0))) for i in [1,2,3,4]]
        data = ws.get_all_values()
        target = next((i + 1 for i, r in enumerate(data) if r and r[0].strip().lower() == MONTH_NAME.lower()), None)
        if not dry_run:
            if target:
                ws.update(f"A{target}:E{target}", [rowvals], value_input_option="USER_ENTERED")
            else:
                ws.append_row(rowvals, value_input_option="USER_ENTERED")
        emit(f"{pub}: Active Users summary updated.")

    def write_followers(pub):
        sh = gc.open_by_key(SHEET_ID[pub])
        try: ws = sh.worksheet(TAB_FOLLOWERS)
        except gspread.WorksheetNotFound: return
        rows_to_write = []
        for w in [1, 2, 3, 4]:
            wk_str = f"Week {w}"
            g = weekly_follower_gains[pub].get(wk_str, {"LinkedIn": 0, "Twitter": 0, "Facebook": 0})
            rows_to_write.append([MONTH_NAME, wk_str, g["LinkedIn"], g["Twitter"], g["Facebook"]])
        data = ws.get_all_values()
        header = data[0] if data else []
        body = data[1:] if len(data) > 1 else []
        keep_body = [r for r in body if len(r) == 0 or r[0].strip().lower() != MONTH_NAME.lower()]
        if not dry_run:
            ws.clear()
            ws.update([header] + keep_body + rows_to_write, value_input_option="USER_ENTERED")
        emit(f"{pub}: Followers wrote 4 rows for {MONTH_NAME}.")

    for pub in GA4_PROPERTY:
        write_master(pub)
        write_active_users(pub)
        write_followers(pub)

    emit("Done. All sheets updated.")
    return {"month": MONTH_NAME, "year": YEAR, "log": log}


# ============================================================
# HTTP handler (Vercel Python runtime)
# ============================================================
class handler(BaseHTTPRequestHandler):
    def _send(self, code, payload):
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(payload).encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length).decode())
        except Exception as e:
            return self._send(400, {"error": f"Bad request: {e}"})

        access_token  = body.get("access_token")
        refresh_token = body.get("refresh_token")
        if not access_token:
            return self._send(401, {"error": "Missing access_token. Sign in with Google first."})

        creds = Credentials(
            token=access_token,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=os.environ.get("GOOGLE_CLIENT_ID"),
            client_secret=os.environ.get("GOOGLE_CLIENT_SECRET"),
            scopes=SCOPES,
        )

        import base64
        try:
            xls_b64 = body["file_b64"]
            xls_bytes = base64.b64decode(xls_b64.split(",")[-1])
            filename = body.get("filename", "upload.xlsx")
        except Exception as e:
            return self._send(400, {"error": f"Missing or invalid file: {e}"})

        month = body.get("month") or None
        year  = body.get("year") or None
        try:
            month = int(month) if month else None
            year  = int(year) if year else None
        except Exception:
            month, year = None, None

        try:
            result = run_pipeline(
                creds, xls_bytes, filename,
                year=year, month=month,
                normalise_seg=body.get("normalise_seg", True),
                write_mode=body.get("write_mode", "overwrite_month"),
                dry_run=body.get("dry_run", False),
            )
            return self._send(200, {"ok": True, **result})
        except Exception as e:
            return self._send(500, {"error": str(e)})
