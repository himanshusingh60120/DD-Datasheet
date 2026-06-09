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

def active_user_weeks(year, month):
    """
    Active-Users week scheme: exactly 4 weeks, Monday-Sunday aligned.
    Week 1 = day 1 through the FIRST Sunday (a short leading stub when the month
             does not start on Monday — it is NOT merged into the next week).
    Weeks 2 & 3 = full Mon-Sun 7-day blocks.
    Week 4 = the remainder of the month (absorbs every day after week 3 ends,
             including any trailing partial week).
    Example April 2026 (1st = Wed):
        W1 Apr 1-5, W2 Apr 6-12, W3 Apr 13-19, W4 Apr 20-30.
    Returns list of 4 (start_date, end_date) tuples.
    """
    ndays = calendar.monthrange(year, month)[1]
    first = dt.date(year, month, 1)
    last  = dt.date(year, month, ndays)

    # Week 1 ends on the first Sunday (or day 1 itself if the month starts Sunday)
    days_to_sunday = (6 - first.weekday()) % 7
    w1_end = first + dt.timedelta(days=days_to_sunday)
    if w1_end > last:
        w1_end = last

    bounds = [(first, w1_end)]
    cur = w1_end + dt.timedelta(days=1)
    for _ in range(2):                            # weeks 2 and 3: full Mon-Sun
        if cur > last:
            bounds.append((last, last)); continue
        wk_end = cur + dt.timedelta(days=6)
        if wk_end > last:
            wk_end = last
        bounds.append((cur, wk_end))
        cur = wk_end + dt.timedelta(days=1)
    bounds.append((cur, last) if cur <= last else (last, last))  # week 4 = remainder
    return bounds[:4]

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
                 normalise_seg=True, write_mode="overwrite_month", dry_run=False,
                 active_user_months=None):
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
            for wk_num in [1, 2, 3, 4]:
                wk_str = f"Week {wk_num}"
                wk_dates = [d for d in pub_dates if d.year == YEAR and d.month == MONTH and WEEK_MAP.get(d) == wk_str]
                if not wk_dates: continue
                start_date, end_date = min(wk_dates), max(wk_dates)
                for plat in ["LinkedIn", "Twitter", "Facebook"]:
                    start_val = daily_stats[pub][start_date].get(plat, 0)
                    end_val   = daily_stats[pub][end_date].get(plat, 0)
                    weekly_follower_gains[pub][wk_str][plat] = end_val - start_val
        emit("Follower deltas computed.")
    except Exception as e:
        emit(f"Followers skipped: {e}")

    # ---- GA4 fetch ----
    # Active Users uses its OWN week scheme (active_user_weeks): Week 1 is the
    # leading stub through the first Sunday, weeks 2-3 are full Mon-Sun, week 4
    # absorbs the remainder; always 4 weeks. activeUsers is DEDUPLICATED, so we
    # run one query per week over that week's date range and read the single
    # value GA4 returns (summing per-day would double-count repeat visitors).
    AU_WEEKS = active_user_weeks(YEAR, MONTH)   # [(start,end) x4]

    def ga_site_total(property_id, start, end):
        req = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[DateRange(start_date=start, end_date=end)],
            metrics=[Metric(name=GA_METRIC)],
        )
        resp = ga_client.run_report(req)
        return float(resp.rows[0].metric_values[0].value or 0) if resp.rows else 0.0

    def ga_by_page(property_id, start, end):
        req = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[DateRange(start_date=start, end_date=end)],
            dimensions=[Dimension(name=GA_DIMENSION)],
            metrics=[Metric(name=GA_METRIC)],
            limit=250000,
        )
        resp = ga_client.run_report(req)
        return [(row.dimension_values[0].value,
                 float(row.metric_values[0].value or 0)) for row in resp.rows]

    ga_week = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    site_wide_weekly_traffic = defaultdict(lambda: defaultdict(float))

    last_day = dt.date(YEAR, MONTH, calendar.monthrange(YEAR, MONTH)[1])
    yesterday = dt.date.today() - dt.timedelta(days=1)
    if last_day > yesterday:
        emit(f"⚠ {MONTH_NAME} {YEAR} not fully elapsed (GA4 data only through {yesterday}); later weeks may be partial.")

    for pub, pid in GA4_PROPERTY.items():
        dom = DOMAIN[pub]
        try:
            page_count = 0
            for idx, (w_start, w_end) in enumerate(AU_WEEKS, start=1):
                wk = f"Week {idx}"
                s = w_start.strftime("%Y-%m-%d")
                e = w_end.strftime("%Y-%m-%d")
                site_wide_weekly_traffic[pub][wk] = ga_site_total(pid, s, e)
                for path, val in ga_by_page(pid, s, e):
                    page_count += 1
                    key = dom + "/na" if str(path).strip().lower() == "(not set)" else normalise_url(path, dom)
                    if key is None: continue
                    ga_week[pub][key][wk] += val
            emit(f"GA4 {pub}: {page_count} page-rows | "
                 f"W1={int(round(site_wide_weekly_traffic[pub]['Week 1']))} "
                 f"W2={int(round(site_wide_weekly_traffic[pub]['Week 2']))} "
                 f"W3={int(round(site_wide_weekly_traffic[pub]['Week 3']))} "
                 f"W4={int(round(site_wide_weekly_traffic[pub]['Week 4']))}")
        except Exception as e:
            emit(f"GA4 error {pub}: {e}")

    # ---- Join individual links ----
    # The row's displayed "Week" stays as-is (original calendar logic). But GA4
    # per-link traffic is now keyed by the Active-Users week scheme, so we look up
    # using the AU week that CONTAINS the link's publish date.
    def au_week_for_date(d):
        for idx, (ws_, we_) in enumerate(AU_WEEKS, start=1):
            if ws_ <= d <= we_:
                return f"Week {idx}"
        return None

    for pub, rows in rows_by_pub.items():
        gw = ga_week.get(pub, {})
        for row in rows:
            au_wk = au_week_for_date(row["date"])
            au = gw.get(row["url_key"], {}).get(au_wk) if au_wk else None
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

    def write_active_users(pub, month_name=None, weekly_totals=None):
        """Write/update one Active-Users row: [MonthName, W1, W2, W3, W4].
        Defaults to the run's detected MONTH_NAME / site_wide_weekly_traffic,
        but can be called with an explicit month_name + weekly_totals dict
        (keys 'Week 1'..'Week 4') to fill any month independently."""
        mname = month_name or MONTH_NAME
        totals = weekly_totals if weekly_totals is not None else site_wide_weekly_traffic[pub]
        sh = gc.open_by_key(SHEET_ID[pub])
        try: ws = sh.worksheet(TAB_ACTIVE_USERS)
        except Exception: return
        rowvals = [mname] + [int(round(totals.get(f"Week {i}", 0))) for i in [1,2,3,4]]
        data = ws.get_all_values()
        target = next((i + 1 for i, r in enumerate(data) if r and r[0].strip().lower() == mname.lower()), None)
        if not dry_run:
            if target:
                ws.update(f"A{target}:E{target}", [rowvals], value_input_option="USER_ENTERED")
            else:
                ws.append_row(rowvals, value_input_option="USER_ENTERED")
        emit(f"{pub}: Active Users summary updated for {mname}.")

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

    def fetch_weekly_site_totals(pub, y, m):
        """Return {'Week 1':v,...,'Week 4':v} of GA4 site-wide active users
        for the given year/month, using the Active-Users week scheme."""
        pid = GA4_PROPERTY[pub]
        weeks = active_user_weeks(y, m)
        totals = {}
        for idx, (w_start, w_end) in enumerate(weeks, start=1):
            s = w_start.strftime("%Y-%m-%d")
            e = w_end.strftime("%Y-%m-%d")
            try:
                totals[f"Week {idx}"] = ga_site_total(pid, s, e)
            except Exception as ex:
                emit(f"GA4 error {pub} {calendar.month_name[m]} Week {idx}: {ex}")
                totals[f"Week {idx}"] = 0.0
        return totals

    # Determine which months get an Active-Users row. If the frontend sent a
    # list, use it (deduped, sorted in calendar order so the sheet stays
    # sequential). Otherwise fall back to the single detected/selected month.
    au_months = None
    if active_user_months:
        seen = []
        for mm in active_user_months:
            try: mm = int(mm)
            except Exception: continue
            if 1 <= mm <= 12 and mm not in seen:
                seen.append(mm)
        au_months = sorted(seen)
    if not au_months:
        au_months = [MONTH]

    for pub in GA4_PROPERTY:
        write_master(pub)
        # Active Users: one row per selected month, in calendar order.
        for mm in au_months:
            mname = calendar.month_name[mm]
            if mm == MONTH:
                # already fetched above for the detected month — reuse it
                totals = site_wide_weekly_traffic[pub]
            else:
                last_d = dt.date(YEAR, mm, calendar.monthrange(YEAR, mm)[1])
                if last_d > yesterday:
                    emit(f"⚠ {mname} {YEAR} not fully elapsed (GA4 only through {yesterday}); later weeks may be partial.")
                totals = fetch_weekly_site_totals(pub, YEAR, mm)
            write_active_users(pub, month_name=mname, weekly_totals=totals)
        write_followers(pub)

    emit("Done. All sheets updated.")
    return {"month": MONTH_NAME, "year": YEAR,
            "active_user_months": [calendar.month_name[m] for m in au_months],
            "log": log}


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

        active_user_months = body.get("active_user_months") or None
        if active_user_months and not isinstance(active_user_months, list):
            active_user_months = [active_user_months]

        try:
            result = run_pipeline(
                creds, xls_bytes, filename,
                year=year, month=month,
                normalise_seg=body.get("normalise_seg", True),
                write_mode=body.get("write_mode", "overwrite_month"),
                dry_run=body.get("dry_run", False),
                active_user_months=active_user_months,
            )
            return self._send(200, {"ok": True, **result})
        except Exception as e:
            return self._send(500, {"error": str(e)})
