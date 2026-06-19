"""
Weekly Cost Analysis — Packaging Material Dashboard (HR1 / CH1 / MU1)
Reads Summary + Daily Use PMS from Google Sheets, generates multi-warehouse HTML.
"""
import json, logging, re, sys
from datetime import datetime
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger(__name__)
SCRIPT_DIR = Path(__file__).parent

# ── auth ──────────────────────────────────────────────────────────────────────

def get_gc():
    try:
        import gspread
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.oauth2.credentials import Credentials as OC
        from google.auth.transport.requests import Request
    except ImportError:
        log.error("Missing packages. Run: pip install -r requirements.txt")
        sys.exit(1)

    SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    token_path = SCRIPT_DIR / "token.json"
    creds_path = SCRIPT_DIR / "credentials.json"

    creds = None
    if token_path.exists():
        creds = OC.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    return gspread.authorize(creds)


def find_worksheet(sh, name):
    """Find worksheet by name, ignoring leading/trailing whitespace."""
    for ws in sh.worksheets():
        if ws.title.strip().lower() == name.strip().lower():
            return ws
    available = [w.title for w in sh.worksheets()]
    raise ValueError(f"Worksheet '{name}' not found. Available: {available}")


# ── helpers ───────────────────────────────────────────────────────────────────

def safe_float(v):
    if v is None: return 0.0
    s = re.sub(r"[^\d.\-]", "", str(v))
    try: return float(s)
    except: return 0.0


def get_category(name):
    n = name.lower()
    if "temper" in n: return "Temper Proof"
    if "zipper" in n or "strapper" in n: return "Zipper Bag"
    if "corrugated" in n or "alfa" in n: return "Corrugated Box"
    if "shrink" in n: return "Shrink Wrap"
    if "thermosheet" in n: return "Thermosheet"
    if "gunny" in n: return "Gunny Bag"
    if "bubble" in n: return "Bubble Roll"
    return "Other"


IS_DATE = re.compile(r"^\d{1,2}-[A-Za-z]+$")

# Week definitions for June 2026
WEEKS = [
    ("W1",  "1–7 Jun",   ["1-June","2-June","3-June","4-June","5-June","6-June","7-June"]),
    ("W2",  "8–14 Jun",  ["8-June","9-June","10-June","11-June","12-June","13-June","14-June"]),
    ("W3",  "15–21 Jun", ["15-June","16-June","17-June","18-June","19-June","20-June","21-June"]),
    ("W4",  "22–30 Jun", ["22-June","23-June","24-June","25-June","26-June","27-June","28-June","29-June","30-June","1-July"]),
]

CAT_COLORS = {
    "Corrugated Box": "#378ADD",
    "Temper Proof":   "#D85A30",
    "Zipper Bag":     "#1D9E75",
    "Shrink Wrap":    "#7F77DD",
    "Thermosheet":    "#BA7517",
    "Gunny Bag":      "#639922",
    "Bubble Roll":    "#E8A838",
    "Other":          "#888780",
}


# ── fetch data ────────────────────────────────────────────────────────────────

def fetch_all(config):
    gc = get_gc()
    sh = gc.open_by_key(config["spreadsheet_id"])

    log.info("Reading Summary…")
    sum_rows = sh.worksheet("Summary").get_all_values()

    log.info("Reading Daily Use PMS…")
    daily_rows = sh.worksheet(config["daily_use_sheet"]).get_all_values()

    return sum_rows, daily_rows


# ── parse Summary → price map ─────────────────────────────────────────────────

def parse_prices(sum_rows):
    """Returns dict: sku_desc -> {unit_price, tax, consumed_value}"""
    prices = {}
    hdr = sum_rows[0] if sum_rows else []

    # locate columns
    def ci(name):
        for i, h in enumerate(hdr):
            if name.lower() in h.lower():
                return i
        return -1

    c_desc  = ci("sku desc") if ci("sku desc") >= 0 else 1
    c_price = ci("unit price") if ci("unit price") >= 0 else 5
    c_tax   = ci("tax") if ci("tax") >= 0 else 6
    c_cval  = ci("consumed stock value") if ci("consumed stock value") >= 0 else 13

    for r in sum_rows[1:]:
        if not r or len(r) <= max(c_desc, c_price, c_tax, c_cval):
            continue
        desc  = str(r[c_desc]).strip()
        price = safe_float(r[c_price])
        tax   = safe_float(r[c_tax])
        cval  = safe_float(r[c_cval])
        if not desc:
            continue
        if desc not in prices:
            prices[desc] = {"unit_price": price, "tax": tax, "consumed_value": cval}
        else:
            # Multiple vendor rows for same SKU desc — sum consumed values,
            # keep weighted-average price based on consumed value
            prices[desc]["consumed_value"] += cval
            if price > 0:
                prices[desc]["unit_price"] = price  # use latest non-zero price

    log.info(f"Price map: {len(prices)} SKUs")
    return prices


# ── parse historical month totals from col V / W ─────────────────────────────

MONTH_ORDER = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
               "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}

def parse_historical(sum_rows):
    """
    Reads col V (Month) and col W (Consumed Value) from Summary.
    Returns list of {month_label, value} sorted chronologically.
    """
    results = {}
    for r in sum_rows[1:]:
        if len(r) <= 22:
            continue
        label = str(r[21]).strip()
        val   = safe_float(r[22])
        if label and val > 0 and label not in results:
            results[label] = val

    def sort_key(item):
        parts = item[0].replace("-", " ").split()
        mon = MONTH_ORDER.get(parts[0][:3].lower(), 0) if parts else 0
        yr  = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
        return (yr, mon)

    return sorted(results.items(), key=sort_key)


# ── parse Daily Use PMS → materials ──────────────────────────────────────────

def parse_daily(daily_rows):
    """Returns (materials_list, date_labels_list)"""
    if not daily_rows:
        return [], []

    hdr = daily_rows[0]

    # Find description column by header name; fall back to col index 1
    desc_col = 1
    for i, h in enumerate(hdr):
        hl = str(h).lower().strip()
        if "sku desc" in hl or hl == "description":
            desc_col = i
            break

    date_cols = {}  # label -> col_index
    for i, h in enumerate(hdr):
        if IS_DATE.match(str(h).strip()):
            date_cols[str(h).strip()] = i

    date_labels = list(date_cols.keys())
    materials = []

    for r in daily_rows[1:]:
        if len(r) < 2:
            continue
        desc = str(r[desc_col]).strip() if desc_col < len(r) and r[desc_col] else ""
        if not desc:
            continue

        daily = {}
        for label, ci in date_cols.items():
            val = safe_float(r[ci]) if ci < len(r) else 0.0
            if val > 0:
                daily[label] = val

        materials.append({
            "name":     desc,
            "category": get_category(desc),
            "daily":    daily,
            "total":    sum(daily.values()),
        })

    log.info(f"Daily: {len(materials)} materials, {len(date_labels)} date columns")
    return materials, date_labels


# ── compute weekly costs ──────────────────────────────────────────────────────

def compute_weekly(materials, prices):
    """
    Returns enriched materials list with weekly cost breakdown.
    Weekly cost = consumed_value (col13 from Summary) × (week_qty / total_qty)
    This ensures totals match the pre-calculated Summary values exactly.
    """
    enriched = []
    for m in materials:
        p = prices.get(m["name"], {})
        unit_price      = p.get("unit_price", 0.0)
        tax             = p.get("tax", 0.0)
        consumed_value  = p.get("consumed_value", 0.0)  # pre-calculated, ex-tax

        # If no consumed_value in Summary, fall back to qty × price (no tax — col13 is ex-tax)
        if consumed_value == 0 and unit_price > 0 and m["total"] > 0:
            consumed_value = m["total"] * unit_price

        week_data = {}
        total_cost = 0.0
        for wk, label, dates in WEEKS:
            qty  = sum(m["daily"].get(d, 0.0) for d in dates)
            # Proportional split: week's share of consumed_value
            prop = qty / m["total"] if m["total"] > 0 else 0.0
            cost = consumed_value * prop
            week_data[wk] = {"qty": qty, "cost": cost}
            total_cost += cost

        enriched.append({
            **m,
            "unit_price":  unit_price,
            "tax":         tax,
            "week_data":   week_data,
            "total_cost":  total_cost,
        })

    return enriched


# ── aggregate stats ───────────────────────────────────────────────────────────

def aggregate(enriched):
    # weekly totals
    week_totals = {wk: 0.0 for wk, *_ in WEEKS}
    cat_week    = {}  # cat -> {wk: cost}
    active_weeks = set()

    for m in enriched:
        cat = m["category"]
        if cat not in cat_week:
            cat_week[cat] = {wk: 0.0 for wk, *_ in WEEKS}
        for wk, _, _ in WEEKS:
            c = m["week_data"][wk]["cost"]
            week_totals[wk] += c
            cat_week[cat][wk] += c
            if c > 0:
                active_weeks.add(wk)

    return week_totals, cat_week, active_weeks


# ── HTML generation ───────────────────────────────────────────────────────────

def fmt_inr(v):
    """Format float as Indian rupee string: ₹1,23,456"""
    v = int(round(v))
    if v == 0:
        return "₹0"
    s = str(v)
    if len(s) <= 3:
        return f"₹{s}"
    last3 = s[-3:]
    rest = s[:-3]
    parts = []
    while len(rest) > 2:
        parts.append(rest[-2:])
        rest = rest[:-2]
    if rest:
        parts.append(rest)
    return "₹" + ",".join(reversed(parts)) + "," + last3


def generate_html(enriched, week_totals, cat_week, generated_at, historical=None):
    # Chart.js inline
    chartjs_path = SCRIPT_DIR / "chart.umd.js"
    if chartjs_path.exists():
        chartjs_tag = f"<script>{chartjs_path.read_text(encoding='utf-8')}</script>"
    else:
        chartjs_tag = '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>'

    week_labels  = [label for _, label, _ in WEEKS]
    week_keys    = [wk for wk, *_ in WEEKS]
    week_vals    = [round(week_totals[wk], 2) for wk in week_keys]

    total_month  = sum(week_vals)

    # top 20 by cost
    top20 = sorted([m for m in enriched if m["total_cost"] > 0],
                   key=lambda x: -x["total_cost"])[:20]

    # categories sorted by total cost
    cats_sorted = sorted(cat_week.keys(),
                         key=lambda c: -sum(cat_week[c].values()))

    # build serialisable lists outside f-strings to avoid {{ }} ambiguity
    top20_json = json.dumps([
        {"name": m["name"], "category": m["category"],
         "unit_price": m["unit_price"], "total_cost": m["total_cost"]}
        for m in top20
    ])
    all_mat_rows = sorted([m for m in enriched if m["total_cost"] > 0],
                          key=lambda x: -x["total_cost"])
    all_mat_json = json.dumps([
        {
            "name":       m["name"],
            "category":   m["category"],
            "unit_price": round(m["unit_price"], 2),
            "tax":        m["tax"],
            "week_data":  {wk: {"qty": round(v["qty"], 0), "cost": round(v["cost"], 2)}
                           for wk, v in m["week_data"].items()},
            "total_cost": round(m["total_cost"], 2),
        }
        for m in all_mat_rows
    ])

    # JS data
    data_block = (
        f"const WEEK_LABELS = {json.dumps(week_labels)};\n"
        f"const WEEK_VALS = {json.dumps(week_vals)};\n"
        f"const WEEK_KEYS = {json.dumps(week_keys)};\n"
        f"const CAT_WEEK = {json.dumps(cat_week)};\n"
        f"const CATS_SORTED = {json.dumps(cats_sorted)};\n"
        f"const TOP20 = {top20_json};\n"
        f"const ALL_MAT = {all_mat_json};\n"
        f"const GEN_TIME = {json.dumps(generated_at.strftime('%d %b %Y, %I:%M %p'))};\n"
        f"const TOTAL_MONTH = {round(total_month, 2)};\n"
    )

    # KPI values
    w1 = fmt_inr(week_totals.get("W1", 0))
    w2 = fmt_inr(week_totals.get("W2", 0))
    w3 = fmt_inr(week_totals.get("W3", 0))
    w4 = fmt_inr(week_totals.get("W4", 0))

    # MRR = (cost to date / days elapsed) × total days in month
    import calendar
    days_in_month = calendar.monthrange(generated_at.year, generated_at.month)[1]
    days_elapsed  = generated_at.day
    mrr = (total_month / days_elapsed) * days_in_month if days_elapsed > 0 else 0

    # historical months from Summary col V/W — last 4 entries
    hist_list = (historical or [])[-4:]  # keep last 4 months chronologically

    def pct_change(new, old):
        if old == 0: return None
        return (new - old) / old * 100

    def pct_badge(pct):
        if pct is None: return ""
        arrow = "▲" if pct > 0 else "▼"
        clr   = "#E24B4A" if pct > 0 else "#1D9E75"
        return f'<span style="font-size:11px;color:{clr};margin-left:6px">{arrow} {abs(pct):.1f}%</span>'

    # build cards html for historical months
    hist_cards_html = ""
    for i, (label, val) in enumerate(hist_list):
        prev_val = hist_list[i-1][1] if i > 0 else None
        pct      = pct_change(val, prev_val) if prev_val else None
        badge    = pct_badge(pct)
        hist_cards_html += (
            f'<div class="kpi" style="border:.5px solid rgba(55,138,221,0.3)">'
            f'<label style="color:#378ADD">{label} Final</label>'
            f'<span style="color:#378ADD">{fmt_inr(val)}</span>{badge}'
            f'</div>'
        )
    # pad if fewer than 4
    while hist_cards_html.count('<div class="kpi"') < 4:
        hist_cards_html += '<div class="kpi" style="opacity:0.3"><label>—</label><span>—</span></div>'

    # june MRR vs last historical month final
    may_val_num = hist_list[-1][1] if hist_list else 0
    jun_pct     = pct_change(mrr, may_val_num)
    jun_badge   = pct_badge(jun_pct)

    html = _html_template(chartjs_tag)
    return (html
        .replace("##DATA_BLOCK##",    data_block)
        .replace("##GEN_TIME##",      generated_at.strftime("%d %b %Y, %I:%M %p"))
        .replace("##TOTAL_MONTH##",   fmt_inr(total_month))
        .replace("##JUN_BADGE##",     jun_badge)
        .replace("##MRR##",           fmt_inr(mrr))
        .replace("##MRR_DETAIL##",    f"({days_elapsed} of {days_in_month} days)")
        .replace("##W1##", w1)
        .replace("##W2##", w2)
        .replace("##W3##", w3)
        .replace("##W4##", w4)
        .replace("##HIST_CARDS##",    hist_cards_html)
        .replace("##CHARTJS##",       chartjs_tag)
    )


def _compute_wh_display(enriched, week_totals, cat_week, historical, generated_at):
    """Compute all formatted display values for one warehouse."""
    import calendar

    week_labels = [label for _, label, _ in WEEKS]
    week_keys   = [wk    for wk,  *_    in WEEKS]
    week_vals   = [round(week_totals[wk], 2) for wk in week_keys]
    total_month = sum(week_vals)

    top20 = sorted([m for m in enriched if m["total_cost"] > 0],
                   key=lambda x: -x["total_cost"])[:20]
    cats_sorted = sorted(cat_week.keys(), key=lambda c: -sum(cat_week[c].values()))

    top20_json = json.dumps([
        {"name": m["name"], "category": m["category"],
         "unit_price": m["unit_price"], "total_cost": m["total_cost"]}
        for m in top20
    ])
    all_mat_json = json.dumps([
        {"name": m["name"], "category": m["category"],
         "unit_price": round(m["unit_price"], 2), "tax": m["tax"],
         "week_data": {wk: {"qty": round(v["qty"], 0), "cost": round(v["cost"], 2)}
                       for wk, v in m["week_data"].items()},
         "total_cost": round(m["total_cost"], 2)}
        for m in sorted([m for m in enriched if m["total_cost"] > 0],
                        key=lambda x: -x["total_cost"])
    ])

    days_in_month = calendar.monthrange(generated_at.year, generated_at.month)[1]
    days_elapsed  = generated_at.day
    mrr = (total_month / days_elapsed) * days_in_month if days_elapsed > 0 else 0

    hist_list = (historical or [])[-4:]

    def pct_change(new, old): return (new - old) / old * 100 if old else None
    def pct_badge(pct):
        if pct is None: return ""
        arrow = "▲" if pct > 0 else "▼"
        clr   = "#E24B4A" if pct > 0 else "#1D9E75"
        return f'<span style="font-size:11px;color:{clr};margin-left:6px">{arrow} {abs(pct):.1f}%</span>'

    hist_cards_html = ""
    for i, (label, val) in enumerate(hist_list):
        prev_val = hist_list[i-1][1] if i > 0 else None
        badge    = pct_badge(pct_change(val, prev_val) if prev_val else None)
        hist_cards_html += (
            f'<div class="kpi" style="border:.5px solid rgba(55,138,221,0.3)">'
            f'<label style="color:#378ADD">{label} Final</label>'
            f'<span style="color:#378ADD">{fmt_inr(val)}</span>{badge}</div>'
        )
    while hist_cards_html.count('<div class="kpi"') < 4:
        hist_cards_html += '<div class="kpi" style="opacity:0.3"><label>—</label><span>—</span></div>'

    may_val_num = hist_list[-1][1] if hist_list else 0
    jun_badge   = pct_badge(pct_change(mrr, may_val_num))

    return {
        "week_labels":     week_labels,
        "week_keys":       week_keys,
        "week_vals":       week_vals,
        "total_month":     total_month,
        "total_month_fmt": fmt_inr(total_month),
        "cat_week":        cat_week,
        "cats_sorted":     cats_sorted,
        "top20_json":      top20_json,
        "all_mat_json":    all_mat_json,
        "mrr_fmt":         fmt_inr(mrr),
        "mrr_detail":      f"({days_elapsed} of {days_in_month} days)",
        "w1": fmt_inr(week_totals.get("W1", 0)),
        "w2": fmt_inr(week_totals.get("W2", 0)),
        "w3": fmt_inr(week_totals.get("W3", 0)),
        "w4": fmt_inr(week_totals.get("W4", 0)),
        "hist_cards_html": hist_cards_html,
        "jun_badge":        jun_badge,
    }


def generate_multi_html(wh_results, generated_at):
    """
    wh_results: {wh_code: {enriched, week_totals, cat_week, historical}}
    Returns a single HTML page with HR1 / CH1 / MU1 tab navigation.
    """
    chartjs_path = SCRIPT_DIR / "chart.umd.js"
    if chartjs_path.exists():
        chartjs_tag = f"<script>{chartjs_path.read_text(encoding='utf-8')}</script>"
    else:
        chartjs_tag = '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>'

    gen_time_str = generated_at.strftime("%d %b %Y, %I:%M %p")
    wh_list      = list(wh_results.keys())
    first_wh     = wh_list[0]

    kpis = {}
    for wh, data in wh_results.items():
        kpis[wh] = _compute_wh_display(
            data["enriched"], data["week_totals"], data["cat_week"],
            data.get("historical", []), generated_at
        )

    # ── tab buttons ──────────────────────────────────────────────────────────
    tab_buttons_html = ""
    for wh in wh_list:
        active = " active" if wh == first_wh else ""
        tab_buttons_html += (
            f'<button id="tab-{wh}" class="tab-btn{active}" '
            f'onclick="switchTab(\'{wh}\')">{wh}</button>'
        )

    # ── per-warehouse sections ────────────────────────────────────────────────
    sections_html = ""
    for wh, k in kpis.items():
        vis = "" if wh == first_wh else ' style="display:none"'
        sections_html += (
            f'<div id="section-{wh}" class="wh-section"{vis}>\n'
            f'<div class="kpi-top">'
            f'<div class="kpi-left">'
            f'<div class="kpi accent"><label>Total Month Cost · Jun</label>'
            f'<span>{k["total_month_fmt"]}</span>{k["jun_badge"]}</div>'
            f'<div class="kpi" style="border:.5px solid rgba(239,159,39,0.4)">'
            f'<label style="color:#EF9F27">MRR {k["mrr_detail"]}</label>'
            f'<span style="color:#EF9F27">{k["mrr_fmt"]}</span></div>'
            f'</div>'
            f'<div class="kpi-weeks">'
            f'<div class="kpi"><label>Week 1 · 1–7 Jun</label><span>{k["w1"]}</span></div>'
            f'<div class="kpi"><label>Week 2 · 8–14 Jun</label><span>{k["w2"]}</span></div>'
            f'<div class="kpi"><label>Week 3 · 15–21 Jun</label><span>{k["w3"]}</span></div>'
            f'<div class="kpi"><label>Week 4 · 22–30 Jun</label><span>{k["w4"]}</span></div>'
            f'</div></div>\n'
            f'<div class="kpi-hist">{k["hist_cards_html"]}</div>\n'
            f'<div class="two">'
            f'<div class="card"><h2>Weekly total cost trend</h2>'
            f'<div style="position:relative;height:260px"><canvas id="{wh}_trendChart"></canvas></div></div>'
            f'<div class="card"><h2>Cost by category — weekly</h2>'
            f'<div class="leg" id="{wh}_catLegend"></div>'
            f'<div style="position:relative;height:220px"><canvas id="{wh}_catStackChart"></canvas></div></div>'
            f'</div>\n'
            f'<div class="card"><h2>Top 20 materials — total consumption cost (June)</h2>'
            f'<p class="sub">Cost = qty × unit price × (1 + tax)</p>'
            f'<div style="position:relative;height:340px"><canvas id="{wh}_top20Chart"></canvas></div></div>\n'
            f'<div class="card"><h2>Weekly cost breakdown by material</h2>'
            f'<p class="sub">Sorted by total cost · only materials with consumption shown</p>'
            f'<div style="overflow-x:auto"><table class="wt">'
            f'<thead><tr>'
            f'<th>Material</th><th>Category</th><th class="r">Unit Price</th>'
            f'<th class="r">W1 Qty</th><th class="r">W1 Cost</th>'
            f'<th class="r">W2 Qty</th><th class="r">W2 Cost</th>'
            f'<th class="r">W3 Qty</th><th class="r">W3 Cost</th>'
            f'<th class="r">W4 Qty</th><th class="r">W4 Cost</th>'
            f'<th class="r">Total Cost</th>'
            f'</tr></thead><tbody id="{wh}_costTbody"></tbody>'
            f'</table></div></div>\n'
            f'</div>\n'
        )

    # ── JS WH_DATA object ─────────────────────────────────────────────────────
    wh_data_parts = []
    for wh, k in kpis.items():
        wh_data_parts.append(
            f'"{wh}":{{'
            f'"week_labels":{json.dumps(k["week_labels"])},'
            f'"week_vals":{json.dumps(k["week_vals"])},'
            f'"week_keys":{json.dumps(k["week_keys"])},'
            f'"cat_week":{json.dumps(k["cat_week"])},'
            f'"cats_sorted":{json.dumps(k["cats_sorted"])},'
            f'"top20":{k["top20_json"]},'
            f'"all_mat":{k["all_mat_json"]}'
            f'}}'
        )
    wh_data_js   = "const WH_DATA={" + ",".join(wh_data_parts) + "};"
    first_wh_str = json.dumps(first_wh)

    return (_multi_html_template()
        .replace("##GEN_TIME##",       gen_time_str)
        .replace("##TAB_BUTTONS##",    tab_buttons_html)
        .replace("##SECTIONS_HTML##",  sections_html)
        .replace("##WH_DATA_JS##",     wh_data_js)
        .replace("##FIRST_WH##",       first_wh_str)
        .replace("##CHARTJS##",        chartjs_tag)
    )


def _multi_html_template():
    return r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<meta http-equiv="refresh" content="3600">
<title>Packaging Material Dashboard — HR1 · CH1 · MU1</title>
<style>
:root{--bg:#161614;--bg2:#202020;--bg3:#1a1a18;--text:#e4e2dc;--text2:#8a8880;--text3:#555550;--border:rgba(255,255,255,0.10);--border2:rgba(255,255,255,0.18);--r:12px;--rs:8px;--font:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;--green:#1D9E75;--orange:#EF9F27;--red:#E24B4A;--blue:#378ADD}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--font);background:var(--bg);color:var(--text);padding:1.5rem 1.75rem;min-height:100vh}
h1{font-size:20px;font-weight:500}
h2{font-size:13px;font-weight:500;margin:0 0 12px;color:var(--text)}
.hdr{display:flex;justify-content:space-between;align-items:center;margin-bottom:1rem;padding:1rem 1.25rem;background:var(--bg2);border-radius:var(--r);border:.5px solid var(--border)}
.hdr-r{font-size:12px;color:var(--text2);text-align:right;line-height:1.7}
.tab-bar{display:flex;gap:6px;margin-bottom:1.25rem}
.tab-btn{background:var(--bg2);border:.5px solid var(--border);color:var(--text2);padding:8px 22px;border-radius:8px;cursor:pointer;font-family:var(--font);font-size:14px;font-weight:600;letter-spacing:.03em;transition:all .15s}
.tab-btn:hover{border-color:var(--blue);color:var(--text)}
.tab-btn.active{background:var(--blue);border-color:var(--blue);color:#fff}
.kpi-top{display:grid;grid-template-columns:auto 1fr;gap:10px;margin-bottom:10px;align-items:stretch}
.kpi-left{display:flex;flex-direction:column;gap:10px;min-width:170px}
.kpi-weeks{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.kpi-hist{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:1.25rem}
.kpi{background:var(--bg2);border-radius:var(--rs);padding:.9rem 1rem;border:.5px solid var(--border)}
.kpi label{font-size:11px;color:var(--text2);display:block;margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em}
.kpi span{font-size:22px;font-weight:500}
.kpi.accent span{color:var(--green)}
.card{background:var(--bg2);border-radius:var(--r);padding:1rem 1.25rem;border:.5px solid var(--border);margin-bottom:1.1rem}
.two{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:1.1rem}
table.wt{width:100%;border-collapse:collapse;font-size:12px}
table.wt th{padding:7px 10px;text-align:left;color:var(--text2);font-weight:500;border-bottom:1px solid var(--border2);white-space:nowrap}
table.wt th.r,table.wt td.r{text-align:right}
table.wt td{padding:6px 10px;border-bottom:.5px solid var(--border);color:var(--text)}
table.wt tr:last-child td{border-bottom:none}
table.wt tr:hover td{background:rgba(255,255,255,0.03)}
.badge{display:inline-block;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:500}
.leg{display:flex;flex-wrap:wrap;gap:6px 14px;margin-bottom:10px}
.leg-item{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--text2)}
.sub{font-size:11px;color:var(--text2);margin:-8px 0 10px}
.icon-btn{display:inline-flex;align-items:center;gap:6px;background:var(--bg2);border:.5px solid var(--border2);color:var(--text2);padding:7px 13px;border-radius:8px;cursor:pointer;font-family:var(--font);font-size:12px;font-weight:500;transition:all .15s;white-space:nowrap}
.icon-btn:hover{border-color:var(--blue);color:var(--blue)}
body.light{--bg:#f0efe9;--bg2:#ffffff;--bg3:#f5f4f0;--text:#1a1918;--text2:#6a6860;--text3:#a09e98;--border:rgba(0,0,0,0.08);--border2:rgba(0,0,0,0.14)}
body.light .tab-btn{background:#fff;color:var(--text2)}
body.light .tab-btn:hover{color:var(--text)}
body.light .tab-btn.active{color:#fff}
body.light table.wt tr:hover td{background:rgba(0,0,0,0.02)}
</style>
</head>
<body>

<div class="hdr">
  <div>
    <h1>Packaging Material · Weekly Cost Analysis</h1>
    <p style="font-size:12px;color:var(--text2);margin-top:3px">June 2026</p>
  </div>
  <div class="hdr-r" style="display:flex;align-items:center;gap:10px">
    <div style="text-align:right">
      Last updated: ##GEN_TIME##<br>
      <span style="color:var(--text3)">Auto-refreshes every hour</span>
    </div>
    <a href="/refresh" class="icon-btn" title="Refresh data now" style="text-decoration:none">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
      Refresh
    </a>
    <button class="icon-btn" onclick="toggleTheme()" id="themeBtn" title="Toggle light/dark mode">
      <svg id="themeIcon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>
      <span id="themeLabel">Light</span>
    </button>
  </div>
</div>

<div class="tab-bar">##TAB_BUTTONS##</div>

##SECTIONS_HTML##

<p style="font-size:11px;color:var(--text3);text-align:center;margin-top:1.5rem;padding-top:1rem;border-top:.5px solid var(--border)">
  Packaging Material Monitor &nbsp;·&nbsp; ##GEN_TIME## &nbsp;·&nbsp; Source: Google Sheets
</p>

##CHARTJS##
<script>
window.onerror=function(msg,src,line,col,err){document.body.insertAdjacentHTML('afterbegin','<div style="position:fixed;top:0;left:0;right:0;background:#E24B4A;color:#fff;padding:8px 12px;font-family:monospace;font-size:12px;z-index:9999">JS Error: '+msg+' (line '+line+')</div>');return false;};

##WH_DATA_JS##

const CC={"Corrugated Box":"#378ADD","Temper Proof":"#D85A30","Zipper Bag":"#1D9E75","Other":"#888780","Shrink Wrap":"#7F77DD","Thermosheet":"#BA7517","Gunny Bag":"#639922","Bubble Roll":"#E8A838"};

function fmtINR(v){v=Math.round(v);if(v===0)return"₹0";let s=v.toString(),last3=s.slice(-3),rest=s.slice(0,-3),parts=[];while(rest.length>2){parts.unshift(rest.slice(-2));rest=rest.slice(0,-2);}if(rest)parts.unshift(rest);return"₹"+parts.join(",")+","+last3;}

const chartsInited={};

function initCharts(wh){
  const d=WH_DATA[wh];

  const legEl=document.getElementById(wh+"_catLegend");
  if(legEl)d.cats_sorted.forEach(cat=>{
    const s=document.createElement("span");s.className="leg-item";
    s.innerHTML='<span style="width:8px;height:8px;border-radius:2px;background:'+(CC[cat]||"#888")+';display:inline-block"></span>'+cat;
    legEl.appendChild(s);
  });

  try{new Chart(document.getElementById(wh+"_trendChart"),{
    type:"bar",
    data:{labels:d.week_labels,datasets:[{label:"Cost (₹)",data:d.week_vals,
      backgroundColor:d.week_vals.map((_,i)=>["#378ADD","#1D9E75","#D85A30","#EF9F27"][i]||"#888"),
      borderRadius:6,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>" "+fmtINR(c.parsed.y)}}},
      scales:{x:{grid:{color:"rgba(136,135,128,0.1)"},ticks:{color:"#8a8880"}},
              y:{grid:{color:"rgba(136,135,128,0.1)"},ticks:{color:"#8a8880",callback:v=>fmtINR(v)}}}}
  });}catch(e){console.error(wh+"_trendChart:",e);}

  try{const catDs=d.cats_sorted.map(cat=>({
    label:cat,data:d.week_keys.map(wk=>(d.cat_week[cat]||{})[wk]||0),
    backgroundColor:CC[cat]||"#888",borderWidth:0,borderRadius:3
  }));
  new Chart(document.getElementById(wh+"_catStackChart"),{
    type:"bar",data:{labels:d.week_labels,datasets:catDs},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.dataset.label+": "+fmtINR(c.parsed.y)}}},
      scales:{x:{stacked:true,grid:{display:false},ticks:{color:"#8a8880"}},
              y:{stacked:true,grid:{color:"rgba(136,135,128,0.1)"},ticks:{color:"#8a8880",callback:v=>fmtINR(v)}}}}
  });}catch(e){console.error(wh+"_catStackChart:",e);}

  try{new Chart(document.getElementById(wh+"_top20Chart"),{
    type:"bar",
    data:{labels:d.top20.map(m=>m.name),datasets:[{data:d.top20.map(m=>m.total_cost),
      backgroundColor:d.top20.map(m=>CC[m.category]||"#888"),borderRadius:4,borderSkipped:false}]},
    options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>" "+fmtINR(c.parsed.x)}}},
      scales:{x:{grid:{color:"rgba(136,135,128,0.1)"},ticks:{color:"#8a8880",callback:v=>fmtINR(v)}},
              y:{grid:{display:false},ticks:{color:"#8a8880",font:{size:11}}}}}
  });}catch(e){console.error(wh+"_top20Chart:",e);}

  const tbody=document.getElementById(wh+"_costTbody");
  if(tbody)d.all_mat.forEach((m,idx)=>{
    const tr=document.createElement("tr");
    if(idx%2===1)tr.style.background="rgba(255,255,255,0.02)";
    const cc=CC[m.category]||"#888";
    let cells="<td>"+m.name+"</td>";
    cells+='<td><span class="badge" style="background:'+cc+'22;color:'+cc+'">'+m.category+"</span></td>";
    cells+='<td class="r" style="color:var(--text2)">₹'+m.unit_price+"</td>";
    ["W1","W2","W3","W4"].forEach(wk=>{
      const dd=m.week_data[wk]||{qty:0,cost:0};
      cells+='<td class="r">'+(dd.qty||"")+"</td>";
      cells+='<td class="r" style="color:'+(dd.cost>0?"var(--text)":"var(--text3)")+'">'+
             (dd.cost>0?fmtINR(dd.cost):"")+"</td>";
    });
    cells+='<td class="r" style="font-weight:500">'+fmtINR(m.total_cost)+"</td>";
    tr.innerHTML=cells;tbody.appendChild(tr);
  });
}

function switchTab(wh){
  document.querySelectorAll('.wh-section').forEach(s=>s.style.display='none');
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('section-'+wh).style.display='';
  document.getElementById('tab-'+wh).classList.add('active');
  if(!chartsInited[wh]){initCharts(wh);chartsInited[wh]=true;}
}

initCharts(##FIRST_WH##);
chartsInited[##FIRST_WH##]=true;

// ── theme toggle ──────────────────────────────────────────────────────────────
const moonSVG='<circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>';
const sunSVG='<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>';
function applyTheme(light){
  document.body.classList.toggle('light',light);
  document.getElementById('themeIcon').innerHTML=light?sunSVG:moonSVG;
  document.getElementById('themeLabel').textContent=light?'Dark':'Light';
}
function toggleTheme(){
  const isLight=document.body.classList.toggle('light');
  localStorage.setItem('pmsDashTheme',isLight?'light':'dark');
  applyTheme(isLight);
}
applyTheme(localStorage.getItem('pmsDashTheme')==='light');
</script>
</body>
</html>
"""


def _html_template(chartjs_tag):
    return r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<meta http-equiv="refresh" content="3600">
<title>HR1 — Weekly Cost Analysis · Packaging Material</title>
<style>
:root{--bg:#161614;--bg2:#202020;--bg3:#1a1a18;--text:#e4e2dc;--text2:#8a8880;--text3:#555550;--border:rgba(255,255,255,0.10);--border2:rgba(255,255,255,0.18);--r:12px;--rs:8px;--font:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;--green:#1D9E75;--orange:#EF9F27;--red:#E24B4A;--blue:#378ADD}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--font);background:var(--bg);color:var(--text);padding:1.5rem 1.75rem;min-height:100vh}
h1{font-size:20px;font-weight:500}
h2{font-size:13px;font-weight:500;margin:0 0 12px;color:var(--text)}
.hdr{display:flex;justify-content:space-between;align-items:center;margin-bottom:1.25rem;padding:1rem 1.25rem;background:var(--bg2);border-radius:var(--r);border:.5px solid var(--border)}
.hdr-r{font-size:12px;color:var(--text2);text-align:right;line-height:1.7}
.kpi-top{display:grid;grid-template-columns:auto 1fr;gap:10px;margin-bottom:10px;align-items:stretch}
.kpi-left{display:flex;flex-direction:column;gap:10px;min-width:170px}
.kpi-weeks{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.kpi-hist{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:1.25rem}
.kpi{background:var(--bg2);border-radius:var(--rs);padding:.9rem 1rem;border:.5px solid var(--border)}
.kpi label{font-size:11px;color:var(--text2);display:block;margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em}
.kpi span{font-size:22px;font-weight:500}
.kpi.accent span{color:var(--green)}
.card{background:var(--bg2);border-radius:var(--r);padding:1rem 1.25rem;border:.5px solid var(--border);margin-bottom:1.1rem}
.two{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:1.1rem}
.three{display:grid;grid-template-columns:1.4fr 1fr 1fr;gap:12px;margin-bottom:1.1rem}
table.wt{width:100%;border-collapse:collapse;font-size:12px}
table.wt th{padding:7px 10px;text-align:left;color:var(--text2);font-weight:500;border-bottom:1px solid var(--border2);white-space:nowrap}
table.wt th.r,table.wt td.r{text-align:right}
table.wt td{padding:6px 10px;border-bottom:.5px solid var(--border);color:var(--text)}
table.wt tr:last-child td{border-bottom:none}
table.wt tr:hover td{background:rgba(255,255,255,0.03)}
.badge{display:inline-block;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:500}
.dot{width:8px;height:8px;border-radius:2px;display:inline-block;margin-right:5px;vertical-align:middle}
.leg{display:flex;flex-wrap:wrap;gap:6px 14px;margin-bottom:10px}
.leg-item{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--text2)}
.sub{font-size:11px;color:var(--text2);margin:-8px 0 10px}
</style>
</head>
<body>

<div class="hdr">
  <div>
    <h1>HR1 — Weekly Cost Analysis</h1>
    <p style="font-size:12px;color:var(--text2);margin-top:3px">Packaging Material · June 2026</p>
  </div>
  <div class="hdr-r">
    Last updated: ##GEN_TIME##<br>
    <span style="color:var(--text3)">Auto-refreshes every hour</span>
  </div>
</div>

<div class="kpi-top">
  <div class="kpi-left">
    <div class="kpi accent">
      <label>Total Month Cost · Jun</label>
      <span>##TOTAL_MONTH##</span>##JUN_BADGE##
    </div>
    <div class="kpi" style="border:.5px solid rgba(239,159,39,0.4)">
      <label style="color:#EF9F27">MRR ##MRR_DETAIL##</label>
      <span style="color:#EF9F27">##MRR##</span>
    </div>
  </div>
  <div class="kpi-weeks">
    <div class="kpi">
      <label>Week 1 · 1–7 Jun</label>
      <span>##W1##</span>
    </div>
    <div class="kpi">
      <label>Week 2 · 8–14 Jun</label>
      <span>##W2##</span>
    </div>
    <div class="kpi">
      <label>Week 3 · 15–21 Jun</label>
      <span>##W3##</span>
    </div>
    <div class="kpi">
      <label>Week 4 · 22–30 Jun</label>
      <span>##W4##</span>
    </div>
  </div>
</div>
<div class="kpi-hist">##HIST_CARDS##</div>

<div class="two">
  <div class="card">
    <h2>Weekly total cost trend</h2>
    <div style="position:relative;height:260px"><canvas id="trendChart"></canvas></div>
  </div>
  <div class="card">
    <h2>Cost by category — weekly</h2>
    <div class="leg" id="catLegend"></div>
    <div style="position:relative;height:220px"><canvas id="catStackChart"></canvas></div>
  </div>
</div>

<div class="card">
  <h2>Top 20 materials — total consumption cost (June)</h2>
  <p class="sub">Cost = qty × unit price × (1 + tax)</p>
  <div style="position:relative;height:340px"><canvas id="top20Chart"></canvas></div>
</div>

<div class="card">
  <h2>Weekly cost breakdown by material</h2>
  <p class="sub">Sorted by total cost · only materials with consumption shown</p>
  <div style="overflow-x:auto">
  <table class="wt" id="costTable">
    <thead>
      <tr>
        <th>Material</th>
        <th>Category</th>
        <th class="r">Unit Price</th>
        <th class="r">W1 Qty</th>
        <th class="r">W1 Cost</th>
        <th class="r">W2 Qty</th>
        <th class="r">W2 Cost</th>
        <th class="r">W3 Qty</th>
        <th class="r">W3 Cost</th>
        <th class="r">W4 Qty</th>
        <th class="r">W4 Cost</th>
        <th class="r">Total Cost</th>
      </tr>
    </thead>
    <tbody id="costTbody"></tbody>
  </table>
  </div>
</div>

<p style="font-size:11px;color:var(--text3);text-align:center;margin-top:1.5rem;padding-top:1rem;border-top:.5px solid var(--border)">
  HR1 Packaging Material Monitor &nbsp;·&nbsp; ##GEN_TIME## &nbsp;·&nbsp; Source: Google Sheets
</p>

##CHARTJS##
<script>
window.onerror=function(msg,src,line,col,err){document.body.insertAdjacentHTML('afterbegin','<div style="position:fixed;top:0;left:0;right:0;background:#E24B4A;color:#fff;padding:8px 12px;font-family:monospace;font-size:12px;z-index:9999">JS Error: '+msg+' (line '+line+')</div>');return false;};

##DATA_BLOCK##

const CC={
  "Corrugated Box":"#378ADD","Temper Proof":"#D85A30","Zipper Bag":"#1D9E75",
  "Other":"#888780","Shrink Wrap":"#7F77DD","Thermosheet":"#BA7517","Gunny Bag":"#639922"
};

function fmtINR(v){
  v=Math.round(v);
  if(v===0)return "₹0";
  let s=v.toString(),last3=s.slice(-3),rest=s.slice(0,-3),parts=[];
  while(rest.length>2){parts.unshift(rest.slice(-2));rest=rest.slice(0,-2);}
  if(rest)parts.unshift(rest);
  return "₹"+parts.join(",")+","+last3;
}

// Category legend
const legEl=document.getElementById("catLegend");
if(legEl) CATS_SORTED.forEach(cat=>{
  const s=document.createElement("span");s.className="leg-item";
  s.innerHTML=`<span style="width:8px;height:8px;border-radius:2px;background:${CC[cat]||"#888"};display:inline-block"></span>${cat}`;
  legEl.appendChild(s);
});

// Trend line chart
try{
new Chart(document.getElementById("trendChart"),{
  type:"bar",
  data:{
    labels:WEEK_LABELS,
    datasets:[{
      label:"Cost (₹)",
      data:WEEK_VALS,
      backgroundColor:WEEK_VALS.map((_,i)=>["#378ADD","#1D9E75","#D85A30","#EF9F27"][i]||"#888"),
      borderRadius:6,borderSkipped:false
    }]
  },
  options:{
    responsive:true,maintainAspectRatio:false,
    plugins:{
      legend:{display:false},
      tooltip:{callbacks:{label:c=>" "+fmtINR(c.parsed.y)}}
    },
    scales:{
      x:{grid:{color:"rgba(136,135,128,0.1)"},ticks:{color:"#8a8880"}},
      y:{grid:{color:"rgba(136,135,128,0.1)"},ticks:{color:"#8a8880",callback:v=>fmtINR(v)}}
    }
  }
});}catch(e){console.error("trendChart:",e);}

// Stacked bar by category
try{
const catDatasets=CATS_SORTED.map(cat=>({
  label:cat,
  data:WEEK_KEYS.map(wk=>(CAT_WEEK[cat]||{})[wk]||0),
  backgroundColor:CC[cat]||"#888",
  borderWidth:0,borderRadius:3
}));
new Chart(document.getElementById("catStackChart"),{
  type:"bar",
  data:{labels:WEEK_LABELS,datasets:catDatasets},
  options:{
    responsive:true,maintainAspectRatio:false,
    plugins:{
      legend:{display:false},
      tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${fmtINR(c.parsed.y)}`}}
    },
    scales:{
      x:{stacked:true,grid:{display:false},ticks:{color:"#8a8880"}},
      y:{stacked:true,grid:{color:"rgba(136,135,128,0.1)"},ticks:{color:"#8a8880",callback:v=>fmtINR(v)}}
    }
  }
});}catch(e){console.error("catStackChart:",e);}

// Top 20 horizontal bar
try{
new Chart(document.getElementById("top20Chart"),{
  type:"bar",
  data:{
    labels:TOP20.map(m=>m.name),
    datasets:[{
      data:TOP20.map(m=>m.total_cost),
      backgroundColor:TOP20.map(m=>CC[m.category]||"#888"),
      borderRadius:4,borderSkipped:false
    }]
  },
  options:{
    indexAxis:"y",responsive:true,maintainAspectRatio:false,
    plugins:{
      legend:{display:false},
      tooltip:{callbacks:{label:c=>` ${fmtINR(c.parsed.x)}`}}
    },
    scales:{
      x:{grid:{color:"rgba(136,135,128,0.1)"},ticks:{color:"#8a8880",callback:v=>fmtINR(v)}},
      y:{grid:{display:false},ticks:{color:"#8a8880",font:{size:11}}}
    }
  }
});}catch(e){console.error("top20Chart:",e);}

// Cost table
const tbody=document.getElementById("costTbody");
ALL_MAT.forEach((m,idx)=>{
  const wks=["W1","W2","W3","W4"];
  const tr=document.createElement("tr");
  if(idx%2===1)tr.style.background="rgba(255,255,255,0.02)";
  const catCol=CC[m.category]||"#888";
  let cells=`<td>${m.name}</td>`;
  cells+=`<td><span class="badge" style="background:${catCol}22;color:${catCol}">${m.category}</span></td>`;
  cells+=`<td class="r" style="color:var(--text2)">₹${m.unit_price}</td>`;
  wks.forEach(wk=>{
    const d=m.week_data[wk]||{qty:0,cost:0};
    cells+=`<td class="r">${d.qty||""}</td>`;
    cells+=`<td class="r" style="color:${d.cost>0?"var(--text)":"var(--text3)"}">${d.cost>0?fmtINR(d.cost):""}</td>`;
  });
  cells+=`<td class="r" style="font-weight:500">${fmtINR(m.total_cost)}</td>`;
  tr.innerHTML=cells;
  tbody.appendChild(tr);
});
</script>
</body>
</html>
"""


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    config = json.loads((SCRIPT_DIR / "config.json").read_text(encoding="utf-8"))

    warehouses = {
        "HR1": {
            "spreadsheet_id":  config["spreadsheet_id"],
            "daily_use_sheet": config["daily_use_sheet"],
        },
        "CH1": {
            "spreadsheet_id":  config["ch1_spreadsheet_id"],
            "daily_use_sheet": config.get("ch1_daily_use_sheet", "Daily Use PMS"),
        },
        "MU1": {
            "spreadsheet_id":  config["mu1_spreadsheet_id"],
            "daily_use_sheet": config.get("mu1_daily_use_sheet", "Daily Use PMS"),
        },
    }

    gc = get_gc()
    wh_results = {}
    for wh, wcfg in warehouses.items():
        log.info(f"Fetching {wh}…")
        sh         = gc.open_by_key(wcfg["spreadsheet_id"])
        sum_rows   = find_worksheet(sh, "Summary").get_all_values()
        daily_rows = find_worksheet(sh, wcfg["daily_use_sheet"]).get_all_values()

        prices     = parse_prices(sum_rows)
        historical = parse_historical(sum_rows)
        materials, _ = parse_daily(daily_rows)
        enriched   = compute_weekly(materials, prices)
        week_totals, cat_week, _ = aggregate(enriched)

        log.info(f"  {wh} totals: " + " | ".join(
            f"{label} {fmt_inr(week_totals[wk])}" for wk, label, _ in WEEKS))
        wh_results[wh] = {
            "enriched":    enriched,
            "week_totals": week_totals,
            "cat_week":    cat_week,
            "historical":  historical,
        }

    html = generate_multi_html(wh_results, datetime.now())
    out_path = SCRIPT_DIR / "Weekly_Cost.html"
    out_path.write_text(html, encoding="utf-8")
    log.info(f"Dashboard saved → {out_path}")


if __name__ == "__main__":
    main()
