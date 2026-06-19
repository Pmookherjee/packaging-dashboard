"""
HR1 Packaging Material — Weekly Cost Dashboard
Flask web app for Render deployment.
Serves the dashboard HTML, regenerating from Google Sheets at most once per hour.
"""
import json, os, time, threading, logging
from pathlib import Path
from datetime import datetime
from flask import Flask, Response, redirect, url_for
from functools import wraps

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

app = Flask(__name__)
SCRIPT_DIR = Path(__file__).parent

# ── credential bootstrap ───────────────────────────────────────────────────────
# On Render: set GOOGLE_CREDENTIALS and GOOGLE_TOKEN env vars.
# Locally: credentials.json and token.json files are used directly.

def bootstrap_credentials():
    creds_env = os.environ.get("GOOGLE_CREDENTIALS")
    token_env  = os.environ.get("GOOGLE_TOKEN")
    if creds_env:
        (SCRIPT_DIR / "credentials.json").write_text(creds_env, encoding="utf-8")
        log.info("credentials.json written from env var")
    if token_env:
        (SCRIPT_DIR / "token.json").write_text(token_env, encoding="utf-8")
        log.info("token.json written from env var")

bootstrap_credentials()

# ── config ─────────────────────────────────────────────────────────────────────
_cfg_env = {
    "spreadsheet_id":  os.environ.get("SPREADSHEET_ID"),
    "daily_use_sheet": os.environ.get("DAILY_USE_SHEET", "Daily Use PMS "),
}

def get_config():
    cfg_path = SCRIPT_DIR / "config.json"
    if cfg_path.exists():
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    else:
        cfg = {}
    # env vars override file values
    if _cfg_env["spreadsheet_id"]:
        cfg["spreadsheet_id"] = _cfg_env["spreadsheet_id"]
    if _cfg_env["daily_use_sheet"]:
        cfg["daily_use_sheet"] = _cfg_env["daily_use_sheet"]
    return cfg

# ── cache ──────────────────────────────────────────────────────────────────────
CACHE_TTL = 3600  # seconds

_cache = {"html": None, "generated_at": 0, "lock": threading.Lock()}


def build_dashboard():
    import sys
    sys.path.insert(0, str(SCRIPT_DIR))
    from weekly_cost import (get_gc, parse_prices, parse_historical, parse_daily,
                              compute_weekly, aggregate, generate_html)

    config = get_config()
    gc = get_gc()
    sh = gc.open_by_key(config["spreadsheet_id"])

    log.info("Fetching Summary…")
    sum_rows = sh.worksheet("Summary").get_all_values()
    log.info("Fetching Daily Use PMS…")
    daily_rows = sh.worksheet(config["daily_use_sheet"]).get_all_values()

    prices     = parse_prices(sum_rows)
    historical = parse_historical(sum_rows)
    materials, _ = parse_daily(daily_rows)
    enriched   = compute_weekly(materials, prices)
    week_totals, cat_week, _ = aggregate(enriched)

    html = generate_html(enriched, week_totals, cat_week, datetime.now(), historical)
    log.info("Dashboard built successfully")
    return html


def get_cached_html(force=False):
    with _cache["lock"]:
        age = time.time() - _cache["generated_at"]
        if force or _cache["html"] is None or age > CACHE_TTL:
            log.info(f"{'Force-' if force else ''}regenerating dashboard (age={age:.0f}s)…")
            try:
                _cache["html"] = build_dashboard()
                _cache["generated_at"] = time.time()
            except Exception as e:
                log.error(f"Build failed: {e}")
                if _cache["html"] is None:
                    raise
                # serve stale on error
        return _cache["html"]


# ── optional basic auth ────────────────────────────────────────────────────────
DASH_USER = os.environ.get("DASH_USER", "")
DASH_PASS = os.environ.get("DASH_PASS", "")

def check_auth(username, password):
    if not DASH_USER:
        return True  # no auth configured
    return username == DASH_USER and password == DASH_PASS

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not DASH_USER:
            return f(*args, **kwargs)
        auth = app.current_app.request.authorization if hasattr(app, 'current_app') else None
        from flask import request
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return Response(
                "Authentication required.",
                401,
                {"WWW-Authenticate": 'Basic realm="HR1 Dashboard"'}
            )
        return f(*args, **kwargs)
    return decorated

# ── routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
@requires_auth
def index():
    html = get_cached_html()
    return Response(html, mimetype="text/html")


@app.route("/refresh")
@requires_auth
def refresh():
    get_cached_html(force=True)
    return redirect(url_for("index"))


@app.route("/health")
def health():
    age = int(time.time() - _cache["generated_at"])
    return {"status": "ok", "cache_age_seconds": age, "has_data": _cache["html"] is not None}


# ── startup ────────────────────────────────────────────────────────────────────
def warm_cache():
    """Pre-build on startup so first visitor doesn't wait."""
    try:
        get_cached_html()
    except Exception as e:
        log.warning(f"Warm-up failed (will retry on first request): {e}")

# Warm-up disabled on PythonAnywhere (30s worker timeout too short for Sheets fetch)
# threading.Thread(target=warm_cache, daemon=True).start()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
