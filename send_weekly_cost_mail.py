"""
Weekly Cost mailer — HR1 Packaging Material
Generates and sends two versions:
  - Operations : full detail (KPIs + category chart + top 15 + weekly table)
  - Finance    : cost KPIs + 2-row category table (Bubble Roll vs Others)
"""
import json, smtplib, ssl, sys, calendar
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

SCRIPT_DIR = Path(__file__).parent
config     = json.loads((SCRIPT_DIR / "config.json").read_text(encoding="utf-8"))

sys.path.insert(0, str(SCRIPT_DIR))
from weekly_cost import (get_gc, parse_prices, parse_historical, parse_daily,
                          compute_weekly, aggregate, WEEKS, CAT_COLORS, fmt_inr)

# ── fetch data once ───────────────────────────────────────────────────────────
gc = get_gc()
sh = gc.open_by_key(config["spreadsheet_id"])
sum_rows   = sh.worksheet("Summary").get_all_values()
daily_rows = sh.worksheet(config["daily_use_sheet"]).get_all_values()

prices     = parse_prices(sum_rows)
historical = parse_historical(sum_rows)
materials, _ = parse_daily(daily_rows)
enriched   = compute_weekly(materials, prices)
week_totals, cat_week, _ = aggregate(enriched)

week_keys    = [wk for wk, *_ in WEEKS]
week_labels  = [label for _, label, _ in WEEKS]
week_colors  = ["#378ADD", "#1D9E75", "#D85A30", "#EF9F27"]
total_month  = sum(week_totals[wk] for wk in week_keys)
hist_list    = (historical or [])[-4:]

now          = datetime.now()
now_str      = now.strftime("%d %b %Y, %I:%M %p")
days_in_month = calendar.monthrange(now.year, now.month)[1]
days_elapsed  = now.day
mrr           = (total_month / days_elapsed) * days_in_month if days_elapsed > 0 else 0

# ── shared helpers ────────────────────────────────────────────────────────────

def pct_change(new, old): return (new - old) / old * 100 if old else None

def pct_badge(pct):
    if pct is None: return ""
    arrow = "▲" if pct > 0 else "▼"
    color = "#E24B4A" if pct > 0 else "#1D9E75"
    return f'<span style="font-size:11px;color:{color};margin-left:4px">{arrow} {abs(pct):.1f}%</span>'

def kpi_cell(label, val, color, bg, border_color, badge="", width="25%"):
    return f"""<td style="width:{width};padding:4px;vertical-align:top">
      <div style="background:{bg};border-radius:8px;padding:12px 14px;border:1px solid {border_color};border-top:3px solid {color}">
        <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:5px">{label}</div>
        <div style="font-size:18px;font-weight:700;color:{color};display:inline">{val}</div>{badge}
      </div></td>"""

def bar_html(pct, color, height=14):
    w = max(2, int(pct))
    return (f'<div style="background:#f0f0f0;border-radius:3px;height:{height}px;width:100%">'
            f'<div style="background:{color};width:{w}%;height:{height}px;border-radius:3px"></div></div>')

def th(content, align="left"):
    return (f'<th style="padding:8px 10px;text-align:{align};font-weight:600;color:#555;'
            f'font-size:11px;text-transform:uppercase;letter-spacing:0.04em;'
            f'border-bottom:2px solid #ddd;white-space:nowrap">{content}</th>')

def td_c(content, align="left", bold=False, color="#1a1a1a", bg=""):
    bg_s = f"background:{bg};" if bg else ""
    fw   = "font-weight:600;" if bold else ""
    return f'<td style="padding:7px 10px;text-align:{align};{fw}color:{color};{bg_s}border-bottom:1px solid #ebebeb">{content}</td>'

def send_email(subject, html_body, recipients):
    smtp_cfg = config["smtp"]
    email_cfg = config["email"]
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = f"{email_cfg['from_name']} <{email_cfg['from']}>"
    msg["To"]      = ", ".join(recipients)
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"], context=context) as server:
        server.login(smtp_cfg["username"], smtp_cfg["password"])
        server.sendmail(email_cfg["from"], recipients, msg.as_string())
    print(f"  Sent to: {', '.join(recipients)}")

# ── shared KPI blocks ─────────────────────────────────────────────────────────

last_hist_val  = hist_list[-1][1] if hist_list else 0
jun_badge      = pct_badge(pct_change(mrr, last_hist_val) if last_hist_val else None)

kpi_row1 = (
    kpi_cell("Total Month · Jun", fmt_inr(total_month), "#1D9E75", "#f8f8f8", "#e8e8e8", jun_badge, "50%") +
    kpi_cell(f"MRR ({days_elapsed} of {days_in_month} days)", fmt_inr(mrr), "#EF9F27", "#fffbf0", "#ffe58f", "", "50%")
)

kpi_row2 = "".join(
    kpi_cell(f"Week {i+1} · {week_labels[i]}", fmt_inr(week_totals[wk]), week_colors[i], "#f8f8f8", "#e8e8e8", "", "25%")
    for i, wk in enumerate(week_keys)
)

hist_cells = ""
for i, (label, val) in enumerate(hist_list):
    prev_val = hist_list[i-1][1] if i > 0 else None
    badge    = pct_badge(pct_change(val, prev_val) if prev_val else None)
    hist_cells += kpi_cell(f"{label} Final", fmt_inr(val), "#378ADD", "#f0f6ff", "#cce0ff", badge, "25%")

kpi_block = f"""
  <tr><td style="background:#fff;padding:20px 24px 4px 24px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>{kpi_row1}</tr></table>
  </td></tr>
  <tr><td style="background:#fff;padding:4px 24px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>{kpi_row2}</tr></table>
  </td></tr>
  <tr><td style="background:#fff;padding:4px 24px 16px 24px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>{hist_cells}</tr></table>
  </td></tr>"""

def divider():
    return '<tr><td style="background:#fff;padding:0 24px"><hr style="border:none;border-top:1px solid #ebebeb;margin:0"></td></tr>'

def email_wrapper(title, subtitle, body_rows, generated_at_str):
    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:24px 0">
<tr><td align="center">
<table width="720" cellpadding="0" cellspacing="0" style="max-width:720px;width:100%">

  <tr><td style="background:#1a1a1a;border-radius:10px 10px 0 0;padding:20px 24px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td>
        <div style="font-size:20px;font-weight:700;color:#fff">{title}</div>
        <div style="font-size:12px;color:#aaa;margin-top:3px">{subtitle}</div>
      </td>
      <td align="right">
        <div style="font-size:11px;color:#888">Generated: {generated_at_str}</div>
        <div style="font-size:11px;color:#555;margin-top:2px">Source: Google Sheets (live)</div>
      </td>
    </tr></table>
  </td></tr>

  {body_rows}

  <tr><td style="background:#fff;background:#fffbe6;border:1px solid #ffe58f;padding:10px 18px">
    <div style="font-size:12px;color:#7d6608"><strong>Note:</strong> This is a sample mail format — enhancements are in process for the final version.</div>
  </td></tr>

  <tr><td style="padding:14px 0;text-align:center;border-radius:0 0 10px 10px">
    <div style="font-size:11px;color:#aaa">HR1 Packaging Material Monitor &nbsp;·&nbsp; {generated_at_str} &nbsp;·&nbsp; Auto-generated report</div>
  </td></tr>

</table>
</td></tr>
</table></body></html>"""

# ── OPERATIONS EMAIL ──────────────────────────────────────────────────────────

# Category bars
cat_totals  = {cat: sum(cat_week[cat].values()) for cat in cat_week}
cats_sorted = sorted(cat_totals, key=lambda c: -cat_totals[c])
max_cat     = max(cat_totals.values()) if cat_totals else 1

cat_rows_html = ""
for cat in cats_sorted:
    total = cat_totals.get(cat, 0)
    if total == 0: continue
    color = CAT_COLORS.get(cat, "#888")
    pct   = total / max_cat * 100
    cat_rows_html += f"""<tr>
      <td style="padding:6px 10px 6px 0;width:160px;font-size:12px;color:#333;white-space:nowrap">{cat}</td>
      <td style="padding:6px 0;width:100%">{bar_html(pct, color, 16)}</td>
      <td style="padding:6px 0 6px 10px;font-size:12px;font-weight:600;color:#333;text-align:right;white-space:nowrap">{fmt_inr(total)}</td>
    </tr>"""

# Top 15
top15 = sorted([m for m in enriched if m["total_cost"] > 0], key=lambda x: -x["total_cost"])[:15]
max_cost = top15[0]["total_cost"] if top15 else 1
top_rows_html = ""
for i, m in enumerate(top15):
    bg    = "#fafafa" if i % 2 == 0 else "#fff"
    color = CAT_COLORS.get(m["category"], "#888")
    pct   = m["total_cost"] / max_cost * 100
    top_rows_html += f"""<tr style="background:{bg}">
      {td_c(f'<span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:{color};margin-right:6px;vertical-align:middle"></span>{m["name"]}')}
      {td_c(m["category"], color="#666")}
      {td_c(bar_html(pct, color, 12))}
      {td_c(fmt_inr(m["total_cost"]), "right", bold=True)}
    </tr>"""

# Weekly detail
wk_header = "".join(th(f"W{i+1} Qty", "right") + th(f"W{i+1} Cost", "right") for i in range(4))
detail_header = f"<tr>{th('Material')}{th('Category')}{wk_header}{th('Total Cost', 'right')}</tr>"
detail_rows_html = ""
for i, m in enumerate(sorted([m for m in enriched if m["total_cost"] > 0], key=lambda x: -x["total_cost"])):
    bg    = "#fafafa" if i % 2 == 0 else "#fff"
    color = CAT_COLORS.get(m["category"], "#888")
    wk_cells = ""
    for wk in week_keys:
        d = m["week_data"].get(wk, {"qty": 0, "cost": 0})
        wk_cells += (td_c(f'{int(d["qty"]):,}' if d["qty"] else "", "right", color="#555") +
                     td_c(fmt_inr(d["cost"]) if d["cost"] > 0 else "", "right", color="#333" if d["cost"] else "#ccc"))
    detail_rows_html += f"""<tr style="background:{bg}">
      {td_c(f'<span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:{color};margin-right:5px;vertical-align:middle"></span>{m["name"]}')}
      {td_c(m["category"], color="#666")}
      {wk_cells}
      {td_c(fmt_inr(m["total_cost"]), "right", bold=True, color="#1D9E75")}
    </tr>"""

ops_body = kpi_block + divider() + f"""
  <tr><td style="background:#fff;padding:20px 24px">
    <div style="font-size:13px;font-weight:600;color:#1a1a1a;margin-bottom:14px">Cost by Category (June total)</div>
    <table width="100%" cellpadding="0" cellspacing="0">{cat_rows_html}</table>
  </td></tr>""" + divider() + f"""
  <tr><td style="background:#fff;padding:20px 24px">
    <div style="font-size:13px;font-weight:600;color:#1a1a1a;margin-bottom:4px">Top 15 Materials by Consumption Cost</div>
    <div style="font-size:11px;color:#999;margin-bottom:14px">Cost = Consumed Stock Value from Summary (ex-tax)</div>
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>{th("Material")}{th("Category")}{th("Relative Cost")}{th("Total Cost","right")}</tr>
      {top_rows_html}
    </table>
  </td></tr>""" + divider() + f"""
  <tr><td style="background:#fff;padding:20px 24px">
    <div style="font-size:13px;font-weight:600;color:#1a1a1a;margin-bottom:14px">Weekly Breakdown by Material</div>
    <div style="overflow-x:auto">
    <table width="100%" cellpadding="0" cellspacing="0" style="min-width:640px">
      {detail_header}{detail_rows_html}
    </table></div>
  </td></tr>"""

ops_html = email_wrapper(
    "HR1 — Weekly Cost Analysis · Operations",
    "Packaging Material · June 2026",
    ops_body, now_str
)

# ── FINANCE EMAIL ─────────────────────────────────────────────────────────────

# Finance category table — same categories as ops (Bubble Roll now its own category)
fin_cat_totals  = {cat: sum(cat_week[cat].values()) for cat in cat_week}
fin_cats_sorted = sorted(fin_cat_totals, key=lambda c: -fin_cat_totals[c])
fin_max_cat     = max(fin_cat_totals.values()) if fin_cat_totals else 1

# collect Other materials for footnote
other_items = sorted(
    [m["name"] for m in enriched if m["category"] == "Other" and m["total_cost"] > 0]
)
other_footnote = (
    f'<div style="font-size:11px;color:#888;margin-top:10px;line-height:1.7">'
    f'<strong style="color:#555">* Others includes:</strong> '
    + ", ".join(other_items) +
    f'</div>'
) if other_items else ""

fin_cat_rows = ""
for cat in fin_cats_sorted:
    val = fin_cat_totals.get(cat, 0)
    if val == 0: continue
    color      = CAT_COLORS.get(cat, "#888780")
    pct        = val / fin_max_cat * 100
    share      = val / total_month * 100 if total_month else 0
    label      = f"{cat}*" if cat == "Other" else cat
    fin_cat_rows += f"""<tr>
      {td_c(f'<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:{color};margin-right:8px;vertical-align:middle"></span>{label}', bold=True)}
      {td_c(bar_html(pct, color, 16))}
      {td_c(fmt_inr(val), "right", bold=True)}
      {td_c(f'{share:.1f}%', "right", color="#888")}
    </tr>"""
fin_cat_rows += f"""<tr style="background:#f8f8f8">
  {td_c('<strong>Total</strong>')}
  {td_c('')}
  {td_c(f'<strong>{fmt_inr(total_month)}</strong>', "right")}
  {td_c('<strong>100%</strong>', "right", color="#888")}
</tr>"""

fin_body = kpi_block + divider() + f"""
  <tr><td style="background:#fff;padding:20px 24px">
    <div style="font-size:13px;font-weight:600;color:#1a1a1a;margin-bottom:4px">Cost by Category (June total)</div>
    <div style="font-size:11px;color:#999;margin-bottom:14px">Consumed Stock Value ex-tax · HR1 Warehouse</div>
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>{th("Category")}{th("Relative Spend")}{th("Cost","right")}{th("Share","right")}</tr>
      {fin_cat_rows}
    </table>
    {other_footnote}
  </td></tr>"""

fin_html = email_wrapper(
    "HR1 — Weekly Cost Analysis · Finance",
    "Packaging Material · June 2026",
    fin_body, now_str
)

# ── send both ─────────────────────────────────────────────────────────────────
email_cfg = config["email"]
subject   = f"HR1 Packaging Material — Weekly Cost Report ({now.strftime('%d %b %Y')})"

print("Sending Operations email…")
send_email(f"[Operations] {subject}", ops_html, email_cfg["ops_to"])

print("Sending Finance email…")
send_email(f"[Finance] {subject}", fin_html, email_cfg["finance_to"])
