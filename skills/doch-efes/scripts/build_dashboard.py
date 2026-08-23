#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_dashboard.py — בונה דשבורד HTML יחיד, standalone, RTL.

העיצוב יורש את מערכת העיצוב של הסקיל ``file-dashboard`` (פלטת Midnight, כרטיסי
KPI, טאבים) כדי שהדשבורד הזה ייראה כמו שאר הדשבורדים במשרד ולא כאי בפני עצמו.

הקובץ עומד בפני עצמו: אין תלות ב-CDN, אין fetch, כל הגרפים הם SVG/CSS שנבנים
בזמן היצירה. הסיבה מעשית — הדשבורד נשלח ליזם במייל ונפתח מהדיסק, ולעיתים
קרובות בלי רשת. הגופן נטען מ-Google Fonts עם נפילה לגופן מערכת.

שימוש:
    python3 build_dashboard.py project.json -o "דשבורד.html"
"""

from __future__ import annotations

import argparse
import html
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from model import DISCLAIMER, full_output, load_project  # noqa: E402


def esc(x):
    return html.escape(str(x if x is not None else ""))


def money(x, short=True):
    try:
        v = float(x)
    except (TypeError, ValueError):
        return "—"
    if short and abs(v) >= 1_000_000:
        return "₪%.1fM" % (v / 1_000_000)
    if short and abs(v) >= 1_000:
        return "₪%.0fK" % (v / 1_000)
    return "₪%s" % format(int(round(v)), ",")


def full_money(x):
    try:
        return "₪%s" % format(int(round(float(x))), ",")
    except (TypeError, ValueError):
        return "—"


def percent(x, digits=1):
    try:
        return ("%." + str(digits) + "f%%") % (float(x) * 100)
    except (TypeError, ValueError):
        return "—"


# --- גרפים -----------------------------------------------------------------
# כולם SVG שנבנה כמחרוזת. אין ספריית גרפים כי הדשבורד חייב להיות קובץ אחד
# שנפתח בלי רשת, וארבעה גרפים סטטיים לא מצדיקים הטמעת ספרייה שלמה בבסיס 64.

PALETTE = ["#3b82f6", "#06b6d4", "#f59e0b", "#10b981", "#8b5cf6", "#ef4444", "#84cc16"]


def donut(items, size=230, thickness=34):
    """טבעת מבנה עלויות. items = [(label, value)]"""
    total = sum(v for _, v in items) or 1
    r = (size - thickness) / 2
    cx = cy = size / 2
    circ = 2 * 3.14159265 * r
    offset = 0.0
    parts = []
    legend = []
    for i, (label, value) in enumerate(items):
        frac = value / total
        color = PALETTE[i % len(PALETTE)]
        parts.append(
            '<circle cx="%.1f" cy="%.1f" r="%.1f" fill="none" stroke="%s" '
            'stroke-width="%d" stroke-dasharray="%.3f %.3f" '
            'stroke-dashoffset="%.3f" transform="rotate(-90 %.1f %.1f)"><title>%s — %s (%s)</title></circle>'
            % (cx, cy, r, color, thickness, frac * circ, circ, -offset * circ, cx, cy,
               esc(label), full_money(value), percent(frac)))
        legend.append(
            '<div class="lg"><span class="dot" style="background:%s"></span>'
            '<span class="lg-l">%s</span><span class="lg-v">%s</span>'
            '<span class="lg-p">%s</span></div>'
            % (color, esc(label), money(value), percent(frac)))
        offset += frac
    svg = ('<svg viewBox="0 0 %d %d" width="%d" height="%d" role="img">%s'
           '<text x="%.1f" y="%.1f" text-anchor="middle" fill="#f0f4ff" '
           'font-size="19" font-weight="700">%s</text>'
           '<text x="%.1f" y="%.1f" text-anchor="middle" fill="#8b9fc0" font-size="12">'
           'סך עלויות</text></svg>'
           % (size, size, size, size, "".join(parts), cx, cy + 2, money(total), cx, cy + 20))
    return '<div class="donut-wrap">%s<div class="legend">%s</div></div>' % (svg, "".join(legend))


def cashflow_chart(rows, width=980, height=300):
    """עמודות הוצאות/הכנסות + קו יתרת אשראי על אותו ציר זמן."""
    if not rows:
        return ""
    pad_l, pad_b, pad_t = 58, 42, 16
    plot_w, plot_h = width - pad_l - 16, height - pad_b - pad_t
    peak = max(max(r["outflow"], r["revenue"], r["debt_balance"]) for r in rows) or 1
    n = len(rows)
    slot = plot_w / n
    bw = min(slot * 0.32, 26)

    def y(v):
        return pad_t + plot_h - (v / peak) * plot_h

    bars, labels, pts = [], [], []
    for i, row in enumerate(rows):
        # RTL: הרבעון הראשון מימין
        cx = 16 + plot_w - (i + 0.5) * slot
        bars.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="#ef4444" '
                    'opacity=".85" rx="2"><title>%s הוצאות %s</title></rect>'
                    % (cx - bw - 1, y(row["outflow"]), bw,
                       plot_h - (y(row["outflow"]) - pad_t),
                       esc(row["quarter"]), full_money(row["outflow"])))
        bars.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="#10b981" '
                    'opacity=".85" rx="2"><title>%s הכנסות %s</title></rect>'
                    % (cx + 1, y(row["revenue"]), bw, plot_h - (y(row["revenue"]) - pad_t),
                       esc(row["quarter"]), full_money(row["revenue"])))
        pts.append("%.1f,%.1f" % (cx, y(row["debt_balance"])))
        if n <= 16 or i % 2 == 0:
            labels.append('<text x="%.1f" y="%d" text-anchor="middle" fill="#8b9fc0" '
                          'font-size="10">%s</text>' % (cx, height - 22, esc(row["quarter"])))
    grid = []
    for f in (0, 0.25, 0.5, 0.75, 1.0):
        gy = pad_t + plot_h - f * plot_h
        grid.append('<line x1="%d" y1="%.1f" x2="%d" y2="%.1f" stroke="#1f2d45"/>'
                    % (pad_l - 46, gy, width - 16, gy))
        grid.append('<text x="%d" y="%.1f" text-anchor="end" fill="#8b9fc0" font-size="10">%s</text>'
                    % (pad_l - 50, gy + 3, money(peak * f)))
    return ('<svg viewBox="0 0 %d %d" width="100%%" height="%d" preserveAspectRatio="xMidYMid meet">'
            '%s%s<polyline points="%s" fill="none" stroke="#f59e0b" stroke-width="2.5"/>%s</svg>'
            '<div class="chart-key"><span><i style="background:#ef4444"></i>הוצאות</span>'
            '<span><i style="background:#10b981"></i>הכנסות</span>'
            '<span><i style="background:#f59e0b"></i>יתרת אשראי</span></div>'
            % (width, height, height, "".join(grid), "".join(bars), " ".join(pts), "".join(labels)))


def heatmap(sens, threshold):
    """מטריצת רגישות מחיר × עלות, מקודדת בצבע סביב סף הרווח."""
    head = "".join('<th>%+d%%</th>' % p for p in sens["price_steps"])
    body = []
    for i, cs in enumerate(sens["cost_steps"]):
        cells = []
        for v in sens["grid"][i]:
            if v >= threshold:
                cls = "ok"
            elif v >= threshold - 0.03:
                cls = "warn"
            elif v >= 0:
                cls = "bad"
            else:
                cls = "neg"
            cells.append('<td class="%s">%s</td>' % (cls, percent(v)))
        body.append("<tr><th>%+d%%</th>%s</tr>" % (cs, "".join(cells)))
    return ('<table class="heat"><thead><tr><th class="corner">עלות בנייה ↓ / מחיר מכירה →</th>'
            '%s</tr></thead><tbody>%s</tbody></table>' % (head, "".join(body)))


def bar_table(rows, headers, aligns=None):
    th = "".join("<th>%s</th>" % esc(h) for h in headers)
    trs = []
    for row in rows:
        tds = "".join("<td>%s</td>" % (c if isinstance(c, str) and c.startswith("<") else esc(c))
                      for c in row)
        trs.append("<tr>%s</tr>" % tds)
    return '<table class="tbl"><thead><tr>%s</tr></thead><tbody>%s</tbody></table>' % (th, "".join(trs))


CSS = """
:root{--bg:#0a0e1a;--surface:#111827;--surface2:#1a2235;--border:#1f2d45;
--text:#f0f4ff;--text2:#8b9fc0;--a1:#3b82f6;--a2:#06b6d4;--a3:#f59e0b;
--a4:#10b981;--a5:#ef4444;--a6:#8b5cf6;}
*{margin:0;padding:0;box-sizing:border-box}
html{direction:rtl}
body{background:var(--bg);color:var(--text);font-family:'Heebo',
 'Segoe UI','Arial Hebrew',Arial,sans-serif;line-height:1.6;min-height:100vh;
 background-image:radial-gradient(circle at 15% 10%,rgba(59,130,246,.08),transparent 40%),
 radial-gradient(circle at 85% 90%,rgba(6,182,212,.06),transparent 40%);}
.wrap{max-width:1200px;margin:0 auto;padding:28px 20px 60px}
header h1{font-size:29px;font-weight:900;letter-spacing:-.5px}
header .sub{color:var(--text2);font-size:15px;margin-top:4px}
header .meta{color:var(--text2);font-size:13px;margin-top:10px}
.verdict{margin:22px 0;padding:16px 20px;border-radius:12px;border:1px solid var(--border);
 background:var(--surface);font-size:17px;font-weight:600;border-right:5px solid var(--a4)}
.verdict.fail{border-right-color:var(--a5)}
.verdict.warn{border-right-color:var(--a3)}
.verdict small{display:block;font-weight:400;font-size:14px;color:var(--text2);margin-top:6px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin:20px 0}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px 20px}
.kpi .v{font-size:29px;font-weight:800;letter-spacing:-.5px}
.kpi .l{color:var(--text2);font-size:13px;margin-top:2px}
.kpi .n{color:var(--text2);font-size:11px;margin-top:6px}
.kpi.good .v{color:var(--a4)} .kpi.bad .v{color:var(--a5)} .kpi.info .v{color:var(--a2)}
.tabs{display:flex;gap:6px;flex-wrap:wrap;margin:26px 0 0;border-bottom:1px solid var(--border)}
.tab{background:none;border:0;color:var(--text2);font:inherit;font-size:15px;
 padding:10px 16px;cursor:pointer;border-bottom:2px solid transparent}
.tab:hover{color:var(--text)}
.tab.on{color:var(--text);border-bottom-color:var(--a1);font-weight:600}
.panel{display:none;animation:fade .35s ease both}.panel.on{display:block}
@keyframes fade{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;
 padding:20px;margin:18px 0;overflow-x:auto}
.card h2{font-size:18px;margin-bottom:14px;font-weight:700}
.card p.note{color:var(--text2);font-size:13px;margin-top:10px}
.tbl{width:100%;border-collapse:collapse;font-size:14px;min-width:520px}
.tbl th{background:var(--surface2);color:var(--text2);font-weight:600;text-align:right;
 padding:10px 12px;border-bottom:1px solid var(--border);white-space:nowrap}
.tbl td{padding:9px 12px;border-bottom:1px solid var(--border)}
.tbl tbody tr:last-child td{border-bottom:0}
.tbl tbody tr:hover{background:var(--surface2)}
.tbl tr.total td{font-weight:700;background:var(--surface2)}
.donut-wrap{display:flex;gap:28px;align-items:center;flex-wrap:wrap}
.legend{flex:1;min-width:260px}
.lg{display:flex;align-items:center;gap:10px;padding:6px 0;font-size:14px;
 border-bottom:1px solid var(--border)}
.lg:last-child{border-bottom:0}
.dot{width:11px;height:11px;border-radius:3px;flex:0 0 auto}
.lg-l{flex:1}.lg-v{color:var(--text2);font-variant-numeric:tabular-nums}
.lg-p{width:56px;text-align:left;color:var(--text2);font-variant-numeric:tabular-nums}
.chart-key{display:flex;gap:18px;justify-content:center;margin-top:8px;
 color:var(--text2);font-size:13px}
.chart-key i{display:inline-block;width:11px;height:11px;border-radius:3px;margin-left:6px}
.heat{border-collapse:separate;border-spacing:3px;font-size:14px;margin:0 auto}
.heat th{color:var(--text2);font-weight:600;padding:6px 10px;font-size:13px}
.heat th.corner{font-size:12px;text-align:right}
.heat td{padding:11px 14px;text-align:center;border-radius:6px;font-weight:600;
 font-variant-numeric:tabular-nums}
.heat td.ok{background:rgba(16,185,129,.22);color:#6ee7b7}
.heat td.warn{background:rgba(245,158,11,.22);color:#fcd34d}
.heat td.bad{background:rgba(239,68,68,.18);color:#fca5a5}
.heat td.neg{background:rgba(239,68,68,.4);color:#fff}
.flag{display:flex;gap:14px;padding:14px 16px;border-radius:11px;margin-bottom:11px;
 border:1px solid var(--border);background:var(--surface2)}
.flag.red{border-right:4px solid var(--a5)}
.flag.amber{border-right:4px solid var(--a3)}
.flag .ttl{font-weight:700;margin-bottom:3px}
.flag .txt{color:var(--text2);font-size:14px}
.badge{flex:0 0 auto;font-size:12px;padding:3px 10px;border-radius:20px;height:fit-content}
.badge.red{background:rgba(239,68,68,.2);color:#fca5a5}
.badge.amber{background:rgba(245,158,11,.2);color:#fcd34d}
footer{margin-top:34px;padding-top:18px;border-top:1px solid var(--border);
 color:var(--text2);font-size:12.5px;line-height:1.7}
@media print{body{background:#fff;color:#111}.tabs{display:none}.panel{display:block!important}}
"""

JS = """
document.querySelectorAll('.tab').forEach(function(t){
  t.addEventListener('click', function(){
    document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('on')});
    document.querySelectorAll('.panel').forEach(function(x){x.classList.remove('on')});
    t.classList.add('on');
    document.getElementById(t.dataset.p).classList.add('on');
  });
});
"""


def build_html(data):
    m = data["model"]
    r = m["results"]
    meta = m["meta"]
    ok = r["meets_threshold"]
    reds = [f for f in data["flags"] if f.get("level") == "red"]

    # --- כותרת ומסקנה ---
    if ok and not reds:
        vcls, vtext = "", "הפרויקט עומד בסף הכדאיות"
    elif ok:
        vcls, vtext = "warn", "הפרויקט עומד בסף — אך אותרו דגלים אדומים"
    else:
        vcls, vtext = "fail", "הפרויקט אינו עומד בסף הכדאיות"
    vsub = ("רווח יזמי %s, המהווה %s מהעלויות. הסף שהוגדר: %s."
            % (full_money(r["profit_before_tax"]), percent(r["margin_on_cost"]),
               percent(r["profit_threshold"], 0)))

    kpis = [
        ("רווח יזמי", full_money(r["profit_before_tax"]),
         "%s מהעלויות · %s מההכנסות" % (percent(r["margin_on_cost"]),
                                        percent(r["margin_on_revenue"])),
         "good" if ok else "bad"),
        ("IRR שנתי", percent(r["irr_annual"]) if r["irr_annual"] is not None else "—",
         "על ההון העצמי", "info"),
        ("סך עלויות", full_money(r["total_cost"]),
         "%s למ\"ר מכור" % full_money(r["cost_per_sold_sqm"]), ""),
        ("סך הכנסות", full_money(r["total_revenue_net"]),
         "נטו ממע\"מ · ברוטו %s" % money(r["total_revenue_gross"]), ""),
    ]
    kpi_html = "".join(
        '<div class="kpi %s"><div class="v">%s</div><div class="l">%s</div>'
        '<div class="n">%s</div></div>' % (cls, esc(v), esc(label), esc(note))
        for label, v, note, cls in kpis)

    # --- טאב: מבנה עלויות ---
    cost_items = [
        ("קרקע ורכישה", m["land"]["total"]),
        ("בנייה ישירה", m["direct"]["total"]),
        ("עלויות עקיפות", m["indirect"]["total"]),
        ("בצ\"מ", m["contingency"]["amount"]),
        ("מימון", m["finance"]["total"]),
    ]
    cost_rows = [[l["label"], full_money(l["amount"])] for l in m["land"]["lines"]]
    cost_rows += [[l["label"], full_money(l["amount"])] for l in m["direct"]["lines"]]
    cost_rows += [[l["label"], full_money(l["amount"])] for l in m["indirect"]["lines"]]
    cost_rows += [["בצ\"מ (%s)" % percent(m["contingency"]["pct"]),
                   full_money(m["contingency"]["amount"])]]
    cost_rows += [[l["label"], full_money(l["amount"])] for l in m["finance"]["lines"]]

    # --- טאב: הכנסות ---
    rev_rows = [[l["label"], format(int(l["units"]), ","), "%s מ\"ר" % format(round(l["avg_sqm"]), ","),
                 full_money(l["price_per_sqm"]), full_money(l["gross"]), full_money(l["net"])]
                for l in m["revenue"]["lines"]]

    # --- טאב: תזרים ---
    cf_rows = [[row["quarter"], full_money(row["outflow"]), full_money(row["revenue"]),
                full_money(row["equity_used"]), full_money(row["interest"]),
                full_money(row["debt_balance"])] for row in m["cashflow"]]

    # --- טאב: דגלים ---
    if data["flags"]:
        flags_html = "".join(
            '<div class="flag %s"><span class="badge %s">%s</span><div>'
            '<div class="ttl">%s</div><div class="txt">%s</div></div></div>'
            % (f.get("level", "amber"), f.get("level", "amber"),
               "אדום" if f.get("level") == "red" else "צהוב",
               esc(f.get("title", "")), esc(f.get("text", "")))
            for f in data["flags"])
    else:
        flags_html = ('<p class="note">המודל לא זיהה חריגה מהספים שהוגדרו. אין בכך אישור '
                      'לפרויקט — הדגלים נגזרים מהמספרים בלבד.</p>')
    if data["deviations"]:
        flags_html += ('<h2 style="margin-top:22px">סטיות מאסמכתאות חיצוניות</h2>' +
                       "".join('<div class="flag amber"><span class="badge amber">סטייה</span>'
                               '<div><div class="ttl">%s</div><div class="txt">%s</div></div></div>'
                               % (esc(d["label"]), esc(d["text"])) for d in data["deviations"]))

    rate_rows = [["%+.1f נק' אחוז" % s["delta_pp"], percent(s["rate"]),
                  full_money(s["interest"]), full_money(s["profit"]), percent(s["metric"])]
                 for s in data["rate_scenarios"]]

    ident = " · ".join("%s %s" % (k, meta[v]) for k, v in
                       [("גוש", "gush"), ("חלקה", "helka"), ("מגרש", "migrash")]
                       if meta.get(v))

    tabs = [
        ("t1", "סקירה", True),
        ("t2", "עלויות", False),
        ("t3", "הכנסות", False),
        ("t4", "תזרים ומימון", False),
        ("t5", "רגישויות", False),
        ("t6", "דגלים", False),
    ]
    tabs_html = "".join('<button class="tab%s" data-p="%s">%s</button>'
                        % (" on" if on else "", tid, esc(name)) for tid, name, on in tabs)

    panels = []
    panels.append(
        '<div class="panel on" id="t1"><div class="card"><h2>מבנה העלויות</h2>%s</div>'
        '<div class="card"><h2>תזרים רבעוני</h2>%s</div></div>'
        % (donut(cost_items), cashflow_chart(m["cashflow"])))
    panels.append(
        '<div class="panel" id="t2"><div class="card"><h2>פירוט העלויות</h2>%s'
        '<p class="note">סך עלויות הפרויקט: %s</p></div></div>'
        % (bar_table(cost_rows, ["סעיף", "סכום"]), full_money(r["total_cost"])))
    panels.append(
        '<div class="panel" id="t3"><div class="card"><h2>הכנסות צפויות</h2>%s'
        '<p class="note">מחירי המכירה הוזנו %s מע\"מ. הרווח נמדד על ההכנסה נטו, '
        'שכן המע\"מ מועבר לרשות המסים.</p></div></div>'
        % (bar_table(rev_rows, ["רכיב", "יחידות", "שטח ממוצע", "מחיר למ\"ר", "ברוטו", "נטו"]),
           "כוללי" if m["revenue"]["prices_include_vat"] else "ללא"))
    panels.append(
        '<div class="panel" id="t4"><div class="card"><h2>תזרים מזומנים רבעוני</h2>%s</div>'
        '<div class="card"><h2>מבנה המימון</h2>%s</div></div>'
        % (bar_table(cf_rows, ["רבעון", "הוצאות", "הכנסות", "ניצול הון", "ריבית", "יתרת אשראי"]),
           bar_table([["הון עצמי", full_money(m["finance"]["equity"])],
                      ["שיא ניצול אשראי", full_money(m["finance"]["peak_debt"])],
                      ["ריבית שנתית", percent(m["finance"]["annual_rate"])]] +
                     [[l["label"], full_money(l["amount"])] for l in m["finance"]["lines"]],
                     ["סעיף", "ערך"])))
    panels.append(
        '<div class="panel" id="t5"><div class="card"><h2>מחיר מכירה מול עלות בנייה</h2>%s'
        '<p class="note">כל תא = רווח יזמי כאחוז מהעלויות. ירוק — מעל סף %s; '
        'צהוב — עד 3 נקודות מתחת לסף; אדום — מתחת לכך; אדום מלא — הפסד.</p></div>'
        '<div class="card"><h2>תרחישי ריבית</h2>%s</div></div>'
        % (heatmap(data["sensitivity"], r["profit_threshold"]),
           percent(r["profit_threshold"], 0),
           bar_table(rate_rows, ["שינוי בריבית", "ריבית שנתית", "עלות ריבית",
                                 "רווח יזמי", "% מהעלויות"])))
    panels.append('<div class="panel" id="t6"><div class="card"><h2>דגלים ונקודות לבחינה</h2>'
                  '%s</div></div>' % flags_html)

    return """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>דוח כדאיות — %(title)s</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;700;900&display=swap" rel="stylesheet">
<style>%(css)s</style>
</head>
<body>
<div class="wrap">
<header>
  <h1>ניתוח כדאיות כלכלית</h1>
  <div class="sub">%(title)s</div>
  <div class="meta">%(meta)s</div>
</header>

<div class="verdict %(vcls)s">%(vtext)s<small>%(vsub)s</small></div>

<div class="kpis">%(kpis)s</div>

<div class="tabs">%(tabs)s</div>
%(panels)s

<footer>
  <strong>%(disclaimer)s</strong><br>
  הדשבורד נבנה ממודל כלכלי יחיד; המספרים כאן זהים לאלה שבקובץ ה-Excel, בדוח ובמצגת.
  הנחות היסוד, המקורות ותאריכי השליפה מפורטים בגיליון "הנחות יסוד" שבמודל.
</footer>
</div>
<script>%(js)s</script>
</body>
</html>""" % {
        "title": esc(meta.get("project_name", "")),
        "meta": esc(" · ".join(x for x in [meta.get("client"), meta.get("location"),
                                           ident, meta.get("date")] if x)),
        "css": CSS, "js": JS,
        "vcls": vcls, "vtext": esc(vtext), "vsub": esc(vsub),
        "kpis": kpi_html, "tabs": tabs_html, "panels": "".join(panels),
        "disclaimer": esc(DISCLAIMER),
    }


def main():
    ap = argparse.ArgumentParser(description="בונה דשבורד HTML standalone")
    ap.add_argument("project")
    ap.add_argument("-o", "--output", default="דשבורד.html")
    args = ap.parse_args()
    data = full_output(load_project(args.project))
    with open(args.output, "w", encoding="utf-8") as fh:
        fh.write(build_html(data))
    print("נוצר: %s" % args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
