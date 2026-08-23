#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
model.py — מנוע החישוב של דוח האפס.

זהו המקור היחיד לאמת המספרית. ארבעת בוני התוצרים (Excel / Word / PPTX / דשבורד)
כולם קוראים ל-``build_model`` ומקבלים בדיוק את אותם מספרים, כך ששקף במצגת לא יכול
לסתור תא באקסל. אם אתה מוסיף חישוב — הוסף אותו כאן, לא בבונה תוצר בודד.

קלט: ``project.json`` (הסכימה המלאה ב-references/project-schema.md, ודוגמה
ב-assets/project.example.json).
פלט: dict תוצאות + ``results.json`` על הדיסק.

שימוש:
    python3 model.py project.json -o results.json
    from model import load_project, build_model
"""

from __future__ import annotations

import argparse
import copy
import json
import sys
from datetime import date, timedelta

# ---------------------------------------------------------------------------
# עזרי קריאה סלחניים
#
# קלטי המשתמש בדוח אפס מגיעים מהתכתבות, לא מטופס. עדיף לקרוא "8.5" או "8.5%"
# או 0.085 ולהבין נכון, מאשר להתרסק על ValueError באמצע הרצה מול לקוח.
# ---------------------------------------------------------------------------


def num(value, default=0.0):
    """ממיר לערך מספרי. מקבל None, מחרוזות עם פסיקים, ₪, % ורווחים."""
    if value is None or value == "":
        return float(default)
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).strip().replace(",", "").replace("₪", "").replace("%", "")
    cleaned = cleaned.replace("‏", "").replace("‎", "").strip()
    if not cleaned:
        return float(default)
    try:
        return float(cleaned)
    except ValueError:
        return float(default)


def pct(value, default=0.0):
    """ממיר אחוז לשבר עשרוני. **כל שדה שמסתיים ב-_pct נמסר ביחידות אחוז**:
    6 פירושו 6%, 0.7 פירושו 0.7%, 23 פירושו 23%.

    אין כאן ניחוש. גרסה מוקדמת ניסתה לזהות לבד אם הכותב התכוון ל-6 או ל-0.06,
    והכלל "ערך מעל 1 הוא אחוזים" הפך שכ"ט של 1% לשכ"ט של 100% ועמלת אשראי של
    0.75% לעמלה של 75% — סטיות של עשרות מיליוני ש"ח שנראות לגמרי סבירות בתא
    אקסל. בדוח כדאיות עדיף כלל אחד נוקשה שהמשתמש חייב להכיר, על פני היוריסטיקה
    שנכשלת בשקט. אם הוזן 0.06 בכוונה, המודל יקרא זאת כ-0.06% — ולכן בדיקת
    השפיות בסוף ההרצה חייבת לכלול מבט על שיעורי האחוזים בגיליון ההנחות.
    """
    return num(value, default) / 100.0


def _quarter_label(d: date) -> str:
    return "Q%d/%d" % ((d.month - 1) // 3 + 1, d.year)


def _parse_date(value, fallback: date) -> date:
    if not value:
        return fallback
    if isinstance(value, date):
        return value
    txt = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%Y/%m/%d"):
        try:
            from datetime import datetime

            return datetime.strptime(txt, fmt).date()
        except ValueError:
            continue
    return fallback


def _spread(total: float, n: int, curve=None):
    """פורס סכום על פני n רבעונים לפי עקומה נתונה (מנורמלת), או אחיד."""
    if n <= 0:
        return []
    if curve:
        weights = [max(0.0, num(w)) for w in curve][:n]
        weights += [0.0] * (n - len(weights))
        s = sum(weights)
        if s > 0:
            return [total * w / s for w in weights]
    return [total / n] * n


# ---------------------------------------------------------------------------
# טעינת הפרויקט
# ---------------------------------------------------------------------------


def load_project(path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


# ---------------------------------------------------------------------------
# המודל
# ---------------------------------------------------------------------------


def build_model(project: dict, price_delta: float = 0.0, cost_delta: float = 0.0,
                rate_delta_pp: float = 0.0) -> dict:
    """מחשב את המודל המלא.

    ``price_delta``/``cost_delta`` הם שברים (0.1 = +10%) ומשמשים את טבלאות
    הרגישות; ``rate_delta_pp`` הוא הפרש בנקודות אחוז על ריבית המימון.
    הפונקציה טהורה — היא לא נוגעת ב-project — ולכן אפשר להריץ אותה מאות פעמים
    לבניית מטריצת רגישות.
    """
    p = project

    # ---------------- מע"מ ומיסוי ----------------
    tax = p.get("tax", {}) or {}
    vat = pct(tax.get("vat_pct"), 18.0)
    developer_type = (tax.get("developer_type") or "company").strip()
    corporate_rate = pct(tax.get("corporate_tax_pct"), 23.0)
    individual_rate = pct(tax.get("individual_tax_pct"), 50.0)
    profit_tax_rate = corporate_rate if developer_type == "company" else individual_rate

    # ---------------- קרקע ורכישה ----------------
    land = p.get("land", {}) or {}
    land_price = num(land.get("purchase_price"))
    purchase_tax = land_price * pct(land.get("purchase_tax_rate"), 6.0)
    land_legal = land_price * pct(land.get("legal_pct"), 0.0) + num(land.get("legal_amount"))
    brokerage = land_price * pct(land.get("brokerage_pct"), 0.0) + num(land.get("brokerage_amount"))
    land_other = num(land.get("development_levies")) + num(land.get("other"))
    land_total = land_price + purchase_tax + land_legal + brokerage + land_other

    land_lines = [
        {"label": "מחיר רכישת הקרקע", "amount": land_price},
        {"label": "מס רכישה", "amount": purchase_tax},
        {"label": "שכ\"ט עו\"ד ועלויות עסקה", "amount": land_legal},
        {"label": "דמי תיווך", "amount": brokerage},
        {"label": "היטלי פיתוח ואחר", "amount": land_other},
    ]

    # ---------------- עלויות בנייה ישירות ----------------
    # מדד תשומות הבנייה: הצמדה קדימה מהמדד הידוע למועד הביצוע הצפוי.
    index_uplift = pct((p.get("indexation") or {}).get("build_index_uplift_pct"), 0.0)
    cost_factor = (1.0 + index_uplift) * (1.0 + cost_delta)

    direct_lines = []
    for area in p.get("areas", []) or []:
        sqm = num(area.get("sqm"))
        rate = num(area.get("cost_per_sqm"))
        amount = sqm * rate * cost_factor
        direct_lines.append({
            "label": area.get("label") or area.get("use") or "שטח",
            "use": area.get("use") or "",
            "sqm": sqm,
            "cost_per_sqm": rate * cost_factor,
            "amount": amount,
            "source": area.get("source", ""),
        })

    parking = p.get("parking", {}) or {}
    parking_spaces = num(parking.get("spaces"))
    if parking_spaces:
        rate = num(parking.get("cost_per_space"))
        direct_lines.append({
            "label": "חניון תת-קרקעי (%d מקומות)" % int(parking_spaces),
            "use": "חניון",
            "sqm": parking_spaces * num(parking.get("sqm_per_space"), 35.0),
            "cost_per_sqm": rate * cost_factor / max(num(parking.get("sqm_per_space"), 35.0), 1),
            "amount": parking_spaces * rate * cost_factor,
            "source": parking.get("source", ""),
        })

    for extra in p.get("direct_other", []) or []:
        direct_lines.append({
            "label": extra.get("label", "עלות ישירה נוספת"),
            "use": extra.get("use", ""),
            "sqm": num(extra.get("sqm")),
            "cost_per_sqm": 0.0,
            "amount": num(extra.get("amount")) * cost_factor,
            "source": extra.get("source", ""),
        })

    direct_total = sum(line["amount"] for line in direct_lines)

    # ---------------- הכנסות ----------------
    # מחירי מכירה למגורים ברוטו כוללים מע"מ. המודל עובד נטו כדי שרווח יזמי
    # לא ינופח במע"מ שממילא מועבר לרשות המסים.
    revenue = p.get("revenue", {}) or {}
    prices_include_vat = revenue.get("prices_include_vat", True)
    vat_divisor = (1.0 + vat) if prices_include_vat else 1.0

    revenue_lines = []
    for item in revenue.get("items", []) or []:
        units = num(item.get("units"), 1)
        avg_sqm = num(item.get("avg_sqm"))
        ppsqm = num(item.get("price_per_sqm")) * (1.0 + price_delta)
        if item.get("total_price") not in (None, ""):
            gross = num(item.get("total_price")) * (1.0 + price_delta)
        else:
            gross = units * avg_sqm * ppsqm
        revenue_lines.append({
            "label": item.get("label", "רכיב הכנסה"),
            "units": units,
            "avg_sqm": avg_sqm,
            "price_per_sqm": ppsqm,
            "gross": gross,
            "net": gross / vat_divisor,
            "sellable_sqm": units * avg_sqm,
            "kind": item.get("kind", "sale"),
            "total_price": num(item.get("total_price")) if item.get("total_price") not in (None, "") else None,
            "source": item.get("source", ""),
        })

    # ---- נכס מניב ----
    # שני זרמים נפרדים, ולא אחד: דמי השכירות שנצברים בשנות ההחזקה, ושווי המימוש
    # המהוון בסופן. גרסה קודמת חישבה רק את המימוש, וכך "נעלמו" שנות ההשכרה —
    # בפרויקט לוגיסטי טיפוסי אלה עשרות מיליוני ש"ח שנפלו בין הכיסאות.
    income_asset = p.get("income_asset") or {}
    exit_value_gross = 0.0
    noi = num(income_asset.get("annual_noi"))
    cap_rate = pct(income_asset.get("cap_rate"), 0.0)
    holding_years = num(income_asset.get("holding_years"), 0)
    asset_sqm = num(income_asset.get("sqm"))

    if noi and holding_years:
        operating_gross = noi * holding_years * (1.0 + price_delta)
        revenue_lines.append({
            "label": "הכנסות שכירות בתקופת ההחזקה (%g שנים)" % holding_years,
            "units": holding_years,
            "avg_sqm": asset_sqm,
            "price_per_sqm": (noi / asset_sqm) if asset_sqm else 0.0,
            "gross": operating_gross,
            "net": operating_gross / vat_divisor,
            # הכנסה תפעולית אינה מכירת שטח. ספירתה כשטח מכיר מנפחת את
            # "עלות למ"ר מכור" פי מספר שנות ההחזקה והופכת אותו למספר חסר פשר.
            "sellable_sqm": 0.0,
            "kind": "operating",
            "source": income_asset.get("source", "") or "NOI שנתי × שנות החזקה",
        })

    if noi and cap_rate:
        exit_value_gross = (noi / cap_rate) * (1.0 + price_delta)
        revenue_lines.append({
            "label": income_asset.get("label", "מימוש נכס מניב (Exit)"),
            "units": 1,
            "avg_sqm": asset_sqm,
            "price_per_sqm": (exit_value_gross / asset_sqm) if asset_sqm else 0.0,
            "gross": exit_value_gross,
            "net": exit_value_gross / vat_divisor,
            "sellable_sqm": asset_sqm,
            "kind": "exit",
            "source": "היוון NOI בשיעור %.2f%%" % (cap_rate * 100),
        })

    revenue_gross = sum(r["gross"] for r in revenue_lines)
    revenue_net = sum(r["net"] for r in revenue_lines)
    sellable_sqm = sum(r["sellable_sqm"] for r in revenue_lines)

    # ---------------- עלויות עקיפות ----------------
    indirect_lines = []
    for item in p.get("indirect", []) or []:
        basis = (item.get("basis") or "direct").strip()
        if item.get("amount") not in (None, ""):
            amount = num(item.get("amount"))
        else:
            rate = pct(item.get("pct"))
            base = {
                "direct": direct_total,
                "land": land_total,
                "revenue": revenue_net,
                "revenue_gross": revenue_gross,
                "direct_and_land": direct_total + land_total,
            }.get(basis, direct_total)
            amount = rate * base
        indirect_lines.append({
            "label": item.get("label", "עלות עקיפה"),
            "basis": basis,
            "pct": pct(item.get("pct")),
            "amount": amount,
            "group": item.get("group", ""),
            "source": item.get("source", ""),
        })

    # ---- עלויות הטיפול בדיירים (התחדשות עירונית) ----
    # תקן 21 ס' 4.14(ד) מטפל בהן כפרק עצמאי, לא כשורה בתוך העקיפות — וזה נכון
    # מעשית: זה המספר הראשון שהדיירים, הרשות והשמאי שואלים עליו. הפריטים נכנסים
    # לחישוב יחד עם שאר העקיפות, אבל נשמרים גם כקבוצה נפרדת עם סכום משלה.
    tenants = p.get("tenants") or {}
    tenant_lines = []
    if tenants:
        t_units = num(tenants.get("units"))
        rent = num(tenants.get("rent_per_unit_month")) * num(tenants.get("rent_months")) * t_units
        if rent:
            tenant_lines.append({"label": "דמי שכירות לדיירים בתקופת הבנייה",
                                 "amount": rent, "source": tenants.get("rent_source", "")})
        for key, label in [("moving_per_unit", "הובלות ואחסון"),
                           ("legal_per_unit", "ליווי משפטי לדיירים"),
                           ("consultants_per_unit", "יועצים ומפקח מטעם הדיירים")]:
            v = num(tenants.get(key)) * t_units
            if v:
                tenant_lines.append({"label": "%s (%d יח\"ד)" % (label, int(t_units)),
                                     "amount": v, "source": tenants.get("source", "")})
        for key, label in [("organizer", "מארגן וארגון דיירים"),
                           ("betterment_for_tenants", "מיסוי בגין רכישת זכויות הדיירים"),
                           ("other", "עלויות דיירים אחרות")]:
            v = num(tenants.get(key))
            if v:
                tenant_lines.append({"label": label, "amount": v,
                                     "source": tenants.get("source", "")})
        t_base = sum(l["amount"] for l in tenant_lines)
        t_cont = pct(tenants.get("contingency_pct"), 0.0) * t_base
        if t_cont:
            tenant_lines.append({
                "label": "בצ\"מ ייעודי לדיירים (סרבנים, הארכת שכירות)",
                "amount": t_cont,
                "source": "אחוז מעלויות הדיירים, לפי ההנחות"})
        for line in tenant_lines:
            indirect_lines.append({"label": line["label"], "basis": "fixed", "pct": 0.0,
                                   "amount": line["amount"], "group": "tenants",
                                   "source": line.get("source", "")})

    tenants_total = sum(l["amount"] for l in tenant_lines)

    betterment = num(tax.get("betterment_levy"))
    if betterment:
        indirect_lines.append({
            "label": "היטל השבחה", "basis": "fixed", "pct": 0.0,
            "amount": betterment, "source": tax.get("betterment_source", ""),
        })
    rmi_fee = num(tax.get("rmi_permit_fee"))
    if rmi_fee:
        indirect_lines.append({
            "label": "דמי היתר רמ\"י", "basis": "fixed", "pct": 0.0,
            "amount": rmi_fee, "source": tax.get("rmi_source", ""),
        })

    indirect_total = sum(line["amount"] for line in indirect_lines)

    # ---------------- בצ"מ ----------------
    cont_pct = pct(p.get("contingency_pct"), 5.0)
    cont_base_key = (p.get("contingency_base") or "direct_and_indirect").strip()
    cont_base = {
        "direct": direct_total,
        "direct_and_indirect": direct_total + indirect_total,
        "total": direct_total + indirect_total + land_total,
    }.get(cont_base_key, direct_total + indirect_total)
    contingency = cont_pct * cont_base

    cost_before_finance = land_total + direct_total + indirect_total + contingency

    # ---------------- לוח זמנים ותזרים ----------------
    sched = p.get("schedule", {}) or {}
    start = _parse_date(sched.get("start"), date.today())
    quarters = int(num(sched.get("quarters"), 12))
    quarters = max(quarters, 1)

    quarter_dates = []
    d = date(start.year, ((start.month - 1) // 3) * 3 + 1, 1)
    for _ in range(quarters):
        quarter_dates.append(d)
        d = (d + timedelta(days=100)).replace(day=1)
        d = date(d.year, ((d.month - 1) // 3) * 3 + 1, 1)

    # הקרקע נרכשת בפועל ברבעון אחד; שאר העלויות נפרסות לפי עקומות.
    land_curve = sched.get("land_curve") or ([1] + [0] * (quarters - 1))
    land_flow = _spread(land_total, quarters, land_curve)
    direct_flow = _spread(direct_total, quarters, sched.get("cost_curve"))
    indirect_flow = _spread(indirect_total, quarters, sched.get("indirect_curve") or sched.get("cost_curve"))
    cont_flow = _spread(contingency, quarters, sched.get("cost_curve"))
    revenue_flow = _spread(revenue_net, quarters, sched.get("revenue_curve"))

    # ---------------- מימון ----------------
    fin = p.get("finance", {}) or {}
    equity = num(fin.get("equity"))
    annual_rate = pct(fin.get("credit_rate_pct"), 0.0)
    if not annual_rate:
        annual_rate = pct(fin.get("prime"), 0.0) + pct(fin.get("spread"), 0.0)
    annual_rate += rate_delta_pp / 100.0
    q_rate = annual_rate / 4.0

    # תזרים רבעוני עם ריבית מצטברת על היתרה המנוצלת של אשראי הליווי.
    # ההון העצמי נספג ראשון; רק מה שמעבר לו נושא ריבית.
    equity_left = equity
    balance = 0.0            # יתרת אשראי מנוצל (חיובית = חוב)
    peak_debt = 0.0
    interest_total = 0.0
    rows = []
    for i in range(quarters):
        out = land_flow[i] + direct_flow[i] + indirect_flow[i] + cont_flow[i]
        inflow = revenue_flow[i]
        net = inflow - out

        equity_used = 0.0
        if net < 0 and equity_left > 0:
            equity_used = min(equity_left, -net)
            equity_left -= equity_used
            net += equity_used

        balance = max(0.0, balance - net)
        interest = balance * q_rate
        interest_total += interest
        balance += interest
        peak_debt = max(peak_debt, balance)

        rows.append({
            "quarter": _quarter_label(quarter_dates[i]),
            "date": quarter_dates[i].isoformat(),
            "land": land_flow[i],
            "direct": direct_flow[i],
            "indirect": indirect_flow[i],
            "contingency": cont_flow[i],
            "outflow": out,
            "revenue": inflow,
            "equity_used": equity_used,
            "interest": interest,
            "debt_balance": balance,
            "net": inflow - out,
        })

    loan_limit = num(fin.get("loan_limit")) or peak_debt
    arrangement_fees = loan_limit * pct(fin.get("fees_pct"), 0.0)
    warranty_cost = revenue_gross * pct(fin.get("warranty_pct"), 0.0)
    finance_total = interest_total + arrangement_fees + warranty_cost

    finance_lines = [
        {"label": "ריבית אשראי ליווי", "amount": interest_total,
         "source": "ריבית שנתית %.2f%% על יתרה מנוצלת" % (annual_rate * 100)},
        {"label": "עמלות פתיחה וניהול אשראי", "amount": arrangement_fees, "source": ""},
        {"label": "ערבויות חוק מכר", "amount": warranty_cost, "source": ""},
    ]

    # ---------------- תוצאות ----------------
    total_cost = cost_before_finance + finance_total
    profit_before_tax = revenue_net - total_cost
    profit_tax = max(0.0, profit_before_tax) * profit_tax_rate
    profit_after_tax = profit_before_tax - profit_tax

    margin_on_cost = profit_before_tax / total_cost if total_cost else 0.0
    margin_on_revenue = profit_before_tax / revenue_net if revenue_net else 0.0
    roc = profit_after_tax / equity if equity else 0.0
    cost_per_sold_sqm = total_cost / sellable_sqm if sellable_sqm else 0.0

    # IRR על תזרים ההון העצמי: יציאות = הון שהוזרם, כניסה אחרונה = החזר ההון
    # שהוזרם בפועל בתוספת הרווח לאחר מס.
    equity_flows = [-r["equity_used"] for r in rows]
    equity_injected = sum(r["equity_used"] for r in rows)
    if equity_flows:
        equity_flows[-1] += equity_injected + profit_after_tax
    irr_annual = _irr(equity_flows)

    # IRR לא-ממונף: אותו פרויקט בלי אשראי. מפריד בין "הפרויקט טוב" לבין
    # "המינוף עשה את העבודה" — הבחנה שמשנה את ההחלטה כשהמרווח דק.
    unlev = [r["revenue"] - (r["outflow"]) for r in rows]
    irr_unlevered = _irr(unlev)

    project_years = len(rows) / 4.0
    roc = profit_after_tax / equity if equity else 0.0

    # ---- מדדי נכס מניב ----
    # לפרויקט להשכרה השאלה אינה "כמה אחוז רווח" אלא "האם התשואה על העלות
    # גבוהה מספיק מעל שיעור ההיוון שבו נמכור". מרווח דק הוא הסיכון האמיתי,
    # והוא בלתי נראה במדד הרווח היזמי.
    income_metrics = None
    if noi:
        yield_on_cost = noi / total_cost if total_cost else 0.0
        spread_bps = (yield_on_cost - cap_rate) * 10000 if cap_rate else None
        annual_debt_service = peak_debt * annual_rate
        income_metrics = {
            "annual_noi": noi,
            "holding_years": holding_years,
            "exit_cap_rate": cap_rate,
            "exit_value": exit_value_gross,
            "yield_on_cost": yield_on_cost,
            "spread_bps": spread_bps,
            "dscr": (noi / annual_debt_service) if annual_debt_service else None,
            "ltc": (peak_debt / total_cost) if total_cost else 0.0,
            "breakeven_cap_rate": (noi / total_cost) if total_cost else None,
        }

    # ---- סף הבקרה ----
    # סף הרווח היזמי של 15%–18% הוא נוהג של ייזום למכירה. cost-benchmarks.md
    # אומר במפורש שהוא לא אומת לנדל"ן מניב, ולכן אין להחיל אותו שם: בפרויקט
    # מניב המבחן הוא מרווח התשואה, ומדידה לפי הסף הלא נכון נותנת תשובה
    # בטוחה-למראה ושגויה.
    thresholds = p.get("thresholds") or {}
    kind = project_kind(p)
    if kind == "income":
        threshold = pct(thresholds.get("min_profit_on_cost_pct"), 0.0)
        min_spread = num(thresholds.get("min_spread_bps"), 150)
        threshold_basis = "spread"
    else:
        threshold = pct(thresholds.get("min_profit_on_cost_pct"), 15.0)
        min_spread = None
        threshold_basis = "margin"

    if kind == "income" and income_metrics and income_metrics["spread_bps"] is not None:
        meets = income_metrics["spread_bps"] >= min_spread
    else:
        meets = margin_on_cost >= threshold

    return {
        "meta": p.get("meta", {}),
        "kind": kind,
        "vat_pct": vat,
        "developer_type": developer_type,
        "profit_tax_rate": profit_tax_rate,
        "land": {"lines": land_lines, "total": land_total},
        "direct": {"lines": direct_lines, "total": direct_total},
        "indirect": {"lines": indirect_lines, "total": indirect_total},
        "tenants": {"lines": tenant_lines, "total": tenants_total},
        "contingency": {"pct": cont_pct, "base": cont_base, "amount": contingency},
        "finance": {
            "lines": finance_lines, "total": finance_total,
            "equity": equity, "annual_rate": annual_rate,
            "peak_debt": peak_debt, "loan_limit": loan_limit,
            "interest": interest_total,
        },
        "revenue": {
            "lines": revenue_lines, "gross": revenue_gross, "net": revenue_net,
            "sellable_sqm": sellable_sqm, "prices_include_vat": prices_include_vat,
        },
        "cashflow": rows,
        "results": {
            "total_cost": total_cost,
            "cost_before_finance": cost_before_finance,
            "total_revenue_net": revenue_net,
            "total_revenue_gross": revenue_gross,
            "profit_before_tax": profit_before_tax,
            "profit_tax": profit_tax,
            "profit_after_tax": profit_after_tax,
            "margin_on_cost": margin_on_cost,
            "margin_on_revenue": margin_on_revenue,
            "irr_annual": irr_annual,
            "irr_unlevered": irr_unlevered,
            # תשואה מצטברת על פני חיי הפרויקט, לא שנתית. התווית חשובה: 15%
            # על פני חמש שנים ו-15% בשנה הם שני דברים שונים לגמרי.
            "roc": roc,
            "project_years": project_years,
            "cost_per_sold_sqm": cost_per_sold_sqm,
            "peak_debt": peak_debt,
            "equity_injected": equity_injected,
            "profit_threshold": threshold,
            "min_spread_bps": min_spread,
            # לפי מה נמדדת ההצלחה: "margin" = רווח יזמי כאחוז מהעלויות
            # (ייזום למכירה), "spread" = מרווח התשואה מול ההיוון (נכס מניב).
            # הבונים קוראים את זה כדי לא לכתוב "מתחת לסף 15%" על פרויקט מניב.
            "threshold_basis": threshold_basis,
            "meets_threshold": meets,
        },
        "income_metrics": income_metrics,
    }


def project_kind(project):
    """מסווג את **כלכלת** הפרויקט. אל תבלבל בין זה לבין ``meta.mode``.

    ``kind`` נגזר מהנתונים וקובע כיצד הפרויקט נמדד:
      ``income``  — נכס מניב. המבחן הוא מרווח התשואה מול ההיוון, לא רווח יזמי.
      ``renewal`` — התחדשות עירונית: יש דיירים ודירות תמורה.
      ``sale``    — ייזום למכירה. המבחן הוא הרווח היזמי כאחוז מהעלויות.

    ``meta.mode`` הוא עניין אחר לגמרי — **מבנה הדוח** (``takan21`` / ``full`` /
    ``short``). גרסה קודמת גזרה kind מ-mode, וכך פרויקט מגורים רגיל שהוגש
    במבנה המלא סווג בטעות כהתחדשות עירונית. יזם יכול לרצות דוח מפורט בלי
    שיהיו בפרויקט דיירים.
    """
    explicit = ((project.get("meta") or {}).get("kind") or "").strip()
    if explicit in ("income", "renewal", "sale"):
        return explicit
    if project.get("income_asset"):
        return "income"
    if project.get("tenants"):
        return "renewal"
    return "sale"


def _irr(flows, lo=-0.95, hi=10.0, tol=1e-7, iters=200):
    """IRR רבעוני בחיפוש בינארי, מוחזר כשיעור שנתי.

    חיפוש בינארי ולא ניוטון-רפסון: תזרים של פרויקט נדל"ן משנה סימן פעם אחת
    בלבד, אז ה-NPV מונוטוני בטווח והבינארי מתכנס תמיד — בעוד ניוטון קופץ
    לאינסוף כשהנגזרת קטנה. מחזיר None כשאין שורש בטווח (למשל פרויקט מפסיד),
    כי מספר IRR שקרי גרוע מהיעדר מספר.
    """
    def npv(rate):
        return sum(f / ((1.0 + rate) ** i) for i, f in enumerate(flows))

    if not flows or all(f == 0 for f in flows):
        return None
    if npv(lo) * npv(hi) > 0:
        return None
    for _ in range(iters):
        mid = (lo + hi) / 2.0
        v = npv(mid)
        if abs(v) < tol:
            break
        if npv(lo) * v < 0:
            hi = mid
        else:
            lo = mid
    q = (lo + hi) / 2.0
    try:
        return (1.0 + q) ** 4 - 1.0
    except (OverflowError, ValueError):
        return None


# ---------------------------------------------------------------------------
# רגישויות
# ---------------------------------------------------------------------------


def sensitivity_grid(project, price_steps=None, cost_steps=None, metric="margin_on_cost"):
    """מטריצה דו-ממדית: מחיר מכירה × עלות בנייה."""
    sens = project.get("sensitivity", {}) or {}
    price_steps = price_steps or sens.get("price_steps") or [-10, -5, 0, 5, 10]
    cost_steps = cost_steps or sens.get("cost_steps") or [-10, -5, 0, 5, 10]
    grid = []
    for cs in cost_steps:
        row = []
        for ps in price_steps:
            m = build_model(project, price_delta=num(ps) / 100.0, cost_delta=num(cs) / 100.0)
            row.append(m["results"][metric])
        grid.append(row)
    return {"price_steps": [num(s) for s in price_steps],
            "cost_steps": [num(s) for s in cost_steps],
            "metric": metric, "grid": grid}


def break_even(project):
    """מוצא את נקודות המפנה: איפה הרווח מתאפס ואיפה הוא פוגש את הסף.

    תקן 21 ס' 4.15 מחייב את זה, אבל התועלת רחבה יותר — ליזם, "עד כמה המחיר
    יכול לרדת לפני שאני מפסיד" הוא מספר שימושי הרבה יותר מאחוז רווח בודד.
    נפתר בחיפוש בינארי על סטיית המחיר, כי הרווח מונוטוני עולה בה.
    """
    base = build_model(project)
    th = base["results"]["profit_threshold"]

    def solve(target):
        lo, hi = -0.95, 3.0
        f = lambda d: build_model(project, price_delta=d)["results"]["margin_on_cost"] - target
        if f(lo) > 0:
            return lo
        if f(hi) < 0:
            return None
        for _ in range(60):
            mid = (lo + hi) / 2
            if f(mid) < 0:
                lo = mid
            else:
                hi = mid
        return (lo + hi) / 2

    def price_at(delta):
        if delta is None:
            return None
        items = (project.get("revenue") or {}).get("items") or []
        if not items:
            return None
        return num(items[0].get("price_per_sqm")) * (1 + delta)

    zero_d = solve(0.0)
    th_d = solve(th) if th else None
    return {
        "zero_profit_price_delta": zero_d,
        "zero_profit_price_per_sqm": price_at(zero_d),
        "threshold_price_delta": th_d,
        "threshold_price_per_sqm": price_at(th_d),
        "threshold": th,
    }


def rate_scenarios(project, steps_pp=None, metric="margin_on_cost"):
    """תרחישי ריבית — הפרשים בנקודות אחוז מהריבית שבהנחות."""
    sens = project.get("sensitivity", {}) or {}
    steps_pp = steps_pp or sens.get("rate_steps_pp") or [-1, 0, 1, 2, 3]
    out = []
    for s in steps_pp:
        m = build_model(project, rate_delta_pp=num(s))
        out.append({
            "delta_pp": num(s),
            "rate": m["finance"]["annual_rate"],
            "metric": m["results"][metric],
            "profit": m["results"]["profit_before_tax"],
            "interest": m["finance"]["interest"],
        })
    return out


# ---------------------------------------------------------------------------
# דגלים אדומים
#
# הדגלים נגזרים מהמודל ולא נכתבים ביד, כדי שדוח שנבנה מחדש עם מספרים חדשים
# יעדכן את הדגלים אוטומטית. דגלים שהמשתמש/סקיל תכנוני הוסיף ידנית מצטרפים.
# ---------------------------------------------------------------------------


def derive_flags(project, model):
    flags = []
    r = model["results"]
    th = r["profit_threshold"]
    im = model.get("income_metrics")

    # בנכס מניב הכישלון מתבטא במרווח, לא ברווח היזמי. בלי הדגל הזה פרויקט
    # שנופל במבחן שלו עצמו פשוט לא מופיע ברשימת הדגלים.
    if r.get("threshold_basis") == "spread" and im and im.get("spread_bps") is not None:
        spread = im["spread_bps"]
        need = r.get("min_spread_bps") or 150
        if spread < 0:
            flags.append({
                "level": "red",
                "title": "מרווח שלילי — הנכס שווה פחות מעלות הקמתו",
                "text": "התשואה על העלות %.2f%% נמוכה משיעור ההיוון ביציאה %.2f%% "
                        "(מרווח %d נק' בסיס). בתום ההקמה שווי הנכס נמוך מהעלות "
                        "שהושקעה בו, וכל רווח שנותר מגיע מדמי השכירות בלבד."
                        % (im["yield_on_cost"] * 100, im["exit_cap_rate"] * 100, round(spread)),
            })
        elif spread < need:
            flags.append({
                "level": "red",
                "title": "מרווח תשואה צר מהנורמה",
                "text": "המרווח בין התשואה על העלות (%.2f%%) לשיעור ההיוון ביציאה "
                        "(%.2f%%) הוא %d נק' בסיס בלבד, מול %d הנדרשות. תנודה קטנה "
                        "בשכירות או בשיעור ההיוון מוחקת את הרווח."
                        % (im["yield_on_cost"] * 100, im["exit_cap_rate"] * 100,
                           round(spread), round(need)),
            })
        if im.get("dscr") is not None and im["dscr"] < 1.2:
            flags.append({
                "level": "red" if im["dscr"] < 1.0 else "amber",
                "title": "DSCR נמוך",
                "text": "יחס כיסוי שירות החוב %.2f. בנק מלווה נוהג לדרוש 1.2 ומעלה."
                        % im["dscr"],
            })
        if im.get("ltc", 0) > 0.70:
            flags.append({
                "level": "amber",
                "title": "LTC גבוה",
                "text": "החוב מהווה %.0f%% מסך העלות. מעל 70%% בנקים מצמצמים את "
                        "המסגרת או דורשים בטוחות נוספות." % (im["ltc"] * 100),
            })

    # רווח חריג כלפי מעלה הוא לרוב טעות קלט, לא בשורה טובה. עדיף לומר זאת
    # ליזם לפני שהוא מגיש את הדוח לבנק ומתבקש להסביר.
    if r.get("threshold_basis") == "margin" and r["margin_on_cost"] > 0.30:
        flags.append({
            "level": "amber",
            "title": "רווח יזמי גבוה מהמקובל — לבדוק את הקלטים",
            "text": "הרווח היזמי %.1f%% מהעלויות, מעל הטווח המקובל של 15%%–20%%. "
                    "לרוב זה מסמן קלט שגוי — מחיר קרקע נמוך מדי, עלות בנייה חסרה, "
                    "או מחיר מכירה אופטימי. אמת מול עסקאות השוואה לפני הגשה."
                    % (r["margin_on_cost"] * 100),
        })

    if r["margin_on_cost"] < th:
        flags.append({
            "level": "red",
            "title": "רווח יזמי מתחת לסף",
            "text": "הרווח היזמי עומד על %.1f%% מהעלויות, מתחת לסף %.0f%% שבנקים מלווים "
                    "נוהגים לדרוש. במתכונת זו הפרויקט עלול שלא לקבל ליווי בנקאי." %
                    (r["margin_on_cost"] * 100, th * 100),
        })
    elif r["margin_on_cost"] < th + 0.03:
        flags.append({
            "level": "amber",
            "title": "רווח יזמי בשולי הסף",
            "text": "הרווח היזמי %.1f%% מהעלויות — קרוב לסף %.0f%%. סטייה קטנה בעלויות "
                    "או במחירי המכירה מוציאה את הפרויקט מהתכנות." %
                    (r["margin_on_cost"] * 100, th * 100),
        })

    # רגישות: אם ירידת מחירים של 10% מוחקת את הרווח, זה סיכון מהותי בפני עצמו.
    downside = build_model(project, price_delta=-0.10)
    if downside["results"]["profit_before_tax"] <= 0:
        flags.append({
            "level": "red",
            "title": "רגישות קריטית לירידת מחירים",
            "text": "ירידה של 10%% במחירי המכירה מוחקת את הרווח כולו (רווח בתרחיש: %s ₪). "
                    "הפרויקט אינו סופג תנודת שוק רגילה." %
                    format(round(downside["results"]["profit_before_tax"]), ","),
        })

    equity = model["finance"]["equity"]
    if equity and model["results"]["peak_debt"] > 0:
        gearing = model["results"]["peak_debt"] / (equity + model["results"]["peak_debt"])
        if gearing > 0.80:
            flags.append({
                "level": "amber",
                "title": "מינוף גבוה",
                "text": "שיא החוב מהווה %.0f%% ממקורות המימון. מינוף בשיעור זה מייקר את "
                        "האשראי ומקטין את כרית הביטחון." % (gearing * 100),
            })
    if equity == 0:
        flags.append({
            "level": "red",
            "title": "לא הוזן הון עצמי",
            "text": "ללא הון עצמי לא ניתן להעריך IRR ותשואה על ההון, ובנק מלווה לא יעמיד ליווי.",
        })

    if model["contingency"]["pct"] < 0.05:
        flags.append({
            "level": "amber",
            "title": "בצ\"מ נמוך מהמקובל",
            "text": "בצ\"מ בשיעור %.1f%% — מתחת לטווח המקובל של 5%%–10%%. בדוק מול הבנק המלווה." %
                    (model["contingency"]["pct"] * 100),
        })

    # קצב ספיגה: כמה יחידות צריך למכור בחודש כדי לעמוד בלוח הזמנים.
    units = sum(num(i.get("units")) for i in (project.get("revenue", {}) or {}).get("items", []) or [])
    months = len(model["cashflow"]) * 3
    if units and months:
        pace = units / months
        if pace > 3:
            flags.append({
                "level": "amber",
                "title": "קצב מכירות אגרסיבי",
                "text": "לוח הזמנים מניח מכירת %.1f יח\"ד בחודש. קצב זה חורג מהמקובל "
                        "ומחייב אימות מול עסקאות השוואה באזור." % pace,
            })

    # התאמת התוויות ב-benchmarks נעשית לפי מחרוזת מדויקת. כשהיא נכשלת ההערה
    # פשוט לא מופיעה, והמשתמש מאמין שהאימות רץ. עדיף להגיד שהוא לא רץ.
    for bm in project.get("benchmarks", []) or []:
        if not bm.get("label"):
            continue
        if not _label_exists(project, bm["label"]):
            flags.append({
                "level": "amber",
                "title": "אסמכתה ללא התאמה בגיליון ההנחות",
                "text": "הוגדרה אסמכתה עבור \"%s\", אבל אין פרמטר בשם הזה בגיליון "
                        "ההנחות ולכן הערת הסטייה לא תיצמד לשום שורה. תקן את התווית "
                        "כך שתהיה זהה." % bm["label"],
            })

    for extra in project.get("flags", []) or []:
        flags.append({
            "level": extra.get("level", "amber"),
            "title": extra.get("title", "דגל"),
            "text": extra.get("text", ""),
            "source": extra.get("source", ""),
        })
    return flags


# ---------------------------------------------------------------------------
# בדיקת סטייה מול אסמכתאות חיצוניות (ספי 10%)
# ---------------------------------------------------------------------------


def _fmt_val(v):
    """מעצב ערך להערת סטייה. עיגול לשלם הפך שיעור של 6.75% ל-"7" ואת האסמכתה
    5.2% ל-"5", כך שההערה דיווחה על פער שאינו קיים. ערך קטן מ-100 הוא כמעט
    תמיד שיעור אחוז ולא סכום, ולכן נשמרות לו שתי ספרות."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(f) < 100:
        return ("%.2f" % f).rstrip("0").rstrip(".")
    return format(int(round(f)), ",")


def deviation_notes(project, tolerance=0.10):
    """מסמן קלטי משתמש שסוטים מעל הסף מהאסמכתה החיצונית שנמצאה.

    הפונקציה לא משנה שום קלט — היא רק מייצרת הערה. ההחלטה על מספר היא של היזם
    והשמאי, לא של הכלי; מה שהכלי חייב לעשות הוא לא לתת לסטייה לעבור בשקט.
    """
    notes = []
    for bm in project.get("benchmarks", []) or []:
        used = num(bm.get("used"))
        ref = num(bm.get("reference"))
        if not ref:
            continue
        dev = (used - ref) / ref
        if abs(dev) > tolerance:
            notes.append({
                "label": bm.get("label", ""),
                "used": used,
                "reference": ref,
                "deviation": dev,
                "source": bm.get("source", ""),
                "source_date": bm.get("source_date", ""),
                "text": "%s: הוזן %s מול אסמכתה %s (%s%.0f%%). מקור: %s%s" % (
                    bm.get("label", ""),
                    _fmt_val(used), _fmt_val(ref),
                    "+" if dev > 0 else "", dev * 100,
                    bm.get("source", "לא צוין"),
                    (", " + bm["source_date"]) if bm.get("source_date") else "",
                ),
            })
    return notes


def _label_exists(project, label):
    """האם התווית קיימת כפרמטר כלשהו שייכתב לגיליון ההנחות."""
    for area in project.get("areas", []) or []:
        name = area.get("label") or area.get("use") or ""
        if label in ("שטח — %s" % name, "עלות למ\"ר — %s" % name):
            return True
    for item in (project.get("revenue") or {}).get("items", []) or []:
        name = item.get("label", "")
        if label in ("כמות — %s" % name, "שטח ממוצע — %s" % name,
                     "מחיר מכירה למ\"ר — %s" % name):
            return True
    for item in project.get("indirect", []) or []:
        if label == item.get("label"):
            return True
    known = {"מחיר רכישת הקרקע", "שיעור מס רכישה", "שכ\"ט עו\"ד ועלויות עסקה",
             "דמי תיווך", "היטלי פיתוח בגין הקרקע", "עלויות קרקע אחרות",
             "הצמדת עלויות למדד תשומות הבנייה", "מקומות חניה תת-קרקעיים",
             "עלות למקום חניה", "שיעור מע\"מ", "היטל השבחה", "דמי היתר רמ\"י",
             "בלתי צפוי מראש (בצ\"מ)", "הון עצמי", "ריבית אשראי ליווי שנתית",
             "NOI שנתי", "שיעור היוון (Cap Rate)", "שנות החזקה עד המימוש"}
    return label in known


DISCLAIMER = ("מסמך זה הינו ניתוח כדאיות כלכלית ואינו מהווה דוח אפס של שמאי מקרקעין "
              "מוסמך לצורכי ליווי בנקאי.")


def full_output(project):
    """אורז את כל מה שבוני התוצרים צריכים למבנה אחד."""
    model = build_model(project)
    return {
        "project": project,
        "model": model,
        "break_even": break_even(project),
        "sensitivity": sensitivity_grid(project),
        "sensitivity_profit": sensitivity_grid(project, metric="profit_before_tax"),
        "rate_scenarios": rate_scenarios(project),
        "flags": derive_flags(project, model),
        "deviations": deviation_notes(project),
        "disclaimer": DISCLAIMER,
    }


def main():
    ap = argparse.ArgumentParser(description="מנוע החישוב של דוח האפס")
    ap.add_argument("project", help="נתיב ל-project.json")
    ap.add_argument("-o", "--output", default="results.json")
    args = ap.parse_args()

    project = load_project(args.project)
    out = full_output(project)
    with open(args.output, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=2)

    r = out["model"]["results"]
    print("נכתב: %s" % args.output)
    print("סך עלויות:      %15s ₪" % format(round(r["total_cost"]), ","))
    print("סך הכנסות (נטו): %15s ₪" % format(round(r["total_revenue_net"]), ","))
    print("רווח יזמי:       %15s ₪  (%.1f%% מהעלויות, %.1f%% מההכנסות)" % (
        format(round(r["profit_before_tax"]), ","),
        r["margin_on_cost"] * 100, r["margin_on_revenue"] * 100))
    print("IRR שנתי:        %15s" % ("%.1f%%" % (r["irr_annual"] * 100) if r["irr_annual"] is not None else "לא ניתן לחישוב"))
    print("דגלים: %d" % len(out["flags"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
