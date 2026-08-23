#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
regression_test.py — בונה את כל התוצרים לשלושת סוגי הפרויקטים ומאמת אותם.

שלושת הסוגים נבדלים לא רק במספרים אלא בלוגיקה: ייזום למכירה נמדד ברווח יזמי,
נכס מניב נמדד במרווח התשואה, והתחדשות עירונית מוסיפה פרק עלויות דיירים וכפל
שטחים. תיקון שעובד לאחד יכול לשבור אחר בשקט — כפי שקרה כשעלויות הדיירים נכנסו
למנוע ולא לאקסל, והשניים הציגו סכומים שונים.

שימוש:
    python3 regression_test.py            # שלושת הסוגים
    python3 regression_test.py --quick    # בלי אימות הנוסחאות (מהיר)
"""

import argparse
import copy
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
sys.path.insert(0, HERE)

from model import build_model, load_project  # noqa: E402


def fixtures():
    base = load_project(os.path.join(SKILL, "assets", "project.example.json"))

    # התרחישים נגזרים מהדוגמה, אבל ה-kind המפורש שבה חייב לרדת — אחרת הבדיקה
    # מאמתת ערך שהועתק במקום את ההיסק שהיא אמורה לבדוק.
    income = copy.deepcopy(base)
    income["meta"].pop("kind", None)
    income["meta"].update(project_name="מרכז לוגיסטי — בדיקת רגרסיה", mode="short")
    income["revenue"] = {"prices_include_vat": False, "items": []}
    income["income_asset"] = {"annual_noi": 8872848, "cap_rate": 7.25,
                              "holding_years": 3, "sqm": 19200,
                              "label": "מימוש הנכס", "source": "בדיקה"}
    income["areas"] = [{"use": "לוגיסטי", "label": "מחסנים", "sqm": 18000,
                        "cost_per_sqm": 3660, "source": "cost-benchmarks.md"}]
    income["parking"] = {}
    income["tax"]["betterment_levy"] = 0
    income["benchmarks"] = []

    renewal = copy.deepcopy(base)
    renewal["meta"].pop("kind", None)
    renewal["meta"].update(project_name="פינוי-בינוי — בדיקת רגרסיה", mode="takan21")
    renewal["tenants"] = {"units": 48, "rent_per_unit_month": 4500, "rent_months": 30,
                          "moving_per_unit": 12000, "organizer": 1500000,
                          "contingency_pct": 10, "source": "בדיקה"}
    renewal["assumptions_notes"] = [{"label": "מקדם שטחי שירות", "value": "28%",
                                     "source": "הנחה"}]
    return {"sale": base, "renewal": renewal, "income": income}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true",
                    help="דלג על אימות הנוסחאות (formulas איטי)")
    args = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix="doch-efes-regression-")
    failures = []

    for name, project in fixtures().items():
        path = os.path.join(tmp, "%s.json" % name)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(project, fh, ensure_ascii=False)

        model = build_model(project)
        print("── %s (kind=%s, מבחן=%s) ──"
              % (name, model["kind"], model["results"]["threshold_basis"]))
        if model["kind"] != name:
            failures.append("%s: הסיווג יצא %s" % (name, model["kind"]))

        for script, ext in [("build_excel", "xlsx"), ("build_report", "docx"),
                            ("build_deck", "pptx"), ("build_dashboard", "html")]:
            out = os.path.join(tmp, "%s.%s" % (name, ext))
            r = subprocess.run([sys.executable, os.path.join(HERE, script + ".py"),
                                path, "-o", out], capture_output=True, text=True)
            ok = r.returncode == 0 and os.path.exists(out) and os.path.getsize(out) > 1000
            print("   %-18s %s" % (script, "✓" if ok else "✗ " + r.stderr.strip()[:80]))
            if not ok:
                failures.append("%s/%s" % (name, script))

        if not args.quick:
            r = subprocess.run([sys.executable, os.path.join(HERE, "verify_excel.py"),
                                path, os.path.join(tmp, "%s.xlsx" % name)],
                               capture_output=True, text=True)
            last = [l for l in r.stdout.strip().split("\n") if l.strip()][-1:]
            print("   %-18s %s" % ("verify_excel", last[0] if last else "אין פלט"))
            if r.returncode != 0:
                failures.append("%s/verify_excel" % name)

    print()
    if failures:
        print("נכשלו: %s" % ", ".join(failures))
        return 1
    print("כל הבדיקות עברו. תוצרים ב-%s" % tmp)
    return 0


if __name__ == "__main__":
    sys.exit(main())
