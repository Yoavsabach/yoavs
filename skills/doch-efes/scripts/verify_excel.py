#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_excel.py — מוודא שהנוסחאות באקסל מחשבות את מה שהמודל חישב.

למה זה קיים: openpyxl כותב נוסחאות אבל לא מריץ אותן, ולכן קובץ יכול להיראות
תקין לגמרי ולהכיל נוסחה שמצביעה לתא שגוי. בדיוק זה קרה בפיתוח — עמודת תזרים
ההון הפנתה לתא הרווח באחת שורה מטה, וה-IRR באקסל יצא 0% בזמן שהמודל הראה 19.8%.
בלי הרצה בפועל אין דרך לתפוס את זה בעין.

הסקריפט מריץ את חוברת העבודה במנוע ``formulas``, שולף את תאי התוצאה ומשווה
אותם למה ש-``model.py`` מחשב. פער מעל הסובלנות = כשל.

שימוש:
    python3 verify_excel.py project.json "מודל כלכלי.xlsx"

מחזיר קוד יציאה 1 אם נמצא פער, כך שאפשר לשרשר אותו לבדיקה אוטומטית.
"""

from __future__ import annotations

import argparse
import os
import sys
import warnings

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from model import build_model, load_project  # noqa: E402

# התאים בגיליון "תוצאות" ומה הם אמורים להיות במודל. התוויות משמשות לאיתור
# השורה לפי עמודה A, כדי שהבדיקה לא תישבר כשמוסיפים שורה לגיליון.
CHECKS = [
    ("סך עלויות הפרויקט", "total_cost", 1.0),
    ("סך הכנסות (נטו ממע\"מ)", "total_revenue_net", 1.0),
    ("רווח יזמי לפני מס", "profit_before_tax", 1.0),
    ("רווח יזמי — % מהעלויות", "margin_on_cost", 0.0005),
    ("רווח יזמי — % מההכנסות", "margin_on_revenue", 0.0005),
    ("מס על הרווח", "profit_tax", 1.0),
    ("רווח לאחר מס", "profit_after_tax", 1.0),
    ("IRR שנתי על ההון העצמי", "irr_annual", 0.005),
    ("תשואה על ההון (ROC)", "roc", 0.005),
    ("עלות למ\"ר מכור", "cost_per_sold_sqm", 1.0),
    ("שיא ניצול אשראי", "peak_debt", 1.0),
]


def evaluate(xlsx_path):
    """מריץ את החוברת ומחזיר מיפוי 'גיליון!תא' → ערך."""
    try:
        import formulas
    except ImportError:
        return None
    warnings.filterwarnings("ignore")
    model = formulas.ExcelModel().loads(xlsx_path).finish()
    sol = model.calculate()
    out = {}
    for key, val in sol.items():
        # המפתח נראה כך: "'[model.xlsx]תוצאות'!B4"
        if "]" not in key or "!" not in key:
            continue
        sheet = key.split("]", 1)[1].split("'", 1)[0]
        cell = key.rsplit("!", 1)[1]
        try:
            v = val.value[0, 0]
        except Exception:
            continue
        out["%s!%s" % (sheet, cell)] = v
    return out


def main():
    ap = argparse.ArgumentParser(description="משווה את תוצאות האקסל למודל")
    ap.add_argument("project")
    ap.add_argument("xlsx")
    args = ap.parse_args()

    expected = build_model(load_project(args.project))["results"]
    cells = evaluate(args.xlsx)
    if cells is None:
        print("המנוע 'formulas' אינו מותקן — לא ניתן לאמת את הנוסחאות.")
        print("התקן:  pip install formulas    ואז הרץ שוב.")
        return 2

    # איתור שורות לפי התווית בעמודה A של גיליון התוצאות
    labels = {}
    for key, val in cells.items():
        if key.startswith("תוצאות!A") and isinstance(val, str):
            labels[val.strip()] = key.split("!A")[1]

    failures, checked = [], 0
    print("%-34s %18s %18s" % ("מדד", "אקסל", "מודל"))
    print("-" * 74)
    for label, model_key, tol in CHECKS:
        row = labels.get(label)
        if row is None:
            failures.append("לא נמצאה השורה '%s' בגיליון התוצאות" % label)
            continue
        got = cells.get("תוצאות!B%s" % row)
        want = expected.get(model_key)
        if want is None:
            continue
        if not isinstance(got, (int, float)):
            failures.append("%s: האקסל החזיר %r ולא מספר" % (label, got))
            continue
        checked += 1
        delta = abs(float(got) - float(want))
        mark = "" if delta <= tol else "  ← פער"
        if delta > tol:
            failures.append("%s: אקסל %.4f מול מודל %.4f (פער %.4f)" % (label, got, want, delta))
        print("%-34s %18.4f %18.4f%s" % (label, got, want, mark))

    print("-" * 74)
    if failures:
        print("נמצאו %d פערים:" % len(failures))
        for f in failures:
            print("  • %s" % f)
        return 1
    print("כל %d התאים שנבדקו תואמים למודל." % checked)
    return 0


if __name__ == "__main__":
    sys.exit(main())
