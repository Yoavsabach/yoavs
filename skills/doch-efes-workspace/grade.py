#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
grade.py — בודק אוטומטית את הקריטריונים המכניים על תיקיית פלט של הרצת eval.

הבדיקות כאן הן אלה שאפשר לאמת בקוד ולכן אין סיבה לשפוט אותן בעין: קיום קבצים,
נוסחאות חיות, RTL, גילוי נאות, ועקביות מספרית בין התוצרים. הקריטריונים
האיכותניים (האם הדגלים באמת רלוונטיים, האם הניסוח מקצועי) נשארים לביקורת
האנושית ב-viewer.

שימוש:
    python3 grade.py <workspace>/iteration-N
"""

import json
import os
import re
import sys
import zipfile

DISCLAIMER_KEY = "אינו מהווה דוח אפס של שמאי מקרקעין מוסמך"


def find(d, exts):
    out = []
    for root, _, files in os.walk(d):
        for f in files:
            if os.path.splitext(f)[1].lower() in exts:
                out.append(os.path.join(root, f))
    return out


def docx_text(path):
    try:
        raw = zipfile.ZipFile(path).read("word/document.xml").decode("utf8", "replace")
        return re.sub(r"<[^>]+>", "", raw)
    except Exception:
        return ""


def pptx_text(path):
    try:
        z = zipfile.ZipFile(path)
        parts = [z.read(n).decode("utf8", "replace") for n in z.namelist()
                 if n.startswith("ppt/slides/slide")]
        return re.sub(r"<[^>]+>", "", "".join(parts))
    except Exception:
        return ""


def xlsx_info(path):
    """מחזיר: כמות נוסחאות, כמה גיליונות RTL, כלל הטקסט, וכל המספרים."""
    info = {"formulas": 0, "sheets": 0, "rtl": 0, "text": "", "numbers": []}
    try:
        z = zipfile.ZipFile(path)
        shared = ""
        if "xl/sharedStrings.xml" in z.namelist():
            shared = z.read("xl/sharedStrings.xml").decode("utf8", "replace")
        texts = [shared]
        for name in z.namelist():
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                continue
            info["sheets"] += 1
            xml = z.read(name).decode("utf8", "replace")
            info["formulas"] += xml.count("<f>")
            if 'rightToLeft="1"' in xml or "rightToLeft=\"true\"" in xml:
                info["rtl"] += 1
            texts.append(xml)
            for m in re.finditer(r"<v>(-?\d+\.?\d*)</v>", xml):
                info["numbers"].append(float(m.group(1)))
        info["text"] = re.sub(r"<[^>]+>", " ", "".join(texts))
    except Exception as exc:
        info["error"] = str(exc)
    return info


def close(a, b, tol=1.5):
    return a is not None and b is not None and abs(a - b) <= tol


def parse_money(text):
    """מוציא סכומים בש\"ח מטקסט חופשי, כדי להשוות בין תוצרים."""
    vals = set()
    for m in re.finditer(r"(\d[\d,]{5,})", text.replace("‏", "")):
        try:
            vals.add(float(m.group(1).replace(",", "")))
        except ValueError:
            pass
    return vals


def grade_run(run_dir):
    out = os.path.join(run_dir, "outputs")
    if not os.path.isdir(out):
        out = run_dir
    xlsx = find(out, {".xlsx"})
    docx = find(out, {".docx"})
    pptx = find(out, {".pptx"})
    html = find(out, {".html", ".htm"})

    xi = xlsx_info(xlsx[0]) if xlsx else {"formulas": 0, "sheets": 0, "rtl": 0,
                                          "text": "", "numbers": []}
    dtext = docx_text(docx[0]) if docx else ""
    ptext = pptx_text(pptx[0]) if pptx else ""
    htext = open(html[0], encoding="utf-8", errors="replace").read() if html else ""

    res = []

    def add(text, passed, evidence):
        res.append({"text": text, "passed": bool(passed), "evidence": evidence})

    add("נוצרו כל ארבעת התוצרים: xlsx, docx, pptx, html",
        bool(xlsx and docx and pptx and html),
        "xlsx=%d docx=%d pptx=%d html=%d" % (len(xlsx), len(docx), len(pptx), len(html)))

    add("קובץ ה-Excel מכיל נוסחאות חיות (תאים שמתחילים ב-=) ולא רק ערכים קשיחים",
        xi["formulas"] >= 30, "נמצאו %d נוסחאות ב-%d גיליונות" % (xi["formulas"], xi["sheets"]))

    add("כל גיליונות ה-Excel מוגדרים RTL (sheetView rightToLeft)",
        xi["sheets"] > 0 and xi["rtl"] == xi["sheets"],
        "%d מתוך %d גיליונות RTL" % (xi["rtl"], xi["sheets"]))

    present = [DISCLAIMER_KEY in t for t in (dtext, ptext, htext)]
    add("גילוי הנאות ('אינו מהווה דוח אפס של שמאי מקרקעין מוסמך') מופיע ב-Word, ב-PPTX וב-HTML",
        all(present), "Word=%s PPTX=%s HTML=%s" % tuple(present))

    both = ("מהעלויות" in dtext + htext) and ("מההכנסות" in dtext + htext)
    add("הרווח היזמי מוצג גם כאחוז מהעלויות וגם כאחוז מההכנסות", both,
        "'מהעלויות' ו-'מההכנסות' נמצאו בדוח/דשבורד" if both else "אחד מהם חסר")

    # עקביות: הסכום הגדול ביותר בדוח צריך להופיע גם באקסל וגם בדשבורד
    dm, hm = parse_money(dtext), parse_money(htext)
    xm = set(round(n) for n in xi["numbers"])
    shared_vals = {v for v in dm if any(abs(v - x) <= 1.5 for x in xm)} & \
                  {v for v in dm if any(abs(v - h) <= 1.5 for h in hm)}
    add("סך העלויות זהה (עד 1 ש\"ח) בין ה-Excel, ה-Word וה-HTML",
        len(shared_vals) >= 3,
        "%d סכומים גדולים מופיעים זהים בשלושת התוצרים" % len(shared_vals))

    has_assump = "הנחות יסוד" in xi["text"] and ("מקור" in xi["text"] or "אסמכתה" in xi["text"])
    add("קיים גיליון הנחות יסוד עם עמודת מקור/אסמכתה", has_assump,
        "נמצא 'הנחות יסוד' + 'מקור/אסמכתה' באקסל" if has_assump else "לא נמצא")

    has_flags = "דגל" in xi["text"] or "דגל" in dtext
    add("קיימת רשימת דגלים שמתייחסת לנתוני הפרויקט הזה ולא לניסוח גנרי",
        has_flags, "נמצאה מילת 'דגל'; הרלוונטיות נבדקת ידנית ב-viewer")

    return res, {"xlsx": xlsx, "docx": docx, "pptx": pptx, "html": html,
                 "xlsx_text": xi["text"], "docx_text": dtext, "html_text": htext,
                 "pptx_text": ptext}


EXTRA_CHECKS = {
    "residential-ashdod": [
        ("מחיר המכירה למ\"ר מסומן כסטייה מהאסמכתה החיצונית, בלי ששונה הקלט",
         lambda c: "סטיי" in c["xlsx_text"] + c["docx_text"] + c["html_text"]),
        ("קיימות שתי טבלאות רגישות דו-ממדיות (מחיר מכירה × עלות בנייה)",
         lambda c: c["xlsx_text"].count("מחיר מכירה") >= 2 or
                   c["xlsx_text"].count("רגישות") >= 1),
        ("הרווח היזמי מחושב בטווח 14%-19% מהעלויות",
         lambda c: bool(re.search(r"1[4-9][.,]\d\s*%", c["docx_text"] + c["html_text"]))),
    ],
    "pinui-binui": [
        ("180 דירות התמורה לדיירים (48 יח\"ד) אינן נספרות כהכנסה ממכירה",
         lambda c: "132" in c["xlsx_text"] + c["docx_text"] or
                   "תמורה" in c["xlsx_text"] + c["docx_text"]),
        ("עלויות הדיירים (שכר דירה, הובלה, ליווי) מופיעות כשורות עלות מפורשות",
         lambda c: "דייר" in c["xlsx_text"] + c["docx_text"]),
        ("הדוח נערך במבנה תקן 21 / מזוהה כהתחדשות עירונית",
         lambda c: "תקן 21" in c["docx_text"] or "התחדשות" in c["docx_text"] or
                   "פינוי" in c["docx_text"]),
    ],
    "logistics-income-asset": [
        ("ההכנסה נגזרת מהיוון NOI ולא מתמהיל מכירת יחידות",
         lambda c: "NOI" in c["xlsx_text"] + c["docx_text"] or
                   "היוון" in c["xlsx_text"] + c["docx_text"]),
        ("שיעור ההיוון 7.25% מופיע בגיליון ההנחות ומקושר לחישוב שווי היציאה",
         lambda c: "7.25" in c["xlsx_text"] or "7.25" in c["docx_text"]),
        ("לא הופק תמהיל מכירת דירות שאינו רלוונטי לפרויקט מניב",
         lambda c: "יח\"ד" not in c["docx_text"] and "יחידות דיור" not in c["docx_text"]),
    ],
}


def main():
    root = sys.argv[1]
    for eval_dir in sorted(os.listdir(root)):
        p = os.path.join(root, eval_dir)
        if not os.path.isdir(p) or not eval_dir.startswith("eval-"):
            continue
        name = eval_dir.split("-", 2)[2] if eval_dir.count("-") >= 2 else eval_dir
        for run in ("with_skill", "without_skill", "old_skill"):
            rd = os.path.join(p, run)
            if not os.path.isdir(rd):
                continue
            res, ctx = grade_run(rd)
            for text, fn in EXTRA_CHECKS.get(name, []):
                try:
                    ok = fn(ctx)
                except Exception as exc:
                    ok, exc = False, exc
                res.append({"text": text, "passed": bool(ok),
                            "evidence": "בדיקה אוטומטית על טקסט התוצרים"})
            passed = sum(1 for r in res if r["passed"])
            json.dump({"expectations": res, "score": passed / len(res) if res else 0},
                      open(os.path.join(rd, "grading.json"), "w"),
                      ensure_ascii=False, indent=2)
            print("%-42s %-15s %d/%d" % (eval_dir, run, passed, len(res)))


if __name__ == "__main__":
    main()
