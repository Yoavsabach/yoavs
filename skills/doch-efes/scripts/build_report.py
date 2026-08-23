#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_report.py — בונה את הדוח הנרטיבי כקובץ Word בעברית RTL מלא.

כללי ה-RTL כאן אינם המצאה מקומית: הם מיושמים לפי hebrew-doc-studio
(references/rtl-bidi-cheatsheet.md), ושלושת התנאים שחייבים להתקיים יחד הם —
  1. בסיס פסקה RTL (``w:bidi``) במיקום הנכון בסכמת ``w:pPr``.
  2. שורה מעורבת נשארת ב-run אחד. Word מריץ את אלגוריתם ה-bidi בעצמו; פיצול
     ידני לפי כתב שובר רווחים, פיסוק, סוגריים ותאריכים.
  3. גופן וגודל גם על ``w:cs``/``w:szCs`` — עברית היא "כתב מורכב" ב-Word,
     ו-``w:ascii``/``w:sz`` לבדם לא חלים עליה.
מלכודת נוספת שמיושמת כאן: תחת ``w:bidi``, ``w:jc="right"`` מתפרש כקצה הלוגי
ומרונדר **שמאלה**. ליישור-ימין פשוט לא מגדירים ``w:jc`` כלל.

שימוש:
    python3 build_report.py project.json -o "דוח אפס.docx" [--mode takan21|short]
"""

from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from model import DISCLAIMER, full_output, load_project, num  # noqa: E402

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor
except ImportError:  # pragma: no cover
    print("חסרה תלות. התקן:  pip install python-docx", file=sys.stderr)
    raise

FONT = "David"          # קיים כמעט בכל Windows עם Office — ראה hebrew-doc-studio
SIZE = 12
NAVY = "1F3864"

# סדר הילדים החוקי של CT_PPr. w:bidi חייב לבוא לפני w:jc, אחרת Word
# ו-Google Docs מתעלמים מה-bidi והפסקה נשפכת שמאלה.
_PPR_ORDER = [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
    'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
    'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
    'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
    'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
    'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
    'sectPr', 'pPrChange',
]


def _insert_ordered(pPr, element):
    tag = element.tag.split('}')[-1]
    idx = _PPR_ORDER.index(tag)
    for child in pPr:
        ctag = child.tag.split('}')[-1]
        if ctag in _PPR_ORDER and _PPR_ORDER.index(ctag) > idx:
            child.addprevious(element)
            return
    pPr.append(element)


def rtl_para(paragraph, align="right"):
    pPr = paragraph._p.get_or_add_pPr()
    _insert_ordered(pPr, pPr.makeelement(qn('w:bidi'), {}))
    # ליישור-ימין לא מגדירים jc — קצה ההובלה של פסקת bidi הוא ימין בכל renderer.
    if align == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "justify":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return paragraph


def rtl_run(paragraph, text, bold=False, italic=False, size=SIZE, color=None, font=FONT):
    """כותב run יחיד עם כל דגלי הכתב המורכב. השורה כולה ב-run אחד במכוון."""
    run = paragraph.add_run(text)
    rPr = run._r.get_or_add_rPr()
    rPr.append(rPr.makeelement(qn('w:rFonts'), {
        qn('w:ascii'): font, qn('w:hAnsi'): font, qn('w:cs'): font}))
    if bold:
        rPr.append(rPr.makeelement(qn('w:b'), {}))
        rPr.append(rPr.makeelement(qn('w:bCs'), {}))
    if italic:
        rPr.append(rPr.makeelement(qn('w:i'), {}))
        rPr.append(rPr.makeelement(qn('w:iCs'), {}))
    if color:
        rPr.append(rPr.makeelement(qn('w:color'), {qn('w:val'): color}))
    rPr.append(rPr.makeelement(qn('w:sz'), {qn('w:val'): str(int(size * 2))}))
    rPr.append(rPr.makeelement(qn('w:szCs'), {qn('w:val'): str(int(size * 2))}))
    rPr.append(rPr.makeelement(qn('w:rtl'), {}))
    return run


def money(x):
    try:
        return "%s ₪" % format(int(round(float(x))), ",")
    except (TypeError, ValueError):
        return "—"


def percent(x, digits=1):
    try:
        return ("%." + str(digits) + "f%%") % (float(x) * 100)
    except (TypeError, ValueError):
        return "—"


class Report:
    def __init__(self, data, mode="takan21"):
        self.d = data
        self.m = data["model"]
        self.mode = mode
        self.kind = data["model"].get("kind", "sale")
        self._ch = 0
        self._sub = 0
        self.doc = Document()
        self._setup()

    def _setup(self):
        sec = self.doc.sections[0]
        sec.left_margin = sec.right_margin = Mm(25)
        sec.top_margin = sec.bottom_margin = Mm(22)
        # כיוון הסקשן — בלעדיו העמוד עצמו LTR והכותרות העליונות מתהפכות.
        sec._sectPr.append(sec._sectPr.makeelement(qn('w:bidi'), {}))
        style = self.doc.styles["Normal"]
        style.font.name = FONT
        style.font.size = Pt(SIZE)
        rpr = style.element.get_or_add_rPr()
        rpr.append(rpr.makeelement(qn('w:rFonts'), {
            qn('w:ascii'): FONT, qn('w:hAnsi'): FONT, qn('w:cs'): FONT}))

    # -- אבני בניין --------------------------------------------------------

    def chapter(self, text):
        """כותרת פרק ממוספרת. המספר נגזר ממונה רץ ולא נכתב בכותרת עצמה —
        גרסה קודמת קידדה אותו קשיח, וכשמצב מקוצר דילג על פרקים יצא דוח
        שממוספר 3, 5, 7, 9. מסמך שמוגש ליזם או לבנק לא יכול להיראות כך."""
        self._ch += 1
        self._sub = 0
        return self.h("%d. %s" % (self._ch, text), 1)

    def sub(self, text):
        self._sub += 1
        return self.h("%d.%d %s" % (self._ch, self._sub, text), 2)

    def h(self, text, level=1):
        sizes = {0: 20, 1: 15, 2: 13, 3: 12}
        # סגנון Heading אמיתי ולא רק טקסט מודגש: בלעדיו אין חלונית ניווט
        # ב-Word ואי אפשר לייצר תוכן עניינים — בדוח בן 13 פרקים זה מורגש.
        style = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}.get(level)
        p = self.doc.add_paragraph(style=style) if style else self.doc.add_paragraph()
        rtl_para(p, "center" if level == 0 else "right")
        rtl_run(p, text, bold=True, size=sizes.get(level, 12), color=NAVY)
        p.paragraph_format.space_before = Pt(14 if level <= 1 else 10)
        p.paragraph_format.space_after = Pt(6)
        return p

    def para(self, text, bold=False, size=SIZE, align="justify"):
        p = self.doc.add_paragraph()
        rtl_para(p, align)
        rtl_run(p, text, bold=bold, size=size)
        p.paragraph_format.line_spacing = 1.4
        p.paragraph_format.space_after = Pt(6)
        return p

    def bullets(self, items):
        for item in items:
            p = self.doc.add_paragraph(style="List Bullet")
            rtl_para(p)
            rtl_run(p, item)
            p.paragraph_format.line_spacing = 1.35

    def table(self, headers, rows, widths=None):
        t = self.doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        # bidiVisual — בלעדיו העמודה הראשונה נופלת משמאל גם במסמך RTL.
        tblPr = t._tbl.tblPr
        tblPr.append(tblPr.makeelement(qn('w:bidiVisual'), {}))
        for i, htext in enumerate(headers):
            cell = t.rows[0].cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            rtl_para(p, "center")
            rtl_run(p, htext, bold=True, size=11, color="FFFFFF")
            shd = cell._tc.get_or_add_tcPr().makeelement(
                qn('w:shd'), {qn('w:val'): 'clear', qn('w:fill'): NAVY})
            cell._tc.get_or_add_tcPr().append(shd)
        for row in rows:
            cells = t.add_row().cells
            for i, value in enumerate(row):
                cells[i].text = ""
                p = cells[i].paragraphs[0]
                rtl_para(p, "right" if i == 0 else "center")
                rtl_run(p, str(value), size=11)
        if widths:
            for i, w in enumerate(widths):
                for row in t.rows:
                    row.cells[i].width = Mm(w)
        self.doc.add_paragraph()
        return t

    def page_break(self):
        self.doc.add_page_break()

    # -- פרקי הדוח ---------------------------------------------------------

    def cover(self):
        meta = self.m["meta"]
        for _ in range(4):
            self.doc.add_paragraph()
        self.h("ניתוח כדאיות כלכלית — דוח אפס", 0)
        p = self.doc.add_paragraph()
        rtl_para(p, "center")
        rtl_run(p, meta.get("project_name", ""), bold=True, size=16)
        for label, key in [("מזמין הדוח", "client"), ("מיקום", "location"),
                           ("עורך הדוח", "preparer"), ("תאריך", "date")]:
            if meta.get(key):
                p = self.doc.add_paragraph()
                rtl_para(p, "center")
                rtl_run(p, "%s: %s" % (label, meta[key]), size=12)
        ident = " | ".join("%s %s" % (k, meta[v]) for k, v in
                           [("גוש", "gush"), ("חלקה", "helka"), ("מגרש", "migrash")]
                           if meta.get(v))
        if ident:
            p = self.doc.add_paragraph()
            rtl_para(p, "center")
            rtl_run(p, ident, size=12, bold=True)
        for _ in range(3):
            self.doc.add_paragraph()
        p = self.doc.add_paragraph()
        rtl_para(p, "center")
        rtl_run(p, DISCLAIMER, italic=True, size=10, color="7F7F7F")
        self.page_break()

    def executive_summary(self):
        """תמצית מנהלים בעמוד הראשון.

        היזם קורא את העמוד הזה ומחליט; שאר הדוח הוא ההנמקה. לכן המסקנה נאמרת
        כאן במפורש — כולל מסקנה שלילית — ולא נגררת לסוף.
        """
        r = self.m["results"]
        self.h("תמצית מנהלים", 1)
        meta = self.m["meta"]
        self.para("דוח זה בוחן את הכדאיות הכלכלית של %s. הבחינה נערכה על בסיס הנתונים "
                  "התכנוניים, עלויות הבנייה ומחירי המכירה המפורטים בפרקים שלהלן."
                  % (meta.get("project_name") or "הפרויקט"))

        self.table(
            ["מדד", "תוצאה"],
            [
                ["סך עלויות הפרויקט", money(r["total_cost"])],
                ["סך הכנסות צפויות (נטו ממע\"מ)", money(r["total_revenue_net"])],
                ["רווח יזמי לפני מס", money(r["profit_before_tax"])],
                ["רווח יזמי — % מהעלויות", percent(r["margin_on_cost"])],
                ["רווח יזמי — % מההכנסות", percent(r["margin_on_revenue"])],
            ] + ([
                ["תשואה על העלות", percent((self.m.get("income_metrics") or {}).get("yield_on_cost"), 2)],
                ["שיעור היוון ביציאה", percent((self.m.get("income_metrics") or {}).get("exit_cap_rate"), 2)],
                ["מרווח", "%d נק' בסיס" % round((self.m.get("income_metrics") or {}).get("spread_bps") or 0)],
            ] if r.get("threshold_basis") == "spread" else []) + [
                ["IRR שנתי על ההון העצמי",
                 percent(r["irr_annual"]) if r["irr_annual"] is not None else "לא ניתן לחישוב"],
                [self._sqm_label(), money(r["cost_per_sold_sqm"])],
                ["שיא ניצול אשראי ליווי", money(r["peak_debt"])],
            ], widths=[90, 60])

        th = r["profit_threshold"]
        im = self.m.get("income_metrics") or {}
        if r.get("threshold_basis") == "spread":
            spread = im.get("spread_bps")
            need = r.get("min_spread_bps") or 150
            if r["meets_threshold"]:
                verdict = ("התשואה על העלות %s מול שיעור היוון ביציאה %s — מרווח של "
                           "%d נקודות בסיס, מעל %d הנדרשות. הרווח לפני מס %s. בכפוף "
                           "לאימות דמי השכירות ושיעור ההיוון מול עסקאות השוואה, "
                           "הפרויקט עומד במבחן."
                           % (percent(im.get("yield_on_cost"), 2),
                              percent(im.get("exit_cap_rate"), 2),
                              round(spread or 0), round(need),
                              money(r["profit_before_tax"])))
            else:
                verdict = ("התשואה על העלות %s מול שיעור היוון ביציאה %s — מרווח של "
                           "%d נקודות בסיס בלבד, מול %d הנדרשות. בנכס מניב זהו המבחן "
                           "הקובע, ולא סף הרווח היזמי: מרווח צר פירושו שהנכס שווה "
                           "כמעט בדיוק את עלות הקמתו. במתכונת הנוכחית הפרויקט אינו "
                           "עומד במבחן."
                           % (percent(im.get("yield_on_cost"), 2),
                              percent(im.get("exit_cap_rate"), 2),
                              round(spread or 0), round(need)))
        elif r["meets_threshold"]:
            verdict = ("הפרויקט מציג רווח יזמי של %s, המהווה %s מסך העלויות — מעל סף "
                       "ה-%s הנהוג בליווי בנקאי. בכפוף לאימות ההנחות ולטיפול בדגלים "
                       "המפורטים בפרק הדגלים, הפרויקט נמצא כדאי כלכלית."
                       % (money(r["profit_before_tax"]), percent(r["margin_on_cost"]),
                          percent(th, 0)))
        else:
            verdict = ("הפרויקט מציג רווח יזמי של %s, המהווה %s מסך העלויות — מתחת לסף "
                       "ה-%s הנהוג בליווי בנקאי. במתכונת הנוכחית הפרויקט אינו עומד "
                       "בסף הכדאיות, ונדרש שינוי מהותי בפרמטרים (מחיר הקרקע, תמהיל "
                       "השטחים, מחירי המכירה או מבנה המימון) בטרם קבלת החלטת השקעה."
                       % (money(r["profit_before_tax"]), percent(r["margin_on_cost"]),
                          percent(th, 0)))
        self.h("מסקנה", 2)
        self.para(verdict, bold=True)

        reds = [f for f in self.d["flags"] if f.get("level") == "red"]
        if reds:
            self.h("דגלים אדומים מרכזיים", 2)
            self.bullets(["%s — %s" % (f["title"], f["text"]) for f in reds])
        self.page_break()

    def identification(self):
        meta = self.m["meta"]
        self.chapter("זיהוי המקרקעין ומטרת הדוח")
        self.para("מטרת הדוח היא בחינת הכדאיות הכלכלית של הפרויקט עבור היזם, לצורך "
                  "קבלת החלטת השקעה ובחינת התכנות מול גורם מממן.")
        rows = [[label, meta.get(key, "—") or "—"] for label, key in
                [("שם הפרויקט", "project_name"), ("מיקום", "location"), ("גוש", "gush"),
                 ("חלקה", "helka"), ("מגרש", "migrash"), ("מצב תכנוני", "plan"),
                 ("מזמין הדוח", "client"), ("תאריך עריכה", "date")]]
        self.table(["פרט", "תוכן"], rows, widths=[55, 95])

    def planning(self):
        self.chapter("המצב התכנוני וזכויות הבנייה")
        meta = self.m["meta"]
        if meta.get("plan"):
            self.para("המצב התכנוני החל על המקרקעין: %s." % meta["plan"])
        self.para("להלן פירוט השטחים שנלקחו בחשבון בבניית המודל, לפי שימוש:")
        rows = [[line["label"], "%s מ\"ר" % format(round(line["sqm"]), ","),
                 money(line["cost_per_sqm"]) if line["cost_per_sqm"] else "—",
                 line.get("source") or "—"]
                for line in self.m["direct"]["lines"]]
        self.table(["רכיב", "שטח / כמות", "עלות ליחידה", "מקור"], rows, widths=[52, 30, 32, 36])
        self.para("אומדני היטל ההשבחה ודמי ההיתר, ככל שנכללו, הם אומדנים בלבד. "
                  "קביעה מחייבת נעשית בשומת הוועדה המקומית או ברמ\"י, וניתנת להשגה "
                  "בפני שמאי מכריע.")

    def costs(self):
        self.chapter("אומדן העלויות")
        self.sub("קרקע ורכישה")
        self.table(["סעיף", "סכום"],
                   [[l["label"], money(l["amount"])] for l in self.m["land"]["lines"]] +
                   [["סה\"כ קרקע ורכישה", money(self.m["land"]["total"])]], widths=[95, 55])

        self.sub("עלויות בנייה ישירות")
        self.table(["רכיב", "שטח / כמות", "עלות ליחידה", "סה\"כ"],
                   [[l["label"], format(round(l["sqm"]), ","),
                     money(l["cost_per_sqm"]) if l["cost_per_sqm"] else "—",
                     money(l["amount"])] for l in self.m["direct"]["lines"]] +
                   [["סה\"כ עלויות ישירות", "", "", money(self.m["direct"]["total"])]],
                   widths=[52, 28, 32, 38])

        self.sub("עלויות עקיפות")
        rows = [[("מזה: " if l.get("group") == "tenants" else "") + l["label"],
                 percent(l["pct"]) if l["pct"] else "—", money(l["amount"])]
                for l in self.m["indirect"]["lines"]]
        self.table(["סעיף", "שיעור", "סכום"],
                   rows + [["סה\"כ עלויות עקיפות", "", money(self.m["indirect"]["total"])]],
                   widths=[80, 30, 40])
        if (self.m.get("tenants") or {}).get("lines"):
            self.para("השורות המסומנות \"מזה\" מפורטות שוב בפרק עלויות הטיפול "
                      "בדיירים. הן נספרות פעם אחת בלבד בסך העלויות.", size=11)

        self.sub("בלתי צפוי מראש (בצ\"מ)")
        c = self.m["contingency"]
        self.para("בצ\"מ נקבע בשיעור %s מבסיס של %s, ובסך %s. הטווח המקובל בליווי "
                  "בנקאי הוא 5%%–10%%; שיעור נמוך מכך נבחן על ידי הבנק המלווה."
                  % (percent(c["pct"]), money(c["base"]), money(c["amount"])))

    def takan21_part_b(self):
        """חלק ב' — התמורה לבעלי היחידות האופייניות (ס' 5.1–5.13).

        "דירה אופיינית" בתקנות היא דירה טיפוסית עד קומה רביעית, **בלי** הצמדות
        ספציפיות. חלק ב' הוא הבסיס שעליו נשענים שמאי פינוי-בינוי ושמאי הדיירים,
        ולכן המבנה חשוב לא פחות מהמספרים.
        """
        tc = self.d.get("tenant_consideration") or {}
        self.chapter("חלק ב' — התמורה לבעלי היחידות האופייניות (תקן 21, ס' 5)")
        need = "— להשלמה בידי עורך השומה"
        meta = self.m["meta"]
        self.table(["ס'", "פרט", "תוכן"], [
            ["5.1", "סוג השומה", "בדיקת התמורה לבעלי היחידות האופייניות במקבץ"],
            ["5.2", "זהות מזמין השומה", meta.get("client") or need],
            ["5.3", "מטרת השומה", "בחינת התמורה לפי חוק פינוי בינוי ותקנותיו"],
            ["5.4", "המועד הקובע", meta.get("date") or need],
            ["5.5", "מועד הביקור וזהות המבקר", need],
            ["5.6", "פרטי זיהוי היחידות", "%d יחידות אופייניות" % tc.get("units", 0)
             if tc.get("units") else need],
            ["5.7", "תיאור המקבץ והסביבה", need],
            ["5.8", "המצב התכנוני המאושר", meta.get("plan") or need],
            ["5.9", "המצב המשפטי", need],
            ["5.10", "עקרונות, גורמים ושיקולים", "ראה פרק המתודולוגיה"],
        ], widths=[14, 56, 80])

        if not tc.get("typical"):
            self.para("לא הוזנו נתוני דירות אופייניות. סעיפים 5.11–5.13 מחייבים "
                      "שומת שווי לדירה לפני ואחרי, שהיא עבודת שמאי מקרקעין מוסמך "
                      "ואינה נגזרת מהמודל הכלכלי. השלם את הבלוק "
                      "`tenant_consideration.typical_units` או צרף שומה נפרדת.",
                      bold=True)
            return

        self.sub("5.11–5.12 עקרונות התמורה והתחשיב")
        rows = []
        for t in tc["typical"]:
            rows.append([
                t["label"], str(t["count"]),
                "%s מ\"ר" % format(round(t["existing_sqm"]), ","), money(t["existing_value"]),
                "%s מ\"ר" % format(round(t["new_sqm"]), ","), money(t["new_value"]),
                money(t["extras_total"]) if t["extras_total"] else "—",
                money(t["uplift"]),
                ("%.2f" % t["ratio"]) if t["ratio"] else "—",
            ])
        self.table(["דירה אופיינית", "כמות", "שטח קיים", "שווי לפני",
                    "שטח חדש", "שווי אחרי", "הצמדות", "פער", "יחס"],
                   rows, widths=[26, 12, 18, 24, 18, 24, 18, 24, 12])

        self.sub("5.13 השומה")
        if tc.get("avg_ratio"):
            self.para("היחס המשוקלל בין שווי הדירה אחרי לשווי לפני עומד על %.2f, "
                      "והפער הכולל לכלל %d היחידות האופייניות הוא %s."
                      % (tc["avg_ratio"], tc["units"], money(tc["total_uplift"])), bold=True)
        self.para("השוויים שלעיל נמסרו כקלט ואינם נקבעו על ידי הכלי. קביעת שווי "
                  "לפני ואחרי היא עבודת שמאי מקרקעין מוסמך, ובמחלוקת — של שמאי "
                  "פינוי-בינוי שמונה לפי תקנות פינוי ובינוי (פיצויים), התשע\"א-2011.")

    def takan21_part_c(self):
        """חלק ג' — התמורה בגין דירה מסוימת (ס' 6.1–6.13).

        ההבדל מחלק ב' אינו טכני: בדירה מסוימת ההצמדות הספציפיות — גג, חצר,
        מחסן — כן נספרות, וזה בדיוק מה שדייר בודד בא לברר.
        """
        tc = self.d.get("tenant_consideration") or {}
        if not tc.get("specific"):
            return
        self.chapter("חלק ג' — התמורה בגין דירה מסוימת (תקן 21, ס' 6)")
        meta = self.m["meta"]
        need = "— להשלמה בידי עורך השומה"
        self.table(["ס'", "פרט", "תוכן"], [
            ["6.1", "סוג השומה", "בדיקת הכדאיות הכלכלית של התמורה בגין דירה מסוימת"],
            ["6.3", "מטרת השומה", "לפי תקנות פינוי ובינוי (פיצויים), התשע\"א-2011"],
            ["6.4", "המועד הקובע", meta.get("date") or need],
            ["6.5", "מועד הביקור וזהות המבקר", need],
            ["6.9", "המצב המשפטי וההצמדות", need],
        ], widths=[14, 56, 80])

        self.sub("6.11–6.12 התמורה לדירה והתחשיב")
        rows = []
        for t in tc["specific"]:
            extras = "; ".join("%s %s" % (k, money(v)) for k, v in (t["extras"] or {}).items())
            rows.append([
                t["label"],
                "%s מ\"ר" % format(round(t["existing_sqm"]), ","), money(t["existing_value"]),
                "%s מ\"ר" % format(round(t["new_sqm"]), ","), money(t["new_value"]),
                extras or "—", money(t["uplift"]),
                ("%.2f" % t["ratio"]) if t["ratio"] else "—",
            ])
        self.table(["הדירה", "שטח קיים", "שווי לפני", "שטח חדש", "שווי אחרי",
                    "הצמדות ספציפיות", "פער", "יחס"], rows,
                   widths=[26, 18, 24, 18, 24, 30, 24, 12])
        self.para("הצמדות ספציפיות נספרות בחלק ג' בלבד. בחלק ב' הן אינן חלק "
                  "מאפיון הדירה הטיפוסית, בהתאם להגדרת \"דירה אופיינית\" בתקנות.")

    def betterment_chapter(self):
        """תקן 21 ס' 4.14(ה) מחייב להציג חיוב בהיטל השבחה וגם פטור. בפינוי-בינוי
        הפטור אינו ודאי, וההפרש בין התרחישים הוא לעיתים ההבדל בין כדאי ללא."""
        sc = self.d.get("betterment_scenarios") or []
        if not sc:
            return
        self.chapter("תרחישי היטל השבחה")
        self.para("היטל ההשבחה בפינוי-בינוי תלוי בהחלטת הרשות המקומית ואינו ודאי "
                  "מראש. להלן התוצאה בכל תרחיש:")
        self.table(["תרחיש", "היטל", "רווח יזמי", "% מהעלויות", "עומד בסף"],
                   [[x["label"], money(x["amount"]), money(x["profit"]),
                     percent(x["margin"]), "כן" if x["meets"] else "לא"] for x in sc],
                   widths=[50, 30, 34, 24, 20])

    def tenant_costs(self):
        """תקן 21 ס' 4.14(ד) מטפל בעלויות הדיירים כפרק עצמאי, ובצדק: זה המספר
        הראשון שהדיירים, הרשות והשמאי שואלים עליו. פיזור השורות בתוך העקיפות
        מסתיר בדיוק את מה שצריך להיות גלוי."""
        t = self.m.get("tenants") or {}
        if not t.get("lines"):
            return
        self.chapter("עלויות הטיפול בדיירים")
        self.para("בפרויקט התחדשות עירונית העלויות הכרוכות בטיפול בדיירים הן מרכיב "
                  "עצמאי, ולא סעיף משנה בעלויות העקיפות. להלן פירוטן:")
        self.table(["סעיף", "סכום", "מקור"],
                   [[l["label"], money(l["amount"]), l.get("source") or "—"]
                    for l in t["lines"]] +
                   [["סה\"כ עלויות הטיפול בדיירים", money(t["total"]), ""]],
                   widths=[70, 45, 35])
        share = t["total"] / self.m["results"]["total_cost"] if self.m["results"]["total_cost"] else 0
        self.para("עלויות הטיפול בדיירים מהוות %s מסך עלויות הפרויקט." % percent(share))

    def finance(self):
        self.chapter("מבנה המימון")
        f = self.m["finance"]
        self.para("מקורות המימון: הון עצמי בסך %s ואשראי ליווי. שיא ניצול האשראי לאורך "
                  "חיי הפרויקט עומד על %s, וזו מסגרת האשראי הנדרשת. הריבית חושבה "
                  "בשיעור שנתי של %s על היתרה המנוצלת בפועל בכל רבעון."
                  % (money(f["equity"]), money(f["peak_debt"]), percent(f["annual_rate"])))
        self.table(["סעיף", "סכום"],
                   [[l["label"], money(l["amount"])] for l in f["lines"]] +
                   [["סה\"כ עלויות מימון", money(f["total"])]], widths=[95, 55])

    def revenue(self):
        self.chapter("אומדן ההכנסות")
        rev = self.m["revenue"]
        word = "המחירים" if self.kind == "income" else "מחירי המכירה"
        self.para("%s הוזנו %s מע\"מ. הרווח נמדד על ההכנסה נטו, שכן המע\"מ מועבר "
                  "לרשות המסים ואינו מהווה הכנסה של היזם."
                  % (word, "כוללי" if rev["prices_include_vat"] else "ללא"))
        self.table(["רכיב", "יחידות", "שטח ממוצע", "מחיר למ\"ר", "סה\"כ ברוטו", "נטו"],
                   [[l["label"], format(int(l["units"]), ","),
                     format(round(l["avg_sqm"]), ","), money(l["price_per_sqm"]),
                     money(l["gross"]), money(l["net"])] for l in rev["lines"]] +
                   [["סה\"כ", "", "", "", money(rev["gross"]), money(rev["net"])]],
                   widths=[36, 18, 22, 26, 26, 26])

    def cashflow(self):
        self.chapter("תזרים המזומנים")
        self.para("התזרים נפרס על פני %d רבעונים. ההון העצמי נספג ראשון, והריבית נצברת "
                  "על היתרה שמעבר לו בלבד." % len(self.m["cashflow"]))
        rows = [[r["quarter"], money(r["outflow"]), money(r["revenue"]),
                 money(r["interest"]), money(r["debt_balance"])]
                for r in self.m["cashflow"]]
        self.table(["רבעון", "הוצאות", "הכנסות", "ריבית", "יתרת אשראי"], rows,
                   widths=[24, 32, 32, 28, 34])

    def results(self):
        self.chapter("תוצאות המודל")
        r = self.m["results"]
        self.table(["מדד", "תוצאה"], [
            ["סך עלויות הפרויקט", money(r["total_cost"])],
            ["סך הכנסות נטו", money(r["total_revenue_net"])],
            ["רווח יזמי לפני מס", money(r["profit_before_tax"])],
            ["רווח יזמי — % מהעלויות", percent(r["margin_on_cost"])],
            ["רווח יזמי — % מההכנסות", percent(r["margin_on_revenue"])],
            ["מס על הרווח", money(r["profit_tax"])],
            ["רווח לאחר מס", money(r["profit_after_tax"])],
            ["IRR שנתי", percent(r["irr_annual"]) if r["irr_annual"] is not None else "לא ניתן לחישוב"],
            ["תשואה מצטברת על ההון (על פני %.1f שנים)" % r["project_years"],
             percent(r["roc"])],
            [self._sqm_label(), money(r["cost_per_sold_sqm"])],
        ], widths=[90, 60])

    def income_asset_metrics(self):
        """בפרויקט מניב הרווח היזמי הוא לא המבחן. המבחן הוא המרווח בין התשואה
        על העלות לשיעור ההיוון שבו הנכס יימכר — מרווח דק שורף את הפרויקט גם
        כשמדד הרווח נראה חיובי."""
        im = self.m.get("income_metrics")
        if not im:
            return
        self.chapter("מדדי הנכס המניב")
        spread = im.get("spread_bps")
        self.table(["מדד", "ערך", "משמעות"], [
            ["NOI שנתי", money(im["annual_noi"]), "הכנסה תפעולית נקייה"],
            ["תשואה על העלות (Yield on Cost)", percent(im["yield_on_cost"], 2),
             "NOI חלקי סך עלות הפרויקט"],
            ["שיעור היוון ביציאה (Exit Cap)", percent(im["exit_cap_rate"], 2),
             "השיעור שבו מהוון שווי המימוש"],
            ["מרווח", "%s נק' בסיס" % (round(spread) if spread is not None else "—"),
             "הנורמה בייזום מניב היא 150–200 נק'"],
            ["שיעור היוון לאיזון", percent(im["breakeven_cap_rate"], 2),
             "מעליו הפרויקט מפסיד"],
            ["DSCR", ("%.2f" % im["dscr"]) if im.get("dscr") else "—",
             "יחס כיסוי שירות החוב"],
            ["LTC", percent(im["ltc"]), "שיעור החוב מסך העלות"],
        ], widths=[52, 38, 60])
        if spread is not None and spread < 150:
            self.para("המרווח של %d נקודות בסיס צר מהנורמה. בפרויקט מניב זהו "
                      "הסיכון המהותי: הנכס שווה כמעט בדיוק את עלות הקמתו, "
                      "ותנודה קטנה בשיעור ההיוון או בדמי השכירות מוחקת את הרווח."
                      % round(spread), bold=True)

    def break_even_chapter(self):
        """תקן 21 ס' 4.15 מחייב נקודת איזון. גם מחוץ לתקן זה המספר שהיזם
        באמת רוצה: כמה המחיר יכול לרדת לפני שהעסקה מפסידה."""
        be = self.d.get("break_even") or {}
        if be.get("zero_profit_price_per_sqm") is None:
            return
        self.chapter("נקודות איזון")
        rows = [["מחיר מכירה לאיזון (רווח אפס)",
                 money(be["zero_profit_price_per_sqm"]),
                 "%s ממחיר הבסיס" % percent(be["zero_profit_price_delta"])]]
        if be.get("threshold_price_per_sqm"):
            rows.append(["מחיר מכירה לעמידה בסף %s" % percent(be["threshold"], 0),
                         money(be["threshold_price_per_sqm"]),
                         "%s ממחיר הבסיס" % percent(be["threshold_price_delta"])])
        self.table(["מדד", "מחיר למ\"ר", "פער ממחיר הבסיס"], rows, widths=[65, 40, 45])
        # תקן 21 ס' 4.15: כשהתכנית אינה כדאית, לקבוע באילו תנאים תהיה. בוועדת
        # התחדשות הציר הנדון הוא מספר היחידות, לא המחיר.
        if be.get("zero_profit_units") or be.get("threshold_units"):
            self.sub("מספר יחידות נדרש")
            urows = []
            if be.get("zero_profit_units"):
                urows.append(["לאיזון (רווח אפס)", "%d יח\"ד" % round(be["zero_profit_units"])])
            if be.get("threshold_units"):
                urows.append(["לעמידה בסף %s" % percent(be["threshold"], 0),
                              "%d יח\"ד" % round(be["threshold_units"])])
            self.table(["מדד", "מספר יחידות למכירה"], urows, widths=[80, 55])

    def _sqm_label(self):
        return "עלות למ\"ר בנוי" if self.kind == "income" else "עלות למ\"ר מכור"

    def sensitivity(self):
        self.chapter("ניתוחי רגישות")
        s = self.d["sensitivity"]
        self.para("הטבלה מציגה את הרווח היזמי כאחוז מהעלויות בשילובים של שינוי במחיר "
                  "המכירה ובעלות הבנייה. תא שלילי או נמוך מהסף מסמן תרחיש שבו הפרויקט "
                  "מפסיק להיות כדאי.")
        headers = ["עלות בנייה \\ מחיר מכירה"] + ["%+d%%" % p for p in s["price_steps"]]
        rows = [["%+d%%" % s["cost_steps"][i]] + [percent(v) for v in row]
                for i, row in enumerate(s["grid"])]
        self.table(headers, rows)

        self.sub("תרחישי ריבית")
        self.table(["שינוי בריבית", "ריבית שנתית", "עלות ריבית", "רווח יזמי", "% מהעלויות"],
                   [["%+.1f נק' אחוז" % sc["delta_pp"], percent(sc["rate"]),
                     money(sc["interest"]), money(sc["profit"]), percent(sc["metric"])]
                    for sc in self.d["rate_scenarios"]])

    def flags(self):
        self.chapter("דגלים ונקודות לבחינה")
        flags = self.d["flags"]
        if not flags:
            self.para("המודל לא זיהה חריגה מהספים שהוגדרו. אין בכך משום אישור לפרויקט — "
                      "הדגלים נגזרים מהמספרים בלבד ואינם מחליפים בדיקה תכנונית ומשפטית.")
        for f in flags:
            self.h("%s %s" % ("🔴" if f.get("level") == "red" else "🟡", f.get("title", "")), 3)
            self.para(f.get("text", ""))
        devs = self.d["deviations"]
        if devs:
            self.sub("סטיות מאסמכתאות חיצוניות")
            self.para("הקלטים הבאים סוטים ביותר מ-10% מהאסמכתה החיצונית שאותרה. הקלט "
                      "לא שונה — ההחלטה על המספר היא של היזם — אך הפער מדווח במפורש:")
            self.bullets([d["text"] for d in devs])

    def methodology(self):
        self.chapter("מתודולוגיה, מקורות ומגבלות")
        self.para("המודל נבנה כמודל תזרימי רבעוני. העלויות והכנסות נפרסות על ציר הזמן, "
                  "עלות המימון מחושבת על היתרה המנוצלת בפועל בכל רבעון, והרווח היזמי "
                  "נמדד הן כאחוז מהעלויות והן כאחוז מההכנסות.")
        self.sub("מקורות הנתונים")
        sources = []
        for group in ("land", "direct", "indirect"):
            for line in self.m[group]["lines"]:
                if line.get("source"):
                    sources.append("%s — %s" % (line["label"], line["source"]))
        for line in self.m["revenue"]["lines"]:
            if line.get("source"):
                sources.append("%s — %s" % (line["label"], line["source"]))
        self.bullets(sources or ["לא צוינו מקורות פרטניים בקלט. יש להשלים לפני הצגה לגורם מממן."])
        notes = self.d["project"].get("assumptions_notes") or []
        if notes:
            self.sub("הנחות איכותניות")
            self.table(["נושא", "הנחה", "מקור"],
                       [[n.get("label", ""), n.get("value", ""), n.get("source") or "—"]
                        for n in notes], widths=[40, 70, 40])
        self.sub("מגבלות")
        self.bullets([
            "הדוח מבוסס על נתונים שנמסרו על ידי המזמין ועל אסמכתאות חיצוניות שאותרו במועד עריכתו.",
            "אומדני היטל השבחה ודמי היתר הם אומדנים; הקביעה המחייבת היא של הוועדה המקומית או רמ\"י.",
            "שיעורי המס ומחירי התשומות משתנים; יש לאמת אותם מחדש בכל עדכון של הדוח.",
            "המודל אינו כולל בחינה משפטית של הזכויות במקרקעין ואינו מהווה בדיקת נאותות.",
        ])
        p = self.doc.add_paragraph()
        rtl_para(p)
        rtl_run(p, DISCLAIMER, bold=True, size=11, color="C00000")

    # -- הרכבה -------------------------------------------------------------

    def takan21_frame(self):
        """מסגרת חלק א' של תקן 21, לפי מבנה הסעיפים 4.1–4.15 שאומת מול נוסח
        התקן הרשמי (references/takan-21.md §5.1).

        מה שהסקיל יכול למלא מהמודל — ממולא. מה שהוא לא יכול (מועד הביקור,
        זהות המבקר, המצב המשפטי) מופיע ככותרת עם דרישת השלמה מפורשת, ולא
        מושמט. השמטה שקטה יוצרת דוח שנראה שלם ואינו עומד בתקן; כותרת ריקה
        מסומנת אומרת לעורך בדיוק מה חסר לפני הגשה.
        """
        self.chapter("פרטי השומה (תקן 21, ס' 4.1–4.7)")
        meta = self.m["meta"]
        need = "— להשלמה בידי עורך השומה"
        self.table(["ס'", "פרט", "תוכן"], [
            ["4.1", "סוג השומה", "בדיקה שמאית כלכלית לתכנית פינוי-בינוי"],
            ["4.2", "זהות מזמין השומה", meta.get("client") or need],
            ["4.3", "יוזם התכנית", meta.get("client") or need],
            ["4.4", "מטרת השומה", "בחינת הכדאיות הכלכלית של התכנית המוצעת"],
            ["4.5", "המועד הקובע", meta.get("date") or need],
            ["4.6", "מועד הביקור במתחם וזהות המבקר", need],
            ["4.7", "פרטי זיהוי המתחם",
             " / ".join(str(meta.get(k, "")) for k in ("gush", "helka", "migrash")) or need],
        ], widths=[14, 52, 84])
        self.para("סעיפים 4.6 ו-4.11 מחייבים ביקור בפועל ובדיקת זכויות, ואינם "
                  "ניתנים לגזירה מהמודל הכלכלי. יש להשלימם בטרם הגשה.", bold=True)

        self.sub("מפת הסעיפים 4.8–4.15 אל פרקי הדוח")
        self.para("התקן מגדיר מבנה סעיפים; הדוח שלהלן ממלא אותו. הטבלה מראה היכן "
                  "כל סעיף נענה, ומה נותר להשלמה ידנית.")
        self.table(["ס'", "הסעיף בתקן", "היכן בדוח"], [
            ["4.8", "תיאור המתחם והסביבה", "להשלמה בידי עורך השומה"],
            ["4.9", "המצב התכנוני המאושר", "פרק המצב התכנוני"],
            ["4.10", "התכנית המוצעת לפינוי בינוי", "פרק המצב התכנוני + פרוגרמת השטחים"],
            ["4.11", "המצב המשפטי (הזכויות בנכס)", "להשלמה — בדיקת נסחים וזכויות"],
            ["4.12", "עקרונות, גורמים ושיקולים", "פרק המתודולוגיה"],
            ["4.13", "עקרונות לתמורות לבעלי הדירות", "חלק ב' שלהלן"],
            ["4.14", "התחשיב (א–ט)", "פרקי העלויות, ההכנסות, המימון והתוצאות"],
            ["4.15", "השומה ונקודות האיזון", "פרק נקודות האיזון"],
        ], widths=[14, 66, 70])

        self.sub("פירוט התחשיב לפי ס' 4.14")
        m = self.m
        rows = [
            ["א", "פרוגרמה תכנונית", "פרק המצב התכנוני"],
            ["ב", "אומדן תקבולים", money(m["results"]["total_revenue_net"])],
            ["ג", "עלות הקמת הפרויקט",
             money(m["direct"]["total"] + m["indirect"]["total"] + m["contingency"]["amount"])],
            ["ד", "עלויות הטיפול בדיירים",
             money((m.get("tenants") or {}).get("total", 0)) if (m.get("tenants") or {}).get("lines")
             else "לא הוזנו — ראה בלוק tenants"],
            ["ה", "עלויות מיסוי",
             money(sum(l["amount"] for l in m["indirect"]["lines"]
                       if l.get("group") == "tenants_tax"
                       or l["label"] in ("היטל השבחה", "דמי היתר רמ\"י")))],
            ["ו", "הוצאות מימון", money(m["finance"]["total"])],
            ["ז", "רווח היזם", money(m["results"]["profit_before_tax"])],
            ["ח", "תשלומים לרכישת זכויות", money(m["land"]["total"])],
            ["ט", "הוצאות מיוחדות", "ראה פירוט העלויות העקיפות"],
        ]
        self.table(["ס'", "רכיב התחשיב", "סכום / הפניה"], rows, widths=[14, 66, 70])

    def build(self, path):
        self.cover()
        self.executive_summary()
        if self.mode == "takan21":
            self.takan21_frame()
        if self.mode == "short":
            # הגרסה היזמית המקוצרת מוותרת על פרקי הזיהוי והמצב התכנוני — אבל
            # **לא** על המימון והתזרים. גרסה קודמת השמיטה גם אותם, ובפרויקטים
            # ארוכים הריבית היא הסיפור כולו; דוח שמשמיט אותה עונה על השאלה
            # הלא נכונה.
            self.costs()
            self.tenant_costs()
            self.finance()
            self.revenue()
            self.cashflow()
            self.results()
            self.income_asset_metrics()
            self.break_even_chapter()
            self.betterment_chapter()
            self.sensitivity()
            self.flags()
            self.methodology()
        else:
            self.identification()
            self.planning()
            self.costs()
            self.tenant_costs()
            self.finance()
            self.revenue()
            self.cashflow()
            self.results()
            self.income_asset_metrics()
            self.break_even_chapter()
            self.betterment_chapter()
            if self.mode == "takan21":
                self.takan21_part_b()
                self.takan21_part_c()
            self.sensitivity()
            self.flags()
            self.methodology()
        self.doc.save(path)
        return path


def main():
    ap = argparse.ArgumentParser(description="בונה את הדוח הנרטיבי כקובץ Word")
    ap.add_argument("project")
    ap.add_argument("-o", "--output", default="דוח אפס.docx")
    ap.add_argument("--mode", choices=["takan21", "full", "short"], default=None,
                    help="מבנה הדוח. takan21 = מסגרת התקן לפינוי-בינוי; "
                         "full = כל הפרקים; short = יזמי מקוצר. "
                         "ברירת מחדל: לפי meta.mode שב-project.json")
    args = ap.parse_args()
    project = load_project(args.project)
    mode = args.mode or (project.get("meta", {}) or {}).get("mode") or "full"
    data = full_output(project)
    path = Report(data, mode=mode).build(args.output)
    print("נוצר: %s  (מצב: %s)" % (path, mode))
    return 0


if __name__ == "__main__":
    sys.exit(main())
