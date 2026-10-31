#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_template.py — يولّد قالب خطاب المخالصة المالية (docx) محايداً.

القالب يحوي: ترويسة الشركة، بيانات الخطاب، ثلاثة جداول (الملخص المالي، بيان
المصروفات، الكشف التراكمي)، وتوقيع. generate_mukhalasa.py يملؤه بالأرقام.

الاستخدام:
    python make_template.py [company.json] -o template.docx

company.json: name / tagline / footer_lines / primary_color / signer / signer_title
"""
import argparse
import json
import os
import sys
import tempfile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_letterhead  # noqa: E402

FONT = "Arial"


def run_fmt(run, size=11, bold=False, color=None):
    make_letterhead._rtl_run(run, size, bold, color)


def para(doc, text, size=11, bold=False, color=None, align="right", after=6):
    p = doc.add_paragraph()
    p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    p.alignment = {"right": WD_ALIGN_PARAGRAPH.RIGHT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
    p.paragraph_format.space_after = Pt(after)
    if text:
        run_fmt(p.add_run(text), size, bold, color)
    return p


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for e in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{e}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "B9B9B9")
        b.append(el)
    tcPr.append(b)


def cell(c, text, bold=False, color=None, fill=None, center=False, size=10.5):
    p = c.paragraphs[0]
    p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.RIGHT
    run_fmt(p.add_run(text), size, bold, color)
    borders(c)
    if fill:
        shade(c, fill)


def table(doc, rows, cols):
    t = doc.add_table(rows=rows, cols=cols)
    t._tbl.tblPr.insert(0, OxmlElement("w:bidiVisual"))
    return t


def build(cfg, out):
    prim = cfg.get("primary_color", "272C3F")
    tmp = os.path.join(tempfile.mkdtemp(), "lh.docx")
    make_letterhead.build(cfg, tmp)
    doc = Document(tmp)
    for p in list(doc.paragraphs):           # احذف الفقرات الفارغة الثلاث
        p._p.getparent().remove(p._p)

    para(doc, "خطاب مخالصة مالية", 16, True, prim, "center", 10)
    para(doc, "التاريخ: 0000-00-00")
    para(doc, "الموضوع: خطاب مخالصة مالية – الدفعة رقم (1)", bold=True)
    para(doc, "المشروع: اسم العميل – المدينة – الحي")
    p = para(doc, "", after=8)
    for txt in ("الأستاذ", "    ", " / اسم العميل  ", "المحترم"):
        run_fmt(p.add_run(txt), 11, True)
    para(doc, "السلام عليكم ورحمة الله وبركاته،")
    para(doc, "نفيدكم بأنه تم استلام الدفعة المحددة أدناه وفق عقد (تكلفة + نسبة) المبرم "
              "بين الطرفين، وفيما يلي بيان المخالصة المالية الخاص بها:", align="justify", after=10)

    para(doc, "أولاً: الملخص المالي للدفعة", 12.5, True, prim, after=4)
    t0 = table(doc, 5, 2)
    labels = ["قيمة الدفعة المستلمة", "أتعاب إدارة المشروع (‎15%‎)",
              "المتاح للمصروفات المباشرة", "إجمالي المصروفات الفعلية", "صافي المتبقي من الدفعة"]
    for i, lab in enumerate(labels):
        cell(t0.rows[i].cells[0], lab, True, fill="EDEDED")
        cell(t0.rows[i].cells[1], "0.00 ر.س", center=True)
    para(doc, "", after=4)

    para(doc, "ثانياً: بيان المصروفات", 12.5, True, prim, after=4)
    t1 = table(doc, 3, 2)
    cell(t1.rows[0].cells[0], "البيان", True, "FFFFFF", prim)
    cell(t1.rows[0].cells[1], "المبلغ", True, "FFFFFF", prim, True)
    cell(t1.rows[1].cells[0], "بند مصروف")
    cell(t1.rows[1].cells[1], "0.00 ر.س", center=True)
    cell(t1.rows[2].cells[0], "الإجمالي", True, fill="EDEDED")
    cell(t1.rows[2].cells[1], "0.00 ر.س", True, fill="EDEDED", center=True)
    para(doc, "", after=4)

    para(doc, "ثالثاً: الكشف التراكمي لجميع الدفعات", 12.5, True, prim, after=4)
    t2 = table(doc, 2, 6)
    heads = ["الدفعة", "قيمة الدفعة", "أتعاب الشركة (‎15%‎)", "المصروفات", "الصافي", "الرصيد التراكمي"]
    for i, h in enumerate(heads):
        cell(t2.rows[0].cells[i], h, True, "FFFFFF", prim, True, 9.5)
    for i in range(6):
        cell(t2.rows[1].cells[i], "0", center=True, size=9.5)
    para(doc, "", after=2)
    para(doc, "الرصيد التراكمي المتبقي حتى تاريخه: 0.00 ر.س", bold=True, after=10)

    para(doc, "رابعاً: طلب الدفعة التالية", 12.5, True, prim, after=4)
    para(doc, "وعليه، وبعد إقفال حساب الدفعة الحالية رقم (1)، نأمل التكرم بتحويل الدفعة "
              "التالية حتى لا يتعطل سير الأعمال وفق الجدول الزمني المتفق عليه.",
         align="justify", after=10)
    para(doc, "وتفضلوا بقبول فائق الاحترام والتقدير،", after=14)
    para(doc, "{{COMPANY}}", bold=True, after=2)
    para(doc, "{{SIGNER}}", after=2)
    para(doc, "{{SIGNER_TITLE}}", after=2)
    doc.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("company", nargs="?")
    ap.add_argument("-o", "--output", default="template.docx")
    a = ap.parse_args()
    cfg = {}
    if a.company:
        with open(a.company, encoding="utf-8") as f:
            cfg = json.load(f)
    build(cfg, a.output)
    print("تم إنشاء", a.output)


if __name__ == "__main__":
    main()
