#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_letterhead.py — يولّد ورقاً رسمياً محايداً (Letterhead) بصيغة Word.

الورق الناتج فيه ترويسة (اسم الشركة + خط)، وتذييل (العنوان وبيانات التواصل)،
وثلاث فقرات فارغة في المتن (تستخدمها سكربتات العروض كنقطة بداية للعنوان).

الاستخدام:
    python make_letterhead.py company.json -o letterhead.docx

لاستخدام ورقك الرسمي الحقيقي بدل هذا: مرّر ملفك مباشرة عبر --template في
سكربت العرض (شرطه: أن يحوي فقرتين فارغتين على الأقل في المتن، والشعار
والتذييل داخل الترويسة/التذييل لا داخل المتن).

company.json (كل الحقول اختيارية؛ الافتراضي نصوص بديلة واضحة):
{
  "name": "اسم شركتك",
  "tagline": "للمقاولات العامة",
  "footer_lines": ["العنوان: المدينة - الحي", "هاتف: 05XXXXXXXX | البريد: info@example.com"],
  "primary_color": "272C3F"
}
"""
import argparse
import json

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

FONT = "Arial"


def _rtl_run(run, size, bold=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for a in ("w:cs", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(a), FONT)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)
    rPr.append(OxmlElement("w:rtl"))


def _para(container, text, size, bold=False, color=None, align="center", border_bottom=None,
          border_top=None, first=False):
    p = container.paragraphs[0] if first else container.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.append(OxmlElement("w:bidi"))
    if border_bottom or border_top:
        pBdr = OxmlElement("w:pBdr")
        for edge, col in (("top", border_top), ("bottom", border_bottom)):
            if col:
                el = OxmlElement(f"w:{edge}")
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), "12")
                el.set(qn("w:space"), "4")
                el.set(qn("w:color"), col)
                pBdr.append(el)
        pPr.append(pBdr)
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    p.paragraph_format.space_after = Pt(2)
    if text:
        _rtl_run(p.add_run(text), size, bold, color)
    return p


def build(cfg, out):
    color = cfg.get("primary_color", "272C3F")
    name = cfg.get("name", "اسم شركتك")
    tagline = cfg.get("tagline", "")
    footer = cfg.get("footer_lines") or [
        "العنوان: المدينة – الحي – الشارع",
        "هاتف: 05XXXXXXXX  |  البريد: info@example.com",
    ]

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)  # A4
    sec.left_margin = sec.right_margin = Inches(0.9)
    sec.top_margin, sec.bottom_margin = Inches(1.2), Inches(1.65)
    sec.header_distance, sec.footer_distance = Inches(0.35), Inches(0.35)

    hdr = sec.header
    _para(hdr, name, 16, True, color, first=True)
    if tagline:
        _para(hdr, tagline, 10, False, "595959", border_bottom=color)
    else:
        hdr.paragraphs[0]._p.get_or_add_pPr()

    ftr = sec.footer
    for i, line in enumerate(footer):
        _para(ftr, line, 9, False, "595959", first=(i == 0), border_top=color if i == 0 else None)

    # الاتجاه RTL على مستوى القسم
    sectPr = sec._sectPr
    bidi = OxmlElement("w:bidi")
    docGrid = sectPr.find(qn("w:docGrid"))
    if docGrid is not None:
        docGrid.addprevious(bidi)
    else:
        sectPr.append(bidi)

    # ثلاث فقرات فارغة في المتن (يستخدمها سكربت العرض للعنوان)
    for _ in range(3):
        p = doc.add_paragraph()
        p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))

    doc.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("company", nargs="?", help="company.json (اختياري)")
    ap.add_argument("-o", "--output", default="letterhead.docx")
    a = ap.parse_args()
    cfg = {}
    if a.company:
        with open(a.company, encoding="utf-8") as f:
            cfg = json.load(f)
    build(cfg, a.output)
    print(f"تم إنشاء {a.output}")


if __name__ == "__main__":
    main()
