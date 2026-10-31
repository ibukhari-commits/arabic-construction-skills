#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_quote.py — عرض سعر خدمات التصميم (داخلي / لاندسكيب / واجهات) بصيغة PowerPoint.

يبني العرض من الصفر (3 شرائح: الغلاف، العرض المالي، المراحل والشروط) بتصميم
محايد قابل للتخصيص من company.json، وبدون أي قالب جاهز — فلا يعتمد على
مواقع أشكال ثابتة ولا يتعطل عند تغيّر عدد البنود.

الاستخدام:
    python generate_quote.py spec.json -o "عرض سعر - العميل.pptx" \
        [--company company.json] [--pricing pricing.json]

spec.json يقبل أحد شكلين لبنود التسعير:
  (أ) "categories": [{"label": "...", "amount": 24000}]        — مبالغ جاهزة
  (ب) "items": [{"type": "interior_residential", "quantity": 300}] — تُحسب من
      pricing.json (سعر الوحدة × الكمية). راجع examples/pricing.example.json.

السكربت لا يفترض أي أسعار: الأسعار كلها من pricing.json أو من المستخدم.
"""
import argparse
import json
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

FONT = "Arial"

COMPANY = {
    "name": "اسم شركتك",
    "tagline": "تصميم داخلي وعمارة",
    "contact_lines": ["هاتف: 05XXXXXXXX", "البريد: info@example.com"],
    "primary_color": "272C3F",
    "accent_color": "B08D57",
}

DEFAULT_PAYMENT_SPLIT = [60, 30, 10]
DEFAULT_STAGES = [
    ("استلام المتطلبات", "جلسة تعريفية ودراسة الموقع والمخططات وتحديد احتياجات العميل."),
    ("المفهوم التصميمي", "لوحات الفكرة والخامات والألوان والتوزيع الوظيفي للفراغات."),
    ("التصور ثلاثي الأبعاد", "مناظير واقعية لكل فراغ لمراجعتها واعتمادها."),
    ("المخططات التنفيذية", "مخططات التنفيذ والتفاصيل وجداول التشطيبات والإنارة."),
    ("التسليم", "تسليم الملفات النهائية وشرحها للمقاول أو لفريق التنفيذ."),
]


# ------------------------------------------------------------------ helpers
def rgb(hexstr):
    return RGBColor.from_string(hexstr)


def fmt_sr(x):
    return f"{x:,.0f} SR"


def style_run(run, size, bold=False, color="000000"):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    rPr = run._r.get_or_add_rPr()
    # خط النصوص المعقدة (العربية) صراحةً
    for tag in ("a:cs",):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)


def set_rtl(paragraph, align=PP_ALIGN.RIGHT):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("rtl", "1")
    paragraph.alignment = align


def add_text(slide, text, left, top, width, height, size=14, bold=False,
             color="000000", align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        set_rtl(p, align)
        run = p.add_run()
        run.text = line
        style_run(run, size, bold, color)
    return tb


def add_rect(slide, left, top, width, height, fill):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def cell_set(cell, text, size=12, bold=False, color="000000", fill=None,
             align=PP_ALIGN.RIGHT):
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    set_rtl(p, align)
    r = p.add_run()
    r.text = text
    style_run(r, size, bold, color)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Inches(0.12)
    if fill:
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(fill)


# ------------------------------------------------------------------ pricing
def categories_from_spec(spec, pricing):
    if spec.get("categories"):
        return [{"label": c["label"], "amount": round(float(c["amount"]))}
                for c in spec["categories"]]
    if not spec.get("items"):
        raise SystemExit("spec.json يحتاج categories أو items")
    if not pricing:
        raise SystemExit("استخدمت items لكن لم تمرر --pricing pricing.json")
    out = []
    for it in spec["items"]:
        rate = pricing["rates"].get(it["type"])
        if rate is None:
            raise SystemExit(f"نوع غير معرّف في pricing.json: {it['type']}")
        out.append({"label": rate["label"],
                    "amount": round(float(rate["price"]) * float(it["quantity"]))})
    return out


def split_payments(total, split):
    if abs(sum(split) - 100) > 1e-9:
        raise SystemExit(f"نسب الدفعات يجب أن تساوي 100، مجموعها {sum(split)}")
    amounts = [round(total * p / 100) for p in split[:-1]]
    amounts.append(total - sum(amounts))      # الأخيرة تمتص فرق التقريب
    return amounts


# ------------------------------------------------------------------ build
def build(spec, out_path, pricing=None):
    cats = categories_from_spec(spec, pricing)
    vat_pct = float(spec.get("vat_percent", 15))
    subtotal = sum(c["amount"] for c in cats)
    vat = round(subtotal * vat_pct / 100)
    total = subtotal + vat
    split = spec.get("payment_split_percent", DEFAULT_PAYMENT_SPLIT)
    pays = split_payments(total, split)

    prim, acc = COMPANY["primary_color"], COMPANY["accent_color"]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    W, H = prs.slide_width, prs.slide_height
    M = Inches(0.7)

    # ---------- 1) الغلاف ----------
    s1 = prs.slides.add_slide(blank)
    add_rect(s1, 0, 0, W, H, prim)
    add_rect(s1, W - Inches(0.35), 0, Inches(0.35), H, acc)
    add_text(s1, COMPANY["name"], M, Inches(0.7), W - 2 * M, Inches(0.8),
             size=26, bold=True, color="FFFFFF")
    add_text(s1, COMPANY["tagline"], M, Inches(1.4), W - 2 * M, Inches(0.5),
             size=14, color=acc)
    add_text(s1, "عرض سعر", M, Inches(2.9), W - 2 * M, Inches(1.2),
             size=54, bold=True, color="FFFFFF")
    add_text(s1, f"{spec['client_honorific']} / {spec['client_name']}", M,
             Inches(4.2), W - 2 * M, Inches(0.6), size=22, color="FFFFFF")
    meta = []
    if spec.get("quote_ref"):
        meta.append(f"المرجع: {spec['quote_ref']}")
    if spec.get("date"):
        meta.append(f"التاريخ: {spec['date']}")
    if meta:
        add_text(s1, "   |   ".join(meta), M, Inches(4.9), W - 2 * M, Inches(0.5),
                 size=14, color=acc)
    add_text(s1, COMPANY.get("contact_lines", []), M, Inches(6.3), W - 2 * M,
             Inches(0.9), size=12, color="DDDDDD")

    # ---------- 2) العرض المالي ----------
    s2 = prs.slides.add_slide(blank)
    add_rect(s2, 0, 0, W, Inches(0.25), prim)
    intro = (f"بناءً على تواصلكم، نشكركم على ثقتكم، ونتشرف بتقديم عرضنا الفني والمالي "
             f"لخدمات {spec['scope_types_text']} للفراغات التالية: "
             f"({spec['spaces_text']}).")
    add_text(s2, f"{spec['client_honorific']} / {spec['client_name']}  "
                 f"{spec['client_closing']}", M, Inches(0.5), W - 2 * M,
             Inches(0.5), size=16, bold=True, color=prim)
    add_text(s2, intro, M, Inches(1.05), W - 2 * M, Inches(0.9), size=13)

    # جدول البنود: عمود المبلغ على اليسار والبيان على اليمين (ترتيب بصري RTL)
    rows = len(cats) + 3
    top = Inches(1.95)
    row_h = Inches(0.4)
    tbl_shape = s2.shapes.add_table(rows + 1, 2, M, top, W - 2 * M, row_h * (rows + 1))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(3.2)
    tbl.columns[1].width = W - 2 * M - Inches(3.2)
    cell_set(tbl.cell(0, 0), "التكلفة", 13, True, "FFFFFF", prim, PP_ALIGN.CENTER)
    cell_set(tbl.cell(0, 1), "مجال العمل", 13, True, "FFFFFF", prim)
    for i, c in enumerate(cats, start=1):
        cell_set(tbl.cell(i, 0), fmt_sr(c["amount"]), 13, False, "000000",
                 "F5F5F5" if i % 2 == 0 else "FFFFFF", PP_ALIGN.CENTER)
        cell_set(tbl.cell(i, 1), c["label"], 13, False, "000000",
                 "F5F5F5" if i % 2 == 0 else "FFFFFF")
    r = len(cats) + 1
    cell_set(tbl.cell(r, 0), fmt_sr(subtotal), 13, True, "000000", "EDEDED", PP_ALIGN.CENTER)
    cell_set(tbl.cell(r, 1), "الإجمالي قبل الضريبة", 13, True, "000000", "EDEDED")
    cell_set(tbl.cell(r + 1, 0), fmt_sr(vat), 13, False, "000000", "FFFFFF", PP_ALIGN.CENTER)
    cell_set(tbl.cell(r + 1, 1), f"ضريبة القيمة المضافة ({vat_pct:g}%)", 13, False, "000000", "FFFFFF")
    cell_set(tbl.cell(r + 2, 0), fmt_sr(total), 14, True, "FFFFFF", acc, PP_ALIGN.CENTER)
    cell_set(tbl.cell(r + 2, 1), "الإجمالي شامل الضريبة", 14, True, "FFFFFF", acc)
    for rr in tbl.rows:
        rr.height = row_h

    # جدول الدفعات
    pay_top = top + row_h * (rows + 1) + Inches(0.3)
    add_text(s2, "جدول الدفعات", M, pay_top, W - 2 * M, Inches(0.4), size=15, bold=True, color=prim)
    labels = spec.get("payment_labels") or [
        "عند توقيع العقد", "عند اعتماد المفهوم التصميمي والتصور ثلاثي الأبعاد",
        "عند تسليم المخططات التنفيذية النهائية"]
    pt = s2.shapes.add_table(len(pays) + 1, 3, M, pay_top + Inches(0.4), W - 2 * M,
                             Inches(0.36) * (len(pays) + 1)).table
    pt.columns[0].width = Inches(2.6)
    pt.columns[1].width = Inches(1.2)
    pt.columns[2].width = W - 2 * M - Inches(3.8)
    for ci, h in enumerate(["المبلغ", "النسبة", "الاستحقاق"]):
        cell_set(pt.cell(0, ci), h, 12, True, "FFFFFF", prim,
                 PP_ALIGN.RIGHT if ci == 2 else PP_ALIGN.CENTER)
    for i, (pct, amt) in enumerate(zip(split, pays), start=1):
        lab = labels[i - 1] if i - 1 < len(labels) else f"الدفعة {i}"
        cell_set(pt.cell(i, 0), fmt_sr(amt), 12, False, "000000", None, PP_ALIGN.CENTER)
        cell_set(pt.cell(i, 1), f"{pct:g}%", 12, False, "000000", None, PP_ALIGN.CENTER)
        cell_set(pt.cell(i, 2), lab, 12, False, "000000")
    for rr in pt.rows:
        rr.height = Inches(0.36)

    # ---------- 3) المراحل والشروط ----------
    s3 = prs.slides.add_slide(blank)
    add_rect(s3, 0, 0, W, Inches(0.25), prim)
    add_text(s3, "مراحل العمل", M, Inches(0.5), W - 2 * M, Inches(0.5), size=20, bold=True, color=prim)
    stages = spec.get("stages") or DEFAULT_STAGES
    n = len(stages)
    gap = Inches(0.2)
    cw = int((W - 2 * M - gap * (n - 1)) / n)
    for i, (t, d) in enumerate(stages):
        # الشريحة الأولى على اليمين
        left = W - M - cw - i * (cw + gap)
        add_rect(s3, left, Inches(1.2), cw, Inches(0.12), acc)
        add_text(s3, f"{i + 1}. {t}", left, Inches(1.4), cw, Inches(0.5), size=13, bold=True, color=prim)
        add_text(s3, d, left, Inches(1.9), cw, Inches(1.3), size=11, color="444444")

    dur = spec.get("contract_duration_days")
    fee = spec.get("modification_fee")
    days = spec.get("modification_extra_days")
    terms = []
    if dur:
        terms.append(f"تم تقدير مدة العقد بـ ({dur}) يوم عمل من تاريخ استلام المتطلبات والمستندات كاملة.")
    if fee is not None:
        extra = f" مع إضافة مدة ({days}) أيام" if days else ""
        terms.append(f"أي تعديل إضافي بعد اعتماد المفهوم التصميمي يُحتسب برسوم ({fee:,.0f}) ريال سعودي{extra}.")
    terms += spec.get("extra_terms", [
        "العرض يشمل أعمال التصميم فقط ولا يشمل الإشراف على التنفيذ ما لم يُذكر صراحةً.",
        "الأسعار المذكورة لا تشمل أي رسوم حكومية أو رسوم جهات الاعتماد.",
        "مدة صلاحية العرض (30) يوماً من تاريخ إصداره.",
    ])
    add_text(s3, "الشروط والأحكام", M, Inches(3.6), W - 2 * M, Inches(0.5), size=18, bold=True, color=prim)
    add_text(s3, [f"•  {t}" for t in terms], M, Inches(4.2), W - 2 * M, Inches(2.8), size=13, color="333333")

    prs.save(out_path)
    return {"subtotal": subtotal, "vat": vat, "total": total,
            "payments": pays, "payments_sum_check": sum(pays) == total,
            "output": out_path}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--company", default=None)
    ap.add_argument("--pricing", default=None)
    a = ap.parse_args()
    with open(a.spec, encoding="utf-8") as f:
        spec = json.load(f)
    if a.company:
        with open(a.company, encoding="utf-8") as f:
            COMPANY.update(json.load(f))
    pricing = None
    if a.pricing:
        with open(a.pricing, encoding="utf-8") as f:
            pricing = json.load(f)
    print(json.dumps(build(spec, a.output, pricing), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.exit(main())
