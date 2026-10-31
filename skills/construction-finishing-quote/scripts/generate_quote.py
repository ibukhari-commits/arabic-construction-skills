#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
توليد عرض سعر أعمال التشطيبات لشركة مقاولات.

يبني كامل محتوى العرض داخل ورق رسمي (--template أو ورق محايد يُولَّد من company.json)
باستخدام python-docx، فيحافظ تلقائياً على الشعار والتذييل لأن هذه العناصر
موجودة في الترويسة/التذييل ولا يتم لمسها إطلاقاً.

التسعير هنا **مقطوع (Lump Sum)** يُدخله المالك يدوياً؛ السكريبت لا يقدّر
الأسعار، بل يحسب الضريبة والإجمالي وتفقيط المبلغ ويوزّع الدفعات حسب النسب.

الاستخدام:
    python generate_quote.py spec.json -o "عرض سعر تشطيبات - <العميل>.docx" \
        [--company company.json] [--template letterhead.docx]

راجع references/input_schema.md للمخطط الكامل.
"""
import argparse
import datetime
import json
import os

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x27, 0x2C, 0x3F)
GRAY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"
RIYAL = "ر.س"


# --------------------------------------------------------------------------
# أدوات RTL / خطوط / جداول  (منقولة من سكيل عرض العظم بعد اختبارها فعلياً؛
# لا تُبسَّط: كل تفصيلة فيها تعالج خطأ اتجاه ظهر في اختبار حقيقي)
# --------------------------------------------------------------------------

def set_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.append(OxmlElement("w:bidi"))
    pPr_rPr = pPr.find(qn("w:rPr"))
    if pPr_rPr is None:
        pPr_rPr = OxmlElement("w:rPr")
        pPr.append(pPr_rPr)
    pPr_rPr.append(OxmlElement("w:rtl"))


def set_run_fonts(run, size=11, bold=False, color=None, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:cs", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(attr), FONT)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)
    if bold:
        rPr.append(OxmlElement("w:bCs"))
    if italic:
        rPr.append(OxmlElement("w:iCs"))
    # حاسم لصحة موضع الأرقام (محايدة الاتجاه) داخل الفقرات العربية
    rPr.append(OxmlElement("w:rtl"))


def keep_with_next(paragraph):
    """يمنع انفصال العنوان عن نصه عند حدود الصفحة — عروض التشطيب طويلة
    (4-6 صفحات) وبدون هذا يظهر عنوان بند وحيداً في آخر الصفحة."""
    paragraph.paragraph_format.keep_with_next = True


ALIGN = {
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def add_p(doc, text="", size=11, bold=False, color=None, align="right",
          space_after=6, space_before=0, indent=None):
    p = doc.add_paragraph()
    set_rtl(p)
    p.alignment = ALIGN[align]
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if indent:
        p.paragraph_format.right_indent = Cm(indent)
    if text:
        set_run_fonts(p.add_run(text), size=size, bold=bold, color=color)
    return p


def add_scope_item(doc, number, title, text, size=11):
    """بند نطاق عمل: رقم + عنوان بارز ثم التوصيف الفني.
    الرقم والعنوان في Run واحد عمداً — فصلهما أو استخدام hanging indent يدفع
    الرقم لليسار في Word رغم اتجاه الفقرة RTL (خطأ ظهر في اختبار فعلي).
    """
    p = doc.add_paragraph()
    set_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.right_indent = Cm(0.3)
    set_run_fonts(p.add_run(f"{number}.  {title}"), size=size, bold=True, color=NAVY)
    keep_with_next(p)
    if text:
        p2 = doc.add_paragraph()
        set_rtl(p2)
        p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p2.paragraph_format.space_after = Pt(4)
        p2.paragraph_format.right_indent = Cm(0.9)
        set_run_fonts(p2.add_run(text), size=size - 0.5)
    return p


def add_bullet(doc, text, size=10.5, space_after=4, indent=1.3):
    p = doc.add_paragraph()
    set_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.right_indent = Cm(indent)
    set_run_fonts(p.add_run(f"–  {text}"), size=size)
    return p


def add_numbered_item(doc, number, text, size=10.5, space_after=6):
    p = doc.add_paragraph()
    set_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.right_indent = Cm(0.3)
    set_run_fonts(p.add_run(f"{number}.  {text}"), size=size)
    return p


def add_section_title(doc, text):
    add_p(doc, text, size=13, bold=True, color=NAVY, space_before=14, space_after=8)
    p = doc.paragraphs[-1]
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "272C3F")
    pBdr.append(bottom)
    pPr.append(pBdr)
    keep_with_next(p)
    return p


def add_group_title(doc, text):
    p = add_p(doc, text, size=11.5, bold=True, color=NAVY, space_before=10, space_after=4)
    keep_with_next(p)
    return p


def set_table_rtl(table):
    table._tbl.tblPr.append(OxmlElement("w:bidiVisual"))


def shade_cell(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_borders(cell, color="B9B9B9", sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def set_cell_text(cell, text, bold=False, color=None, size=10.5, align="right"):
    cell.paragraphs[0].text = ""
    p = cell.paragraphs[0]
    set_rtl(p)
    p.alignment = ALIGN[align]
    set_run_fonts(p.add_run(text), size=size, bold=bold, color=color)
    set_cell_borders(cell)
    cell.vertical_alignment = 1


def fmt_num(x, dec=0):
    return f"{x:,.{dec}f}"


def fmt_money(x, dec=2):
    return f"{fmt_num(x, dec)} {RIYAL}"


# --------------------------------------------------------------------------
# تفقيط المبلغ بالحروف (يمنع أي لبس في قيمة العقد — بند احترافي مطلوب عملياً)
# --------------------------------------------------------------------------

_ONES = ["", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة", "سبعة", "ثمانية",
         "تسعة", "عشرة", "أحد عشر", "اثنا عشر", "ثلاثة عشر", "أربعة عشر",
         "خمسة عشر", "ستة عشر", "سبعة عشر", "ثمانية عشر", "تسعة عشر"]
_TENS = ["", "", "عشرون", "ثلاثون", "أربعون", "خمسون", "ستون", "سبعون",
         "ثمانون", "تسعون"]
_HUNDREDS = ["", "مائة", "مائتان", "ثلاثمائة", "أربعمائة", "خمسمائة",
             "ستمائة", "سبعمائة", "ثمانمائة", "تسعمائة"]


def _under_thousand(n):
    parts = []
    if n >= 100:
        parts.append(_HUNDREDS[n // 100])
        n %= 100
    if n:
        if n < 20:
            parts.append(_ONES[n])
        else:
            unit, ten = n % 10, n // 10
            if unit:
                parts.append(f"{_ONES[unit]} و{_TENS[ten]}")
            else:
                parts.append(_TENS[ten])
    return " و".join(parts)


def _group_name(n, singular, dual, plural):
    if n == 1:
        return singular
    if n == 2:
        return dual
    if 3 <= n <= 10:
        return f"{_under_thousand(n)} {plural}"
    return f"{_under_thousand(n)} {singular}"


def tafqeet(amount):
    """يحوّل مبلغاً بالريال إلى نص عربي: 'فقط ... ريال سعودي لا غير'."""
    amount = round(float(amount), 2)
    riyals = int(amount)
    halalas = int(round((amount - riyals) * 100))
    if riyals == 0:
        words = "صفر"
    else:
        chunks = []
        billions, rem = divmod(riyals, 1_000_000_000)
        millions, rem = divmod(rem, 1_000_000)
        thousands, units = divmod(rem, 1000)
        if billions:
            chunks.append(_group_name(billions, "مليار", "ملياران", "مليارات"))
        if millions:
            chunks.append(_group_name(millions, "مليون", "مليونان", "ملايين"))
        if thousands:
            chunks.append(_group_name(thousands, "ألف", "ألفان", "آلاف"))
        if units:
            chunks.append(_under_thousand(units))
        words = " و".join(chunks)
    # الصيغة التجارية المتعارف عليها في العقود السعودية (بدون تنوين النصب،
    # تفادياً لصياغات نحوية ثقيلة لا تُستخدم عملياً في المستندات المالية)
    text = f"فقط {words} ريال سعودي"
    if halalas:
        text += f" و{_under_thousand(halalas)} هللة"
    return text + " لا غير."


# --------------------------------------------------------------------------
# النصوص الثابتة (قابلة للاستبدال أو الإضافة عبر spec.json)
# --------------------------------------------------------------------------

SCOPE_INTRO_DEFAULT = (
    "تلتزم {{COMPANY}} بتنفيذ كامل أعمال التشطيب "
    "الموصوفة أدناه على أساس (توريد وتركيب)، شاملةً جميع المواد والخامات، "
    "والعمالة الفنية والمساعدة، والمعدات والعدد، والسقالات، وإدارة الموقع "
    "والإشراف الفني، والنقل والمناولة، ورفع المخلفات أولاً بأول، وذلك وفقاً "
    "للمخططات المعتمدة وأصول الصناعة والكود السعودي للبناء، وعلى النحو التالي:"
)

SCOPE_START_NOTE = (
    "يبدأ نطاق هذا العقد بعد اكتمال أعمال الهيكل الإنشائي وأعمال المباني "
    "(البلوك) وتسليم الموقع نظيفاً وخالياً من العوائق."
)

DEFAULT_PAYMENTS = [
    {"pct": 30, "label": "دفعة مقدمة",
     "milestone": "عند توقيع العقد، وقبل تعبئة الموقع وتوريد المواد."},
    {"pct": 20, "label": None,
     "milestone": "عند الانتهاء من أعمال تمديدات السباكة والكهرباء والتكييف "
                  "وأعمال اللياسة (البياض)."},
    {"pct": 20, "label": None,
     "milestone": "عند الانتهاء من أعمال الأرضيات والتكسيات والأسقف المستعارة."},
    {"pct": 20, "label": None,
     "milestone": "عند الانتهاء من أعمال الدهانات والأبواب والشبابيك."},
    {"pct": 10, "label": "دفعة احتجاز نهائية",
     "milestone": "عند إتمام أعمال التشطيب النهائية وتركيب الأدوات الصحية "
                  "والكهربائية، واجتياز الاختبارات، والتنظيف النهائي، والتسليم "
                  "الابتدائي للعميل."},
]

# جدول دفعات بديل للمشاريع التجارية — مرّره عبر spec["payment_schedule"].
# سبب وجوده: التسلسل التنفيذي في المشاريع التجارية يختلف (أعمال ديكور وتصنيع
# مسبق للأثاث بدل الأبواب والشبابيك)، والدفعة الأخيرة ترتبط بتركيب الأثاث
# والإنارة لا بتركيب الأدوات الصحية.
COMMERCIAL_PAYMENTS = [
    {"pct": 30, "label": "دفعة مقدمة",
     "milestone": "عند توقيع العقد، وقبل تعبئة الموقع وتوريد المواد وبدء تصنيع "
                  "بنود الديكور والأثاث."},
    {"pct": 20, "label": None,
     "milestone": "عند الانتهاء من أعمال تعديل المباني، وتمديدات السباكة "
                  "والكهرباء والتكييف، وأعمال اللياسة."},
    {"pct": 20, "label": None,
     "milestone": "عند الانتهاء من أعمال العزل والأرضيات والتكسيات وهياكل "
                  "الأسقف والتجاليد الجبسية."},
    {"pct": 20, "label": None,
     "milestone": "عند الانتهاء من أعمال الدهانات وأعمال الديكور الخاصة "
                  "والكاونترات والألمنيوم والزجاج."},
    {"pct": 10, "label": "دفعة احتجاز نهائية",
     "milestone": "عند تركيب وحدات الإنارة واللوحات والستائر وتوريد وتركيب "
                  "الأثاث، واجتياز الاختبارات، والتنظيف النهائي، والتسليم "
                  "الابتدائي للعميل."},
]

DEFAULT_EXCLUSIONS = [
    "الأعمال الإنشائية والخرسانية بجميع أنواعها، وأي أعمال هدم أو تعديل في "
    "المباني غير منصوص عليها صراحةً في نطاق العمل.",
    "توريد وحدات التكييف (المكائن والوحدات الداخلية والخارجية) وشحن الفريون "
    "والتشغيل، ما لم يُنص على خلاف ذلك في نطاق العمل.",
    "توريد وتركيب المطابخ وخزائن الملابس والأثاث الثابت والمتحرك والمفروشات "
    "والستائر والأجهزة الكهربائية.",
    "المصاعد، والمولدات، وأنظمة الطاقة الشمسية.",
    "أعمال تنسيق الحدائق واللاندسكيب والمسابح والمظلات والإنترلوك والأسوار "
    "الخارجية.",
    "أنظمة التيار المنخفض (المراقبة، الإنذار، الصوتيات، الشبكات) ما لم تُدرج "
    "ضمن نطاق العمل.",
    "الرسوم الحكومية، وتراخيص البلدية، وعدادات ورسوم توصيل المرافق (الكهرباء، "
    "المياه، الصرف الصحي، الاتصالات).",
    "أعمال التصميم الداخلي وإعداد المخططات التنفيذية، ما لم يُتفق عليها بعقد "
    "مستقل.",
]

DEFAULT_CLIENT_RESPONSIBILITIES = [
    "تسليم الموقع للمقاول بعد اكتمال الهيكل الإنشائي وأعمال المباني، خالياً من "
    "العوائق والمخلفات.",
    "توفير مصدر كهرباء ومياه صالحين للاستخدام في الموقع طوال فترة التنفيذ.",
    "اعتماد العينات والخامات والألوان خلال مدة لا تتجاوز (3) أيام عمل من تاريخ "
    "عرضها، حيث يؤثر التأخر في الاعتماد مباشرةً على البرنامج الزمني.",
    "تسليم المخططات المعمارية والتنفيذية وتصاميم الديكور المعتمدة قبل مباشرة "
    "الأعمال، وأي تعديل عليها بعد التنفيذ يُعالَج بأمر تغيير.",
    "توفير مكان آمن داخل الموقع لتخزين المواد والعدد.",
    "سداد الدفعات في مواعيدها وفق الجدول المعتمد؛ ويترتب على أي تأخير في السداد "
    "تمديد مدة التنفيذ بمدة مماثلة.",
]

DEFAULT_TERMS = [
    "الأسعار المذكورة شاملة التوريد والتركيب والعمالة والمعدات وإدارة الموقع، "
    "وغير شاملة ضريبة القيمة المضافة ما لم يُذكر خلاف ذلك صراحةً.",
    "نطاق العرض: أُعِدّ هذا العرض بناءً على المخططات والمواصفات المقدمة حصراً، "
    "وأي بند أو عمل غير مذكور صراحةً يُعتبر مستثنى من النطاق والسعر.",
    "أوامر التغيير: أي أعمال إضافية أو تعديل يطلبه العميل بعد اعتماد نطاق العمل "
    "يتم توثيقه عبر ملحق (أمر تغيير) معتمد من الطرفين، وتُحدد تكلفته ومدته "
    "الزمنية بشكل منفصل قبل التنفيذ.",
    "العينات والخامات: الماركات المذكورة في نطاق العمل أو ما يماثلها من حيث "
    "الجودة والمواصفة الفنية؛ وأي ترقية في مستوى الخامة يرغبها العميل تُسعَّر "
    "بأمر تغيير مستقل.",
    "البنود غير المحددة: بالنسبة للأعمال أو الخامات غير المحددة كلياً أو جزئياً، "
    "تحدد الشركة المواصفة بما يتناسب مع المستوى العام للمشروع وتُرفع للعميل "
    "لاعتمادها قبل التوريد.",
    "حل التعارض: عند وجود أي تعارض بين وثيقة العقد والمواصفات من جهة والمخططات "
    "من جهة أخرى، تكون وثيقة العقد والمواصفات هي المرجع المعتمد.",
    "الضمان: تضمن الشركة عيوب التنفيذ لمدة ({warranty}) من تاريخ التسليم "
    "الابتدائي، ولا يشمل الضمان سوء الاستخدام أو التشغيل الخاطئ أو الأعمال "
    "التي ينفذها الغير.",
    "صلاحية العرض: ({validity}) يوماً من تاريخ الإصدار، وبعدها يخضع العرض "
    "لإعادة التسعير وفقاً لأسعار المواد السائدة.",
]


# --------------------------------------------------------------------------
# البناء
# --------------------------------------------------------------------------

REQUIRED = ["quote_ref", "client_name", "project_name", "project_location",
            "project_type", "scope_groups", "contract_value"]

ORDINALS = {1: "الأولى", 2: "الثانية", 3: "الثالثة", 4: "الرابعة", 5: "الخامسة",
            6: "السادسة", 7: "السابعة", 8: "الثامنة", 9: "التاسعة", 10: "العاشرة"}


def _ordinal(n):
    return ORDINALS.get(n, f"رقم {n}")


def build(spec, template_path, output_path):
    missing = [f for f in REQUIRED if not spec.get(f)]
    if missing:
        raise SystemExit(f"حقول ناقصة في spec.json: {', '.join(missing)}")

    ptype = spec["project_type"]
    if ptype not in ("residential", "commercial"):
        raise SystemExit("project_type يجب أن يكون residential أو commercial")

    value = float(spec["contract_value"])
    vat_pct = float(spec.get("vat_percent", 15))
    vat = round(value * vat_pct / 100, 2)
    total = round(value + vat, 2)
    validity = int(spec.get("validity_days", 30))
    warranty = spec.get("warranty_text", "12 شهراً")
    date = spec.get("date") or datetime.date.today().isoformat()

    payments = spec.get("payment_schedule") or (
        COMMERCIAL_PAYMENTS if ptype == "commercial" else DEFAULT_PAYMENTS)
    pay_sum = sum(float(p["pct"]) for p in payments)
    if abs(pay_sum - 100) > 0.001:
        raise SystemExit(f"مجموع نسب الدفعات = {pay_sum}% ويجب أن يساوي 100%")

    doc = Document(template_path)
    # الترويسة الفعلية تمتد حتى 1.11 بوصة والتذييل يبدأ 1.54 بوصة قبل الأسفل،
    # فهامش القالب الافتراضي (1 بوصة) غير كافٍ. القيمة السفلية 1.85 وليست 1.65
    # (كما في سكيل العظم) لأن عروض التشطيب أطول وتنتهي صفحاتها بفقرات نصية
    # طويلة — سُجّل تداخل فعلي مع التذييل عند 1.65 في اختبار مشروع تجاري.
    doc.sections[0].top_margin = Inches(1.2)
    doc.sections[0].bottom_margin = Inches(1.85)

    # --- العنوان (باستخدام فقرات القالب الفارغة) ---
    title_p = doc.paragraphs[0]
    set_rtl(title_p)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_fonts(title_p.add_run("عرض سعر ومواصفات فنية لأعمال التشطيب"),
                  size=16, bold=True, color=NAVY)

    sub_p = doc.paragraphs[1]
    set_rtl(sub_p)
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_fonts(sub_p.add_run(spec["project_name"]), size=12, color=GRAY)
    sub_p.paragraph_format.space_after = Pt(12)
    doc.paragraphs[2].paragraph_format.space_after = Pt(4)

    # --- جدول معلومات العرض ---
    info = doc.add_table(rows=2, cols=4)
    set_table_rtl(info)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = ["المرجع", "التاريخ", "الصلاحية", "موقع المشروع"]
    values = [spec["quote_ref"], date, f"{validity} يوماً من تاريخ الإصدار",
              spec["project_location"]]
    for i, (lab, val) in enumerate(zip(labels, values)):
        shade_cell(info.rows[0].cells[i], "272C3F")
        set_cell_text(info.rows[0].cells[i], lab, bold=True, color=WHITE,
                      size=10, align="center")
        set_cell_text(info.rows[1].cells[i], val, size=10, align="center")

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # --- جدول الطرفين ---
    parties = doc.add_table(rows=2, cols=2)
    set_table_rtl(parties)
    parties.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_text(parties.rows[0].cells[1], "الطرف الأول (المقاول)", bold=True)
    set_cell_text(parties.rows[0].cells[0], "{{COMPANY}}")
    set_cell_text(parties.rows[1].cells[1], "الطرف الثاني (العميل)", bold=True)
    set_cell_text(parties.rows[1].cells[0], spec["client_name"])

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # --- أولاً: وصف المشروع ---
    add_section_title(doc, "أولاً: وصف المشروع")
    add_p(doc, spec["project_description"], size=11, align="justify", space_after=6)
    add_p(doc, SCOPE_START_NOTE, size=10.5, align="justify", color=GRAY,
          space_after=4)

    # جدول بيانات المشروع (اختياري حسب المتوفر)
    facts = []
    if spec.get("area_m2"):
        facts.append(("المساحة الإجمالية التقريبية", f"{fmt_num(float(spec['area_m2']))} م²"))
    if spec.get("floors"):
        facts.append(("مكونات المشروع", "، ".join(spec["floors"])))
    facts.append(("نوع المشروع",
                  "مشروع سكني" if ptype == "residential" else "مشروع تجاري"))
    facts.append(("أعمال تعديل المباني",
                  "مشمولة ضمن نطاق العمل (مرحلة تمهيدية)"
                  if spec.get("masonry_modification") else "غير مشمولة"))
    ftable = doc.add_table(rows=len(facts), cols=2)
    set_table_rtl(ftable)
    ftable.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k, v) in enumerate(facts):
        shade_cell(ftable.rows[i].cells[1], "EDEDED")
        set_cell_text(ftable.rows[i].cells[1], k, bold=True, size=10)
        set_cell_text(ftable.rows[i].cells[0], v, size=10)

    # --- ثانياً: نطاق العمل ---
    add_section_title(doc, "ثانياً: نطاق العمل والمواصفات الفنية")
    add_p(doc, spec.get("scope_intro") or SCOPE_INTRO_DEFAULT, size=11,
          align="justify", space_after=8)

    n_item = 1
    for group in spec["scope_groups"]:
        if group.get("group_title"):
            add_group_title(doc, group["group_title"])
        if group.get("group_note"):
            add_p(doc, group["group_note"], size=10, color=GRAY, align="justify",
                  space_after=4, indent=0.3)
        for item in group["items"]:
            add_scope_item(doc, n_item, item["title"], item.get("text", ""))
            for b in item.get("bullets", []):
                add_bullet(doc, b)
            n_item += 1

    # --- ثالثاً: قيمة العقد ---
    add_section_title(doc, "ثالثاً: قيمة العقد")
    keep_with_next(add_p(doc, "قيمة إجمالية مقطوعة (Lump Sum) لكامل نطاق العمل "
               "الموصوف أعلاه شاملة التوريد والتركيب:", size=11,
               align="justify", space_after=8))

    price = doc.add_table(rows=4, cols=2)
    set_table_rtl(price)
    price.alignment = WD_TABLE_ALIGNMENT.CENTER
    shade_cell(price.rows[0].cells[1], "272C3F")
    shade_cell(price.rows[0].cells[0], "272C3F")
    set_cell_text(price.rows[0].cells[1], "البيان", bold=True, color=WHITE,
                  size=10, align="center")
    set_cell_text(price.rows[0].cells[0], "القيمة (ريال سعودي)", bold=True,
                  color=WHITE, size=10, align="center")
    set_cell_text(price.rows[1].cells[1],
                  spec.get("price_line_label",
                           "تنفيذ كامل أعمال التشطيب الموصوفة في نطاق العمل "
                           "(توريد وتركيب)"), size=10)
    set_cell_text(price.rows[1].cells[0], fmt_num(value, 2), size=10, align="center")
    shade_cell(price.rows[2].cells[1], "EDEDED")
    shade_cell(price.rows[2].cells[0], "EDEDED")
    set_cell_text(price.rows[2].cells[1], f"ضريبة القيمة المضافة ({vat_pct:g}%)",
                  bold=True, size=10)
    set_cell_text(price.rows[2].cells[0], fmt_num(vat, 2), bold=True, size=10,
                  align="center")
    shade_cell(price.rows[3].cells[1], "272C3F")
    shade_cell(price.rows[3].cells[0], "272C3F")
    set_cell_text(price.rows[3].cells[1], "إجمالي قيمة العقد (شامل الضريبة)",
                  bold=True, color=WHITE, size=10.5)
    set_cell_text(price.rows[3].cells[0], fmt_money(total, 2), bold=True,
                  color=WHITE, size=10.5, align="center")

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    add_p(doc, tafqeet(total), size=11, bold=True, align="right", space_after=4)

    # --- رابعاً: مدة التنفيذ ---
    add_section_title(doc, "رابعاً: مدة التنفيذ")
    add_p(doc, spec.get("duration_text",
                        "تُحدد مدة التنفيذ عند توقيع العقد."),
          size=11, align="justify", space_after=4)
    add_p(doc, "تبدأ المدة من تاريخ استلام الدفعة المقدمة وتسليم الموقع، ولا "
               "تشمل فترات التوقف الناتجة عن تأخر اعتماد العينات أو المخططات أو "
               "الدفعات من طرف العميل، أو عن أوامر التغيير، أو عن الظروف القاهرة.",
          size=10.5, align="justify", color=GRAY)

    # --- خامساً: جدول الدفعات ---
    add_section_title(doc, "خامساً: جدول الدفعات")
    keep_with_next(add_p(doc, f"تُصرف قيمة العقد على ({len(payments)}) دفعات "
               f"مرتبطة بمراحل الإنجاز الفعلية في الموقع:", size=11,
               align="justify", space_after=8))
    ptable = doc.add_table(rows=len(payments) + 1, cols=3)
    set_table_rtl(ptable)
    ptable.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(["الدفعة", "مرحلة الاستحقاق", "النسبة"]):
        shade_cell(ptable.rows[0].cells[i], "272C3F")
        set_cell_text(ptable.rows[0].cells[i], h, bold=True, color=WHITE,
                      size=10, align="center")
    for i, pay in enumerate(payments, start=1):
        cells = ptable.rows[i].cells
        name = f"الدفعة {_ordinal(i)}"
        if pay.get("label"):
            name += f"\n({pay['label']})"
        set_cell_text(cells[0], name, bold=True, size=10, align="center")
        set_cell_text(cells[1], pay["milestone"], size=10)
        set_cell_text(cells[2], f"{float(pay['pct']):g}%", bold=True, size=10,
                      align="center")

    # --- سادساً: بنود خارج نطاق العمل ---
    add_section_title(doc, "سادساً: بنود خارج نطاق العمل (استثناءات)")
    keep_with_next(add_p(doc, "البنود التالية غير مشمولة في هذا العرض ولا في "
               "قيمته، ويمكن تسعيرها بعرض مستقل عند الطلب:", size=10.5,
               align="justify", space_after=6, color=GRAY))
    excl = spec.get("exclusions") or DEFAULT_EXCLUSIONS
    excl = excl + spec.get("extra_exclusions", [])
    for i, t in enumerate(excl, start=1):
        add_numbered_item(doc, i, t)

    # --- سابعاً: مسؤوليات العميل ---
    add_section_title(doc, "سابعاً: مسؤوليات العميل")
    resp = spec.get("client_responsibilities") or DEFAULT_CLIENT_RESPONSIBILITIES
    resp = resp + spec.get("extra_client_responsibilities", [])
    for i, t in enumerate(resp, start=1):
        add_numbered_item(doc, i, t)

    # --- ثامناً: الشروط والأحكام العامة ---
    add_section_title(doc, "ثامناً: الشروط والأحكام العامة")
    terms = spec.get("terms") or DEFAULT_TERMS
    terms = [t.format(warranty=warranty, validity=validity) for t in terms]
    terms = terms + spec.get("extra_terms", [])
    for i, t in enumerate(terms, start=1):
        add_numbered_item(doc, i, t)

    # --- التوقيع ---
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    add_p(doc, "إعداد:", size=11, bold=True, space_after=2)
    add_p(doc, "{{DEPARTMENT}}", size=11, space_after=2)
    add_p(doc, "{{COMPANY}}", size=11, bold=True,
          space_after=2)

    apply_company(doc)
    doc.save(output_path)

    return {
        "quote_ref": spec["quote_ref"],
        "project_type": ptype,
        "masonry_modification": bool(spec.get("masonry_modification")),
        "scope_items_count": n_item - 1,
        "contract_value_excl_vat": value,
        "vat": vat,
        "total_incl_vat": total,
        "amount_in_words": tafqeet(total),
        "payment_percentages": [float(p["pct"]) for p in payments],
        "payment_check_sum": pay_sum,
        "output": output_path,
    }


# --------------------------------------------------------------------------
# بيانات الشركة (تُقرأ من company.json عبر --company) — لا شيء مكتوب بالكود
# --------------------------------------------------------------------------
COMPANY = {
    "name": "اسم شركتك",
    "department": "قسم العقود والمشتريات",
    "tagline": "للمقاولات العامة",
}


def apply_company(doc):
    """يستبدل {{COMPANY}} و{{DEPARTMENT}} في كل المتن والجداول."""
    mapping = {"{{COMPANY}}": COMPANY.get("name", ""),
               "{{DEPARTMENT}}": COMPANY.get("department", "")}

    def fix(paras):
        for p in paras:
            for r in p.runs:
                for k, v in mapping.items():
                    if k in r.text:
                        r.text = r.text.replace(k, v)

    fix(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                fix(c.paragraphs)



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--template", default=None,
                    help="ورقك الرسمي (docx). إن لم يُمرَّر يُولَّد ورق محايد من company.json")
    ap.add_argument("--company", default=None,
                    help="company.json: name / department / tagline / footer_lines / primary_color")
    args = ap.parse_args()

    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)
    if args.company:
        with open(args.company, encoding="utf-8") as f:
            COMPANY.update(json.load(f))

    import os
    import sys
    import tempfile
    template = args.template
    if not template:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import make_letterhead
        template = os.path.join(tempfile.mkdtemp(), "letterhead.docx")
        make_letterhead.build(COMPANY, template)

    result = build(spec, template, args.output)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
