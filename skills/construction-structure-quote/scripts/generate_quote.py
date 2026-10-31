#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
توليد عرض سعر (عظم / الهيكل الخرساني) لشركة مقاولات
يبني كامل محتوى العرض داخل ورق رسمي (--template أو ورق محايد يُولَّد من company.json)
باستخدام python-docx، فيحافظ تلقائياً على الشعار والتذييل وإعدادات الصفحة
لأن هذه العناصر موجودة في الترويسة/التذييل ولا يتم لمسها إطلاقاً.

الاستخدام:
    python generate_quote.py spec.json -o "عرض سعر - <اسم العميل>.docx" \
        [--company company.json] [--template letterhead.docx]

مخطط ملف spec.json:
{
  "quote_ref": "Q-2026-003",          // رقم العرض
  "date": "2026-08-03",                    // اختياري، افتراضي تاريخ اليوم (ميلادي)
  "validity_days": 30,                     // اختياري، افتراضي 30 يوماً
  "project_name": "فيلا الأستاذ فلان",      // اسم المشروع (للعنوان واسم الملف)
  "client_name": "الأستاذ فلان الفلاني",    // اسم العميل
  "project_location": "حي النور، المدينة",      // عنوان المشروع
  "floors": ["الدور الأرضي", "الدور الأول", "الملحق العلوي"],  // الأدوار، بالترتيب
  "area_m2": 774,                          // مساحة المشروع التقديرية (م²)
  "rate_per_m2": 620,                      // سعر المتر (ريال/م²، غير شامل الضريبة)
  "vat_percent": 15                        // اختياري، افتراضي 15%
}

منطق جدول الدفعات (ثابت بغض النظر عن عدد الأدوار):
  20% مقدم + 20% أساسات + [حصة كل دور] + 10% تسليم نهائي = 100%
  حصة الأدوار مجتمعة = 50%. الدور الأول في القائمة يأخذ الباقي بعد تقريب
  بقية الأدوار كل واحد لأقرب 5% للأسفل (نفس منطق نموذج الشركة الأصلي: عند
  3 أدوار تُعطي تلقائياً 20% / 15% / 15%). ملاحظة: هذا افتراض منطقي عام
  اتُّفق عليه لأن التوزيع الفعلي في النموذج المرجعي لم يكن متساوياً — يجب
  عرض جدول الدفعات المطبوع على المالك للتأكد قبل التسليم دائماً.
"""
import argparse
import datetime
import json
import sys

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches

NAVY = RGBColor(0x27, 0x2C, 0x3F)
GRAY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"

RIYAL = "ر.س"


# --------------------------------------------------------------------------
# أدوات مساعدة عامة (RTL / خطوط / حدود / تظليل)
# --------------------------------------------------------------------------

def set_rtl(paragraph):
    """يضبط اتجاه الفقرة RTL. لا يكفي وحده — كل Run لازم يُعلَّم بـ w:rtl أيضاً
    (انظر set_run_fonts)، وإلا ستضطرب Word/LibreOffice في وضع أرقام الترقيم
    ونهايات الأسطر (هذا كان سبب ظهور الترقيم على اليسار في أول نسخة)."""
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    pPr.append(bidi)
    # علّم run الخاص بعلامة نهاية الفقرة نفسها بالـ RTL أيضاً
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
    rFonts.set(qn("w:cs"), FONT)
    rFonts.set(qn("w:ascii"), FONT)
    rFonts.set(qn("w:hAnsi"), FONT)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)
    if bold:
        bCs = OxmlElement("w:bCs")
        rPr.append(bCs)
    if italic:
        iCs = OxmlElement("w:iCs")
        rPr.append(iCs)
    # حاسم لصحة اتجاه الأرقام وعلامات الترقيم داخل الفقرات العربية: بدونه قد
    # يضع Word الأرقام (وهي محايدة الاتجاه) على يسار السطر بدل يمينه.
    rPr.append(OxmlElement("w:rtl"))


def add_p(doc_or_cell, text="", size=11, bold=False, color=None, align="right",
          space_after=6, space_before=0, indent=None):
    p = doc_or_cell.add_paragraph()
    set_rtl(p)
    p.alignment = {
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if indent:
        p.paragraph_format.right_indent = Cm(indent)
    if text:
        r = p.add_run(text)
        set_run_fonts(r, size=size, bold=bold, color=color)
    return p


def add_numbered_item(doc, number, text, size=11, bold_number=True, space_after=8):
    """بند مرقم يدوياً. مقصوداً: رقم ونص في Run واحد بدل Runين منفصلين، وبدون
    hanging indent (w:firstLine/hanging عبر python-docx مرتبطة بالجهة اليسرى
    فيزيائياً بغض النظر عن اتجاه الفقرة، فكانت تُخرج الرقم على يسار السطر بدل
    يمينه في اختبار فعلي — تفادينا المشكلة نهائياً بدل محاولة عكسها)."""
    p = doc.add_paragraph()
    set_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.right_indent = Cm(0.3)
    r = p.add_run(f"{number}.  {text}")
    set_run_fonts(r, size=size, bold=False)
    return p


def add_bullet(doc, text, size=10.5, space_after=6):
    p = doc.add_paragraph()
    set_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.right_indent = Cm(0.5)
    r = p.add_run(f"•  {text}")
    set_run_fonts(r, size=size)
    return p


def add_section_title(doc, text):
    add_p(doc, text, size=13, bold=True, color=NAVY, space_before=14, space_after=8)
    # خط سفلي بسيط تحت العنوان
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


def set_table_rtl(table):
    tblPr = table._tbl.tblPr
    bidi = OxmlElement("w:bidiVisual")
    tblPr.append(bidi)


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
    p.alignment = {
        "right": WD_ALIGN_PARAGRAPH.RIGHT, "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
    }[align]
    r = p.add_run(text)
    set_run_fonts(r, size=size, bold=bold, color=color)
    set_cell_borders(cell)
    cell.vertical_alignment = 1  # center


def fmt_num(x, dec=0):
    if dec == 0:
        return f"{x:,.0f}"
    return f"{x:,.{dec}f}"


def fmt_money(x, dec=2):
    return f"{fmt_num(x, dec)} {RIYAL}"


# --------------------------------------------------------------------------
# حساب جدول الدفعات حسب عدد الأدوار
# --------------------------------------------------------------------------

def compute_payment_shares(n_floors, pool=50.0):
    """يرجع قائمة نسب الأدوار (%) بنفس الطول، تجمع إلى 50.
    الدور الأول يأخذ الباقي؛ بقية الأدوار كل واحد لأقرب 5% للأسفل من الحصة المتساوية."""
    if n_floors <= 1:
        return [pool]
    equal_share = pool / n_floors
    rest_share = int(equal_share // 5) * 5
    if rest_share <= 0:
        rest_share = 5  # حد أدنى معقول حتى لا تصفر الحصة
    rest_total = rest_share * (n_floors - 1)
    first_share = pool - rest_total
    return [first_share] + [rest_share] * (n_floors - 1)


# --------------------------------------------------------------------------
# محتوى نطاق العمل الثابت (البنود 1-9 و11-13) والشروط العامة
# --------------------------------------------------------------------------

SCOPE_INTRO = (
    "تلتزم {{COMPANY}} بتنفيذ كافة الأعمال المتعلقة "
    "بالهيكل الخرساني (\u200f\"مرحلة العظم\"\u200f) وأعمال المباني، ويشمل السعر جميع "
    "المواد المطلوبة (بما في ذلك حديد التسليح، والخرسانة الجاهزة، ومواد البناء، "
    "والعزل)، والعمالة، والمعدات، وإدارة الموقع، وأعمال الشدة الخشبية والسقالات، "
    "وذلك وفقاً للمخططات الهندسية المعتمدة وللنطاق الفني التالي:"
)

SCOPE_ITEMS_BEFORE_FLOORS = [
    "تجهيز الموقع والحفريات: التعبئة، وتحديد المواقع، والحفر حتى الأعماق المطلوبة "
    "للأساسات، ونقل والتخلص من فائض مواد الحفر في مواقع الطمر المعتمدة من البلدية.",
    "الخرسانة العزلاء (النظافة): توريد وصب طبقة خرسانة عزلاء بسماكة 10 سم أسفل جميع الأساسات.",
    "القواعد المنفصلة (الأساسات المسلحة): تنفيذ القواعد الخرسانية المسلحة المنفصلة وفقاً "
    "للمخططات الإنشائية المعتمدة.\n\u200fتوصية هندسية: يجب أن تتوافق الأساسات بشكل تام مع "
    "تقرير فحص التربة (الجيوتقني) الخاص بالمشروع.",
    "أعمدة الرقاب: تنفيذ أعمدة الرقاب الخرسانية المسلحة حتى مستوى الميدة الأرضية.",
    "خزان المياه الأرضي وخزان الصرف الصحي: تنفيذ الخزان الأرضي وخزان الصرف الصحي من "
    "الخرسانة المسلحة، شاملاً الجدران الإنشائية والأسقف وطبقات النظافة.",
    "الميدات الأرضية: تنفيذ الميدات الخرسانية المسلحة الرابطة بين أعمدة الرقاب.",
    "العزل المائي: تطبيق طبقة عزل مائي (بيتوميني) عالية الجودة لجميع العناصر الأرضية "
    "(الأساسات، أعمدة الرقاب، الميدات، والأسطح الخارجية لخزان المياه وخزان الصرف الصحي).",
    "الردم والدمك: ردم محيط الأساسات والميدات على طبقات متتالية باستخدام مواد ردم "
    "معتمدة، شاملاً الترطيب والدمك الميكانيكي.",
    "أرضية الدور الأرضي: تنفيذ أرضية خرسانية غير مسلحة (أو خفيفة التسليح) للدور الأرضي.",
]

SCOPE_ITEMS_AFTER_FLOORS = [
    "العناصر الخرسانية الملحقة: تنفيذ أسوار المبنى الخارجية، وجدار الساتر (الستره) "
    "أعلى السطح، وبئر المصعد، وبيت الدرج.",
    "أعمال المباني (البلوك): توريد وتركيب البلوك الخرساني لجميع الجدران الداخلية "
    "الفاصلة، والجدران الخارجية المحيطة، وأسوار الملكية للمبنى بالكامل وفقاً للمخططات المعمارية.",
    "التنظيف والتسليم: التنظيف المستمر للموقع، والإزالة النهائية لمخلفات البناء، "
    "والتسليم الإنشائي الرسمي للعميل/المكتب الهندسي المشرف.",
]

MEASUREMENT_METHOD_TITLE = "منهجية القياس:"

MEASUREMENT_BULLETS_TEMPLATE = [
    "نسبة 1:1 (متر بمتر): سقف {floor_list}، والأسقف الملحقة الخارجية.",
    "نسبة 1:2 (احتساب مضاعف): تُحتسب مساحة الإسقاط الأفقي لخزان المياه الأرضي وخزان "
    "الصرف الصحي مضروبة في (2).",
    "أسوار الملكية: تُقاس طولياً (بالمتر الطولي) وتُسعّر بنفس سعر المتر المربع المتفق "
    "عليه ({rate} ريال/م.ط).",
    "الجدران المقلوبة (الستره): كل 2 متر طولي يعادل 1 متر مربع (السعر شامل ستره "
    "السطح إن وُجدت).",
    "الفتحات والفراغات: أي فتحة أو فراغ إنشائي أقل من 1 م² ومحاط بالخرسانة من جميع "
    "الجهات الأربع يُحتسب كخرسانة مصمتة (\u200f\"الفاضي مليان\"\u200f).",
]

TERMS_ITEMS = [
    "نطاق العرض: يُعد هذا العرض بناءً على جدول الكميات (BOQ) والمواصفات المقدمة "
    "حصراً. أي بنود أو أعمال غير مذكورة صراحة تُعتبر مستثناة.",
    "أوامر التغيير: أي أعمال إضافية يطلبها العميل ولم تُدرج في جدول الكميات سيتم "
    "توثيقها عبر ملحق \u200f\"أمر تغيير\"\u200f، وتُحدد تكلفتها بشكل منفصل.",
    "حل التعارض: في حال وجود أي تعارض أو تناقض بين وثيقة العقد/المواصفات "
    "والمخططات الهندسية، فإن وثيقة العقد والمواصفات هي المرجع المعتمد.",
    "البنود غير المحددة: بالنسبة للأعمال أو البنود غير المحددة كلياً أو جزئياً، "
    "ستحدد {{COMPANY}} المواصفات بما يتناسب مع المستوى العام "
    "للمشروع، وتُرفع للجهة المشرفة لاعتمادها. وفي حال رغب العميل بالترقية، يُصدر "
    "أمر تغيير بذلك.",
    "حالة تسليم الموقع: يجب تسليم الموقع ل{{COMPANY}} خالياً من "
    "أي مبانٍ أو منشآت أو عوائق قائمة. ولا تتحمل الشركة أي تكلفة هدم أو إزالة "
    "مخلفات. وفي حال رغبة العميل، يُقدَّم عرض سعر منفصل لأعمال الهدم.",
    "مرافق البنية التحتية: يتحمل العميل وحده مسؤولية تكلفة إزالة أو نقل أو إخلاء "
    "أي مرافق بنية تحتية عامة أو خاصة (كشبكات الصرف الصحي، خطوط المياه الرئيسية، "
    "أو خطوط الكهرباء أو الاتصالات).",
    "الرسوم الحكومية: يستثني هذا العرض صراحة أي رسوم حكومية، أو تراخيص بلدية، أو "
    "رسوم توصيل مرافق (مياه، كهرباء، اتصالات)، أو رسوم الدفاع المدني.",
]


# --------------------------------------------------------------------------
# البناء الرئيسي
# --------------------------------------------------------------------------

REQUIRED_FIELDS = ["quote_ref", "client_name", "project_location", "floors",
                    "area_m2", "rate_per_m2"]


def build(spec, template_path, output_path):
    missing = [f for f in REQUIRED_FIELDS if f not in spec or spec[f] in (None, "", [])]
    if missing:
        raise SystemExit(f"حقول ناقصة في spec.json: {', '.join(missing)}")

    floors = spec["floors"]
    if not isinstance(floors, list) or not floors:
        raise SystemExit("حقل floors يجب أن يكون قائمة أدوار غير فارغة")
    n = len(floors)
    area = float(spec["area_m2"])
    rate = float(spec["rate_per_m2"])
    vat_pct = float(spec.get("vat_percent", 15))
    validity = int(spec.get("validity_days", 30))
    date = spec.get("date") or datetime.date.today().isoformat()

    subtotal = round(area * rate, 2)
    vat = round(subtotal * vat_pct / 100, 2)
    total = round(subtotal + vat, 2)

    def _p(v):
        v = float(v)
        return int(v) if v.is_integer() else v

    advance_pct = _p(spec.get("advance_pct", 20))
    foundation_pct = _p(spec.get("foundation_pct", 20))
    final_pct = _p(spec.get("final_pct", 10))
    shares = compute_payment_shares(n, 100.0 - advance_pct - foundation_pct - final_pct)

    doc = Document(template_path)

    # الورق الرسمي الأصلي مضبوط بهامش 1 بوصة من كل جهة، لكن قياسنا الفعلي
    # لمحتوى الترويسة (الشعار + الشريط) يمتد حتى 1.11 بوصة من الأعلى، ومحتوى
    # التذييل (العنوان + التواصل) يبدأ من 1.54 بوصة قبل الأسفل — أي أن هامش
    # القالب الأصلي (1 بوصة) لا يكفي بالفعل ليمنع تلامس أول/آخر سطر مع
    # الترويسة/التذييل عند امتلاء الصفحة. رفعنا الهامش هنا لأقل قيمة آمنة
    # فوق هذا القياس مباشرة (وليس أكثر)، حفاظاً على أقرب شكل ممكن لتناسق
    # القالب الأصلي.
    doc.sections[0].top_margin = Inches(1.2)
    doc.sections[0].bottom_margin = Inches(1.65)

    # --- استخدام فقرات القالب الفارغة الثلاث كنقطة بداية للعنوان ---
    title_p = doc.paragraphs[0]
    set_rtl(title_p)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title_p.add_run("عرض سعر ومواصفات فنية للمقاولة")
    set_run_fonts(r, size=16, bold=True, color=NAVY)

    sub_p = doc.paragraphs[1]
    set_rtl(sub_p)
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub_p.add_run(spec.get("project_name", ""))
    set_run_fonts(r2, size=12, bold=False, color=GRAY)
    sub_p.paragraph_format.space_after = Pt(12)

    # الفقرة الثالثة الفارغة تبقى كفاصل بسيط
    doc.paragraphs[2].paragraph_format.space_after = Pt(4)

    # --- جدول معلومات العرض (المرجع / التاريخ / الصلاحية / الموقع) ---
    info_table = doc.add_table(rows=2, cols=4)
    set_table_rtl(info_table)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = ["المرجع", "التاريخ", "الصلاحية", "موقع المشروع"]
    values = [spec["quote_ref"], date, f"{validity} يوماً من تاريخ الإصدار",
              spec.get("project_location", "")]
    for i, (lab, val) in enumerate(zip(labels, values)):
        shade_cell(info_table.rows[0].cells[i], "272C3F")
        set_cell_text(info_table.rows[0].cells[i], lab, bold=True, color=WHITE,
                      size=10, align="center")
        set_cell_text(info_table.rows[1].cells[i], val, bold=False, size=10,
                      align="center")

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # --- جدول الطرفين ---
    parties = doc.add_table(rows=2, cols=2)
    set_table_rtl(parties)
    parties.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_text(parties.rows[0].cells[1], "الطرف الأول (المقاول)", bold=True, size=10.5)
    set_cell_text(parties.rows[0].cells[0], "{{COMPANY}}", size=10.5)
    set_cell_text(parties.rows[1].cells[1], "الطرف الثاني (العميل)", bold=True, size=10.5)
    set_cell_text(parties.rows[1].cells[0], spec["client_name"], size=10.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # --- أولاً: نطاق العمل ---
    add_section_title(doc, "أولاً: نطاق العمل (المواصفات الفنية)")
    add_p(doc, SCOPE_INTRO, size=11, align="justify", space_after=10)

    n_item = 1
    for text in SCOPE_ITEMS_BEFORE_FLOORS:
        add_numbered_item(doc, n_item, text)
        n_item += 1

    # البند الديناميكي: أعمدة وأسقف الأدوار (بند واحد بعناوين فرعية لكل دور)
    add_numbered_item(doc, n_item, "أعمدة وأسقف الأدوار العلوية:", space_after=4)
    n_item += 1
    for floor in floors:
        add_bullet(doc, f"{floor}: صب الأعمدة الرأسية، تليها أعمال الشدة والتسليح "
                         f"وصب سقف {floor}.", size=10.5, space_after=4)

    for text in SCOPE_ITEMS_AFTER_FLOORS:
        add_numbered_item(doc, n_item, text)
        n_item += 1

    # --- ثانياً: التقدير المالي ومنهجية القياس ---
    add_section_title(doc, "ثانياً: التقدير المالي ومنهجية القياس")
    add_p(doc, f"سعر المتر المربع: {fmt_money(rate)}/م² — غير شامل ضريبة القيمة المضافة.",
          size=11, space_after=2)
    add_p(doc, f"إجمالي المساحة التقديرية للبناء: {fmt_num(area)} م² (تقدير شامل يأخذ "
               f"في الاعتبار الأسقف الفعلية، والخزانات الأرضية المحتسبة، وأسوار "
               f"الملكية، والستره، وقابل للتعديل وفق الحصر النهائي في الموقع).",
          size=11, space_after=10, align="justify")

    price_table = doc.add_table(rows=4, cols=4)
    set_table_rtl(price_table)
    price_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["البيان", "المساحة/الكمية", "سعر الوحدة", "الإجمالي (ريال)"]
    for i, h in enumerate(headers):
        shade_cell(price_table.rows[0].cells[i], "272C3F")
        set_cell_text(price_table.rows[0].cells[i], h, bold=True, color=WHITE,
                      size=10, align="center")
    row1 = price_table.rows[1].cells
    set_cell_text(row1[0], "تنفيذ الهيكل الخرساني وأعمال المباني (شامل المواد والعمالة)",
                  size=10)
    set_cell_text(row1[1], f"{fmt_num(area)} م²", align="center", size=10)
    set_cell_text(row1[2], fmt_money(rate), align="center", size=10)
    set_cell_text(row1[3], fmt_num(subtotal, 2), align="center", size=10)

    row2 = price_table.rows[2].cells
    shade_cell(row2[0], "EDEDED")
    shade_cell(row2[1], "EDEDED")
    shade_cell(row2[2], "EDEDED")
    shade_cell(row2[3], "EDEDED")
    set_cell_text(row2[0], f"ضريبة القيمة المضافة ({vat_pct:g}%)", bold=True, size=10)
    set_cell_text(row2[1], "", size=10)
    set_cell_text(row2[2], "", size=10)
    set_cell_text(row2[3], fmt_num(vat, 2), bold=True, align="center", size=10)

    row3 = price_table.rows[3].cells
    shade_cell(row3[0], "272C3F")
    shade_cell(row3[1], "272C3F")
    shade_cell(row3[2], "272C3F")
    shade_cell(row3[3], "272C3F")
    set_cell_text(row3[0], "إجمالي قيمة العقد التقديرية (شامل الضريبة)", bold=True,
                  color=WHITE, size=10.5)
    set_cell_text(row3[1], "", size=10)
    set_cell_text(row3[2], "", size=10)
    set_cell_text(row3[3], fmt_money(total, 2), bold=True, color=WHITE, align="center",
                  size=10.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    add_p(doc, MEASUREMENT_METHOD_TITLE, size=11, bold=True, space_after=6)
    floor_list_txt = "، ".join(floors[:-1]) + ("، و" + floors[-1] if n > 1 else floors[0])
    for tpl in MEASUREMENT_BULLETS_TEMPLATE:
        text = tpl.format(floor_list=floor_list_txt, rate=fmt_num(rate))
        add_bullet(doc, text, size=10.5)

    # --- ثالثاً: جدول الدفعات ---
    add_section_title(doc, "ثالثاً: جدول الدفعات")
    add_p(doc, f"لدعم تعبئة الموقع وتوريد المواد، تُصرف قيمة العقد على "
               f"{n + 3} دفعات مرتبطة بمراحل إنجاز المشروع:",
          size=11, space_after=8, align="justify")

    installment_no = 1
    add_bullet(doc, f"الدفعة {_ordinal(installment_no)} (دفعة مقدمة): "
                     f"{advance_pct}% من قيمة العقد عند توقيع العقد، وقبل تعبئة "
                     f"الموقع وتوريد المواد.", size=11, space_after=6)
    installment_no += 1
    add_bullet(doc, f"الدفعة {_ordinal(installment_no)}: {foundation_pct}% عند "
                     f"الانتهاء من صب القواعد المنفصلة، وأعمدة الرقاب، وقاعدة/جدران "
                     f"خزان المياه الأرضي وخزان الصرف الصحي.", size=11, space_after=6)
    installment_no += 1
    for i, floor in enumerate(floors):
        extra = ""
        if i == n - 1:
            extra = "، وبئر المصعد، وأسوار الملكية"
        pct = shares[i]
        pct_txt = f"{pct:g}"
        add_bullet(doc, f"الدفعة {_ordinal(installment_no)}: {pct_txt}% عند الانتهاء "
                         f"من صب أعمدة وسقف {floor}{extra}.", size=11, space_after=6)
        installment_no += 1
    add_bullet(doc, f"الدفعة {_ordinal(installment_no)} (دفعة احتجاز نهائية): "
                     f"{final_pct}% عند الانتهاء التام من أعمال المباني (البلوك)، "
                     f"وتنظيف الموقع النهائي، والتسليم الإنشائي الرسمي.",
               size=11, space_after=6)

    # --- رابعاً: الشروط والأحكام العامة ---
    add_section_title(doc, "رابعاً: الشروط والأحكام العامة")
    for text in TERMS_ITEMS:
        add_numbered_item(doc, n_item, text)
        n_item += 1

    # --- التوقيع ---
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    add_p(doc, "إعداد:", size=11, bold=True, space_after=2)
    add_p(doc, "{{DEPARTMENT}}", size=11, space_after=2)
    add_p(doc, "{{COMPANY}}", size=11, bold=True, space_after=2)

    apply_company(doc)
    doc.save(output_path)

    return {
        "quote_ref": spec["quote_ref"],
        "area_m2": area,
        "rate_per_m2": rate,
        "subtotal": subtotal,
        "vat": vat,
        "total": total,
        "floors": floors,
        "payment_shares_percent": shares,
        "payment_check_sum": advance_pct + foundation_pct + final_pct + sum(shares),
        "output": output_path,
    }


ORDINALS = {
    1: "الأولى", 2: "الثانية", 3: "الثالثة", 4: "الرابعة", 5: "الخامسة",
    6: "السادسة", 7: "السابعة", 8: "الثامنة", 9: "التاسعة", 10: "العاشرة",
}


def _ordinal(n):
    return ORDINALS.get(n, f"رقم {n}")


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
