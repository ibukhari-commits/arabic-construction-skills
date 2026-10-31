---
name: arabic-docx-rtl
description: "إصلاح وفحص اتجاه RTL في ملفات Word العربية المولَّدة بـ python-docx — يُستدعى إلزامياً بعد توليد أي مستند docx عربي وقبل تسليمه (عروض أسعار، عقود، خطابات، تقارير)."
---

# فرض اتجاه RTL الصحيح في مستندات Word العربية

## المشكلة ولماذا تتكرر

`python-docx` يضيف العناصر بـ `append()` أي في **نهاية** العنصر الأب، بينما
مخطط OOXML (ECMA-376) يفرض **ترتيباً محدداً** لأبناء `<w:pPr>` و`<w:rPr>`
و`<w:tblPr>` و`<w:sectPr>`. النتيجة:

| الخلل | ما يُولّد | الصحيح |
|---|---|---|
| اتجاه الفقرة | `<w:jc/>` ثم `<w:bidi/>` | `<w:bidi/>` **قبل** `spacing` و`ind` و`jc` |
| اتجاه الجدول | `<w:bidiVisual/>` بعد `tblLook` | `<w:bidiVisual/>` **أول** عنصر، قبل `tblW` و`jc` و`tblLook` |
| اتجاه القسم | `<w:sectPr>` **بلا** `<w:bidi/>` إطلاقاً | `<w:bidi/>` بعد `cols` وقبل `docGrid` |
| الفقرات الفارغة | `doc.add_paragraph()` بلا `<w:bidi/>` | كل فقرة تحمل `<w:bidi/>` |

**لماذا يمر الخطأ دون أن يُكتشف:** LibreOffice — المستخدم في تحويل PDF
للفحص البصري — يتسامح مع الترتيب المخالف فيظهر الـPDF سليماً تماماً، بينما
Microsoft Word يتجاهل العناصر المخالفة أو "يصلح" المستند، فيرى العميل النص
والجداول من اليسار لليمين. **الفحص البصري للـPDF لا يكشف هذه المشكلة
إطلاقاً** — وهذا سبب تكرارها رغم نجاح كل فحص.

## القاعدة

بعد توليد أي `.docx` عربي وقبل تسليمه:

1. شغّل `scripts/fix_rtl_docx.py` (نفس الكود المعروض أدناه) على الملف.
2. تحقق من XML (لا من الـPDF) وفق قائمة الفحص.
3. ثم أجرِ الفحص البصري للـPDF للتخطيط والتداخل مع الترويسة/التذييل — فهو
   يفحص شيئاً آخر ولا يغني عن الخطوة 1.

**لا تُعدّل ترويسة أو تذييل القالب الرسمي (letterhead)** — يقتصر الإصلاح على
`word/document.xml`، لأن الترويسة مصممة مسبقاً وقد تخلط العربية بالإنجليزية عمداً.

## المُطبِّع

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fix_rtl_docx.py — فرض اتجاه RTL صحيح في مستندات Word العربية.
الاستخدام: python3 fix_rtl_docx.py "ملف.docx" [ملف2.docx ...]"""
import sys
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ترتيب أبناء كل عنصر وفق ECMA-376. أي عنصر غير مذكور يبقى في النهاية.
ORDER = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
            "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
            "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
            "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
            "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
            "suppressOverlap", "jc", "textDirection", "textAlignment",
            "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
            "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
            "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
            "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
            "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
            "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
            "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
              "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
              "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
              "tblCellMar", "tblLook", "tblCaption", "tblDescription",
              "tblPrChange"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr",
               "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
               "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
               "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
               "sectPrChange"],
}


def reorder(parent, kind):
    order = ORDER[kind]
    kids = list(parent)
    def key(el):
        tag = el.tag.split("}")[-1]
        return order.index(tag) if tag in order else len(order)
    ordered = sorted(kids, key=key)   # sorted مستقر: المتساويات تحافظ على ترتيبها
    if ordered != kids:
        for el in ordered:
            parent.append(el)
        return True
    return False


def ensure(parent, tag, kind):
    if parent.find(qn("w:" + tag)) is None:
        parent.append(OxmlElement("w:" + tag))
        reorder(parent, kind)
        return True
    return False


def fix(path):
    doc = Document(path)
    body = doc.element.body
    st = dict(sect_bidi=0, para_bidi=0, para_order=0, run_order=0,
              tbl_bidi=0, tbl_order=0)

    # 1) اتجاه القسم — أهم بند: بدونه يعامل Word الصفحة كلها كـ LTR
    for sect in body.iter(W + "sectPr"):
        if ensure(sect, "bidi", "sectPr"):
            st["sect_bidi"] += 1

    # 2) الفقرات: ضمان <w:bidi/> وترتيب pPr وعلامة نهاية الفقرة
    for p in body.iter(W + "p"):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)
        if pPr.find(qn("w:bidi")) is None:
            pPr.append(OxmlElement("w:bidi"))
            st["para_bidi"] += 1
        if reorder(pPr, "pPr"):
            st["para_order"] += 1
        rPr = pPr.find(qn("w:rPr"))
        if rPr is not None:
            if rPr.find(qn("w:rtl")) is None:
                rPr.append(OxmlElement("w:rtl"))
            reorder(rPr, "rPr")

    # 3) الـ runs: ضمان <w:rtl/> وترتيب rPr
    for r in body.iter(W + "r"):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            continue
        if rPr.find(qn("w:rtl")) is None:
            rPr.append(OxmlElement("w:rtl"))
        if reorder(rPr, "rPr"):
            st["run_order"] += 1

    # 4) الجداول: <w:bidiVisual/> بالترتيب الصحيح
    for tbl in body.iter(W + "tbl"):
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl.insert(0, tblPr)
        if tblPr.find(qn("w:bidiVisual")) is None:
            tblPr.append(OxmlElement("w:bidiVisual"))
            st["tbl_bidi"] += 1
        if reorder(tblPr, "tblPr"):
            st["tbl_order"] += 1

    doc.save(path)
    return st


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("الاستخدام: python3 fix_rtl_docx.py ملف.docx [...]")
    for f in sys.argv[1:]:
        s = fix(f)
        print(f"{f}\n  أقسام bidi: {s['sect_bidi']} | فقرات bidi: {s['para_bidi']}"
              f" | pPr مُرتّبة: {s['para_order']} | rPr مُرتّبة: {s['run_order']}"
              f" | جداول bidiVisual: {s['tbl_bidi']} | tblPr مُرتّبة: {s['tbl_order']}")
```

## قائمة الفحص بعد الإصلاح

```bash
mkdir -p /tmp/rtlchk && cd /tmp/rtlchk && unzip -oq "/المسار/مخرجاتي.docx"
python3 - <<'EOF'
import re
d = open('word/document.xml', encoding='utf-8').read()
print("sectPr فيه bidi:", '<w:bidi/>' in d.split('<w:sectPr')[-1])
ps = re.findall(r'<w:p[ >].*?</w:p>', d, re.S)
print("فقرات بلا bidi:", sum(1 for p in ps if '<w:bidi/>' not in p), "من", len(ps))
print("كل جدول bidiVisual أولاً:",
      all(t.startswith('<w:tblPr><w:bidiVisual/>')
          for t in re.findall(r'<w:tblPr>.*?</w:tblPr>', d, re.S)))
print("jc=left (يجب ألا يوجد):", d.count('w:val="left"'))
EOF
```

المطلوب: `True` / `0 من N` / `True` / `0`. أي نتيجة أخرى تعني أن المستند سيظهر
LTR في Word — لا تسلّمه.

## عند كتابة مولّد جديد من الصفر

- مرّر المخرجات على المُطبِّع أعلاه — أبسط من محاولة ضبط الترتيب يدوياً عند
  كل إضافة.
- `<w:bidi/>` في `sectPr` مرة واحدة لكل قسم — أهم بند منفرد، وأكثر ما يُنسى.
- `<w:rtl/>` في `rPr` لكل run **و** في `rPr` الخاص بعلامة نهاية الفقرة داخل `pPr`.
- لا تستخدم `WD_ALIGN_PARAGRAPH.LEFT` في مستند عربي؛ استخدم `RIGHT` أو `JUSTIFY`.
- في الجداول بعد `bidiVisual`، العمود الأول منطقياً (`cells[0]`) يظهر على **اليمين**
  — رتّب الأعمدة على هذا الأساس (البيان/الوصف أولاً، القيمة/الرقم أخيراً).
- أضف `w:cs` في `rFonts` لضمان خط النصوص المعقدة (العربية)، و`szCs` للحجم.
