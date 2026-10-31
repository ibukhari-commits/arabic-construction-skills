# سكيلات عربية للمقاولات والتصميم · Arabic Construction & Design Skills

مجموعة سكيلات (Skills) جاهزة لـ Claude تُنتج مستندات **عربية احترافية بالكامل (RTL)** لشركات المقاولات والمكاتب الهندسية واستديوهات التصميم: عروض أسعار، وخطابات مخالصة، وتحليل مخططات.
A set of ready-to-use Claude Skills that generate **fully right-to-left Arabic business documents** for contractors, engineering offices and interior-design studios.

> مبنية من ممارسة فعلية في سوق المقاولات السعودي، ومنقّحة من أي بيانات خاصة. كل الأسماء والأسعار والأمثلة في هذا المستودع **توضيحية**.
> Built from real-world practice in the Saudi contracting market and scrubbed of private data. All names, prices and samples here are **illustrative**.

## السكيلات · Skills

| السكيل | ماذا يفعل | المخرج |
|---|---|---|
| [`construction-structure-quote`](skills/construction-structure-quote) | عرض سعر مرحلة العظم: نطاق العمل، منهجية القياس، ضريبة، جدول دفعات يُحسب تلقائياً | Word |
| [`construction-finishing-quote`](skills/construction-finishing-quote) | عرض سعر التشطيبات (سكني/تجاري/إعادة تأهيل): توصيف فني، تفقيط المبلغ، دفعات، استثناءات، شروط | Word |
| [`interior-design-quote`](skills/interior-design-quote) | عرض سعر خدمات التصميم (داخلي/لاندسكيب/واجهات) بتسعير من ملف أسعارك | PowerPoint |
| [`payment-settlement-letter`](skills/payment-settlement-letter) | خطاب مخالصة مالية لعقود تكلفة + نسبة مع كشف تراكمي | Word |
| [`floorplan-analyzer`](skills/floorplan-analyzer) | استخراج مساحات الغرف من صور/PDF للمخططات المعمارية العربية | جدول في المحادثة |
| [`arabic-docx-rtl`](skills/arabic-docx-rtl) | إصلاح وفحص اتجاه RTL في ملفات Word العربية المولَّدة بـ python-docx | سكربت |

### لماذا `arabic-docx-rtl` مهم؟
مكتبة `python-docx` تُنتج XML بترتيب عناصر مخالف للمواصفة؛ **LibreOffice يتسامح معه فيبدو الـPDF سليماً، بينما Microsoft Word يعرض المستند من اليسار لليمين**. السكربت يصلح الترتيب ويتحقق منه في ملف docx نفسه لا في المعاينة. مفيد لأي مشروع يولّد Word عربياً، حتى خارج هذه المجموعة.

## البدء السريع · Quick start

```bash
git clone https://github.com/ibukhari-commits/arabic-construction-skills.git
cd arabic-construction-skills/skills/construction-structure-quote
pip install python-docx
python scripts/generate_quote.py examples/spec.example.json \
    -o quote.docx --company examples/company.example.json
python scripts/fix_rtl_docx.py quote.docx
```

بدون `--template` يُولَّد ورق رسمي محايد من `company.json`. لاستخدام ورقك الحقيقي بشعارك: `--template ورقك.docx`.

### الاستخدام كسكيلات Claude
- **Claude Code:** انسخ مجلد السكيل إلى `~/.claude/skills/` (أو `.claude/skills/` داخل مشروعك).
- **Claude.ai / Cowork:** اضغط مجلد السكيل (zip) وارفعه من إعدادات السكيلات (حسب نسخة الواجهة لديك).

ثم اطلب بلغتك الطبيعية: «جهّز عرض عظم لفيلا مساحتها 700 متر ودورين وملحق».

## خصّصها لشركتك · Make it yours

1. انسخ `examples/company.example.json` إلى `company.json` وضع اسم شركتك وبيانات التواصل واللون.
2. راجع **النصوص الثابتة** (نطاق العمل، الشروط، الاستثناءات) داخل `scripts/generate_quote.py` وعدّلها لتناسب عقودك.
3. للتصميم الداخلي: ضع أسعارك في `pricing.json`؛ الأرقام في المثال توضيحية.

## تنبيهات مهمة · Disclaimers

- النصوص والشروط **نماذج استرشادية** وليست استشارة قانونية أو هندسية؛ راجعها مع مختصك قبل الإرسال لأي عميل.
- راجع دائماً الأرقام المحسوبة (الضريبة، الدفعات) قبل التسليم؛ السكربتات تطبع ملخصاً للتحقق.
- لا تضع في المستودع ملفات عملائك أو أسعارك الحقيقية؛ `.gitignore` يتجاهل `company.json` و`pricing.json` و`spec.json` والمخرجات.

---

## English

Each skill follows the same pattern: a `SKILL.md` (when/how Claude should use it), a Python generator under `scripts/`, JSON input schemas under `references/`, and runnable samples under `examples/`.

- **Structure & finishing quotes** → Word proposals with scope of work, VAT, amount-in-words (tafqeet), payment schedules that must sum to exactly 100 %, exclusions and terms.
- **Interior-design quote** → 3-slide PowerPoint built from scratch (no fragile template), pricing from your own `pricing.json`.
- **Payment settlement letter** → cost-plus-fee settlement (`fee = payment × P / (100 + P)`), itemised expenses and a cumulative statement that reuses the previous letter's table.
- **Floorplan analyzer** → room-area extraction from Arabic architectural plans with perimeter checks and irregular-shape handling.
- **Arabic DOCX RTL fixer** → enforces OOXML child ordering and `w:bidi` / `w:bidiVisual` / `w:rtl` so Word renders RTL correctly.

Requirements: Python 3.9+, `python-docx`, `python-pptx`, and LibreOffice (only for visual QA via PDF).

### Contributing
Issues and pull requests are welcome — especially new scope libraries (other trades or countries), additional languages, and sanitized real-world edge cases. Please never include client names, real prices or company letterheads.

### License
MIT © Ibrahim Bukhari
