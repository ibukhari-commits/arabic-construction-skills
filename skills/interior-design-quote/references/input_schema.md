# مخطط spec.json — عرض سعر التصميم

| الحقل | النوع | ملاحظة |
|---|---|---|
| `client_name` | نص | اسم العميل (مطلوب) |
| `client_honorific` / `client_closing` | نص | زوج متطابق: `الأستاذ`/`المحترم` أو `الأستاذة`/`المحترمة` |
| `quote_ref`, `date` | نص | اختياريان، يظهران في الغلاف |
| `scope_types_text` | نص | عبارة بالخدمات، تدخل في الفقرة التمهيدية |
| `spaces_text` | نص | الفراغات (قصير ≈ 70 حرفاً) |
| `contract_duration_days` | رقم | بأيام العمل — يدخله المستخدم ولا يُحسب |
| `items` | قائمة | `[{"type": "...", "quantity": N}]` — الأنواع من `pricing.json` |
| `categories` | قائمة | بديل عن `items`: `[{"label": "...", "amount": N}]` مبالغ جاهزة |
| `vat_percent` | رقم | الافتراضي 15 |
| `payment_split_percent` | قائمة | الافتراضي `[60, 30, 10]` ومجموعها 100 |
| `payment_labels` | قائمة نصوص | وصف استحقاق كل دفعة (اختياري) |
| `modification_fee`, `modification_extra_days` | رقم | شرط التعديل الإضافي (اختياري) |
| `stages` | قائمة | `[["عنوان", "وصف"], ...]` لاستبدال مراحل العمل الافتراضية |
| `extra_terms` | قائمة نصوص | تستبدل الشروط الافتراضية |

يحسب السكربت: `subtotal = Σ amounts`، `vat = round(subtotal × vat%)`، `total = subtotal + vat`،
والدفعات = `round(total × pct)` للأولى، والأخيرة = الباقي (مجموع دقيق).
