---
name: floorplan-analyzer
description: Analyzes architectural floor plans (images or PDFs) sent by clients to identify rooms/spaces, walls, doors, windows, stairs, and interior/exterior boundaries, then calculates net floor area per space. Use this skill whenever the user uploads or references an architectural drawing, floor plan, blueprint, "مخطط معماري", "مخطط دور", or asks to calculate room areas, "مسطحات", "مساحة الفراغات", or wants a breakdown of spaces from a plan. Trigger even if the user just says "حلل لي هذا المخطط" or uploads a plan image/PDF without further instructions — floor plan area analysis is this skill's job. Also trigger for multi-story PDF plan sets, asking the skill to process each floor separately.
---

# Floor Plan Analyzer

Reads architectural floor plans (image or PDF) and produces a structured breakdown of every space with its area, in Arabic, for use by a construction/design firm owner reviewing client drawings.

## Workflow

### 1. Ingest
- If the file is a PDF with multiple pages/floors, process **each page as a separate floor** and produce a separate text summary + table per floor. Never merge floors into one table.
- If the file is an image, treat it as a single floor.
- If resolution is low or text is hard to read, crop and zoom into sub-regions (using PIL via bash) rather than guessing — dimension numbers and room labels are usually small. Zoom into each quadrant of the plan before extracting numbers.

### 2. Establish the boundary and scale
- Identify the **outer building envelope** (outermost continuous wall line) — this separates interior floor area from exterior (roof terrace, garden, setback/court areas visible on ground floor plans).
- Look for a **printed scale** (e.g. 1:100) or a **dimension string** along the outer grid lines (numbers running along axis lines, often with grid bubbles/circles labeled with numbers or Arabic letters — these are structural grid references, not room dimensions, and should NOT be used as room sizes).
- Note total building footprint dimensions if given (e.g. overall width x depth) as a sanity check for the sum of internal areas.

### 3. Identify wall thickness
- If wall thickness is dimensioned on the drawing, use that value.
- Otherwise assume **20 cm (0.20 m)** exterior/interior wall thickness by default, and state this assumption explicitly in the output.

### 4. Identify structural/architectural elements
For every enclosed region, identify:
- **Walls**: solid black/hatched double lines forming room boundaries.
- **Doors**: small gap in a wall with a quarter-circle arc (swing) — note the doorway but don't treat it as breaking the room boundary for area purposes.
- **Windows**: shorter parallel tick marks or thin double lines within a wall segment (no arc).
- **Stairs**: parallel diagonal hatch lines or a zig-zag/arrow pattern within a rectangular shaft — treat the stair shaft as its own space, labeled "بئر السلم" or "الدرج" if unlabeled.
- **Shafts / ducts / voids**: small squares with a diagonal cross (X) pattern that are NOT stairs — these are typically light wells, AC ducts, or vertical service shafts. **Exclude these entirely from the area table** — do not list them, even as "فراغ غير محدد". Just mention their approximate location once in the summary paragraph (e.g. "لوحظ شافت/فتحة خدمات قرب X، غير مدرج في الجدول").
- **Columns**: small solid filled squares/rectangles inside or along walls — note their presence but they don't form separate spaces.

### 5. Read dimensions and room labels
- Prefer **explicit dimension numbers written directly inside or beside a space** (e.g. "4.00" and "4.80" near a room) — multiply length × width directly for that space's area.
- If only one dimension is visible for an irregular space, trace the enclosed boundary using wall centerlines/inner faces, subtracting wall thickness (from step 3), and estimate area from the traced polygon.
- Read Arabic room name labels printed inside each enclosed region exactly as written (e.g. "نوم رئيسي", "مطبخ", "صالة", "صالون رجال", "صالون نساء", "حمام", "مقلط"). Keep the label as-is even if the term is regionally specific — don't "correct" or translate it.
- If a region is enclosed but has **no visible label**, name it sequentially: "فراغ غير محدد ١", "فراغ غير محدد ٢", etc., in a consistent reading order (top-to-bottom, right-to-left, matching Arabic drawing convention).

### 6. Distinguish interior vs exterior spaces
- Interior enclosed rooms → included in the net internal area table.
- Exterior open areas on the same level (roof terraces "سطح", entry porches/verandas, gardens "حديقة", light courts "فناء", setbacks) → list separately under an "الفراغات الخارجية" section, not mixed into the interior total.
- **Mandatory perimeter check**: before finalizing, trace the entire outer boundary of the drawing on all four sides and confirm every zone between the last interior wall and the outer property/building line has been accounted for — as either a named exterior space or explicitly noted as excluded (e.g. setback with no floor area). Exterior terraces/porches are easy to miss because they often lack a bold room-name label and instead only show dimension numbers and curved/bullnose wall lines — the presence of dimensions with no room label near the outer edge is itself a signal to check for a missed exterior space, not a reason to skip it.

### 7. Calculate areas
- Area = length × width for rectangular spaces with explicit dimensions (net/clear internal dimensions, i.e. already inside the wall faces if the written dimension appears to be a clear dimension — most Gulf architectural drawings dimension the clear/net opening).
- **For irregular or non-rectangular spaces (L-shapes, notches, open halls that wrap around other rooms, etc.), never just multiply the two largest bounding dimensions.** Trace the actual enclosed outline, decompose it into rectangular sub-areas (or add/subtract a bounding rectangle minus the notches), sum them, and show the decomposition briefly in the "ملاحظات" column (e.g. "8.0×3.0 + 4.7×1.2 بعد خصم تعرج الجدار"). This applies especially to large open spaces like a central صالة/ممر that abuts irregular room boundaries — verify the boundary shape against the drawing before computing, don't assume a clean rectangle just because two dimension lines are printed.
- Round areas to 2 decimal places (م²).
- Sum interior spaces into a **net internal total**, and separately note the building's gross footprint (from step 2) for the user's own comparison — do not force these to match; a gap is expected due to wall thickness and is fine to mention.

### 8. Output format (always in Arabic, in-conversation — do not create a file unless the user explicitly asks)

Structure the response as:

1. **A short paragraph** (5-8 sentences) explaining what was understood from the drawing: building type/level (سكني، دور أرضي/أول، إلخ if inferable), overall dimensions, wall thickness assumption used, source of dimensions (مكتوبة على المخطط / مقدرة), any ambiguous elements flagged, and whether the file had multiple floors.
2. **A Markdown table** with columns: `#` | `اسم الفراغ` | `الأبعاد (م)` | `المساحة (م²)` | `ملاحظات`
   - "ملاحظات" column notes things like: تقديري، غير مسمى بالمخطط، مقسم لعدة مستطيلات، فراغ خارجي، بئر سلم، إلخ.
3. **A totals line**: إجمالي المساحات الداخلية (المسطح الداخلي), and if exterior spaces exist, إجمالي الفراغات الخارجية separately.
4. If multiple floors were found in one file, repeat 1-3 per floor with a clear floor heading (e.g. "### الدور الأرضي", "### الدور الأول").
5. End with a one-line disclaimer that this is a visual/dimensional estimate from the drawing, not a substitute for a licensed survey, and offer to adjust any assumption (wall thickness, unclear labels, etc.) if the user corrects it.

## Notes
- Never fabricate a dimension that isn't legible — say so in "ملاحظات" (e.g. "الرقم غير واضح، تقدير تقريبي") rather than inventing a precise-looking number.
- Grid axis labels (circled numbers/Arabic letters like ١، ٢، ٣ or أ، ب، ج) are structural reference grids, not room identifiers — don't confuse them with room numbering.
- If the user later uploads a clearer version or corrects a label/dimension, redo just the affected row(s), not the whole table.
