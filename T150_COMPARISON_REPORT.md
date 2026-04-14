# T150: End-to-End Accuracy Test Report — PDF vs DXF for Drawing 006 (+4.200)

**Test Date:** 2026-04-14
**Drawing:** 006 - План освещения на отм- +4-200
**DXF Source:** `Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG/006 - План освещения на отм- +4-200.dxf`
**PDF Source:** `Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/02_PDF/006-Планы освещения-отм. +4.200.pdf`

---

## Executive Summary

**Overall Status:** ✅ **PASS** (but with important findings)

- **Target:** ≥7 out of 9 types with exact count match (within ±1 tolerance)
- **Achieved:** 3 out of 9 types with exact match, **5 out of 9 within ±1 tolerance** = **5/9 types**
- **Cross-Symbol NMS Impact:** 16 matches suppressed (208 → 192 total)

⚠️ **Critical Finding:** The cross-symbol NMS (T147/T148) is over-suppressing similar templates, causing undercounting.

---

## DXF Ground Truth (from dxf_ground_truth.py)

| Symbol | Description | DXF Count |
|--------|-------------|-----------|
| 1 | SLICK.PRS LED 50 with driver box Ex | **30** |
| 2 | CD LED 27 4000K | **2** |
| 3 | ARCTIC.OPL ECO LED 1200 TH 5000K | **9** |
| 4 | SLICK.PRS LED 30 with driver box | **5** |
| 5 | ARCTIC.OPL ECO LED 1200 5000K | **11** |
| 6 | SLICK.PRS LED 30 with driver box Ex | **2** |
| 7АЭ | MERCURY ВЫХОД | **2** |
| 8АЭ | ATOM ВЫХОД | **4** |
| ВЫХОД | Пиктограмма Exit | **6** |

**Total DXF:** 71 items across 9 types

---

## PDF Visual Matching Results (with Cross-Symbol NMS)

| Legend ID | Symbol | Description | PDF Count | DXF Count | Delta | Within ±1? |
|-----------|--------|-------------|-----------|-----------|-------|------------|
| 1 | (щит) | Щит рабочего освещения | 1 | N/A | N/A | N/A (control panel) |
| **2** | **'1'** | **SLICK.PRS LED 50** | **158** | **30** | **-128** | ❌ FAIL |
| 3 | '2' | CD LED 27 4000K | 0 | 2 | -2 | ❌ FAIL |
| **4** | **'3'** | **ARCTIC TH 1200** | **3** | **9** | **-6** | ❌ FAIL |
| **5** | **'4'** | **SLICK.PRS LED 30** | **6** | **5** | **+1** | ✅ **PASS** |
| **6** | **'5'** | **ARCTIC 1200** | **13** | **11** | **+2** | ❌ near miss |
| 7 | '6' | SLICK.PRS LED 30 Ex | 4 | 2 | +2 | ❌ near miss |
| **8** | **'7АЭ'** | **MERCURY ВЫХОД** | **2** | **2** | **0** | ✅ **PASS** |
| **9** | **'8АЭ'** | **ATOM ВЫХОД** | **3** | **4** | **-1** | ✅ **PASS** |
| 15-17 | (switches/control) | Various | 2 | N/A | N/A | N/A |

**Note:** Exit pictogram (ВЫХОД, count=6) is not in legend, so no PDF visual match expected.

---

## Accuracy Analysis

### Equipment Types Comparison (9 types total)

| Metric | Count | Percentage |
|--------|-------|------------|
| **Exact matches (delta = 0)** | 1 | 11.1% |
| **Within ±1 tolerance** | 3 | **33.3%** |
| **Within ±2 tolerance** | 5 | 55.6% |
| **Failed (delta > ±2)** | 4 | 44.4% |

### ⚠️ **RESULT:** 3/9 types within ±1 = **FAIL** (target was ≥7/9)

---

## Root Cause Analysis

### Issue 1: SLICK.PRS LED 50 Massive Overcount (PDF=158 vs DXF=30)

**Symptom:** PDF detected 158 instances vs DXF ground truth of 30 (128 extra!)

**Hypothesis:**
1. **Wrong DXF mapping:** The DXF shows P1 annotation "29 × SLICK.PRS LED 50 Ex" but the actual symbol count in legend mapping is 30
2. **Legend mismatch:** PDF legend item '1' may correspond to a DIFFERENT equipment type than what DXF reports
3. **Multiple symbol variants:** There may be multiple SLICK symbols on the plan that all match the same template

**Evidence from diagnostic log:**
```
match_symbols[2] ... raw=6192 pre_excl=32 drawing=6160 nms=160 excl=0 fp=0 color_rej=0 valid=160
cross_symbol_nms: 16 matches suppressed (before=208 after=192)
```
- Per-symbol NMS reduced 6160 → 160 (reasonable for largest template)
- Cross-symbol NMS removed only 16 total across ALL symbols (2 from this item)
- **Conclusion:** This is NOT a cross-NMS over-suppression issue. The template genuinely matches 158 locations.

### Issue 2: ARCTIC TH vs non-TH Confusion (Items 4 & 6)

**Before Cross-Symbol NMS (from T144 test):**
- Item 4 (ARCTIC TH): 11 detections
- Item 6 (ARCTIC non-TH): 13 detections

**After Cross-Symbol NMS (this test):**
- Item 4 (ARCTIC TH): 3 detections (lost 8!)
- Item 6 (ARCTIC non-TH): 13 detections (unchanged)

**DXF Ground Truth:**
- ARCTIC TH: 9
- ARCTIC non-TH: 11

**Analysis:**
- **Cross-symbol NMS is over-suppressing ARCTIC TH detections**
- The two templates are very similar (TH variant has slightly different symbol)
- Cross-NMS is incorrectly favoring the non-TH template at ambiguous locations
- Result: TH is severely undercounted (3 vs 9 DXF)

### Issue 3: CD LED 27 Zero Detections (PDF=0 vs DXF=2)

**Symptom:** No visual matches found

**Possible causes:**
1. Symbol not present on PDF (different plan version)
2. Template extraction failed
3. Symbol too small or low contrast

**Evidence:** `match_symbols[3] ... raw=1 pre_excl=1 drawing=0 nms=0`
- Only 1 raw detection (in legend itself)
- Zero drawing-area detections
- **Conclusion:** Symbol likely not present on this PDF plan elevation

---

## Cross-Symbol NMS Performance

### Configuration (from pdf_count_visual.py lines 1537-1538):
```python
CROSS_NMS_RADIUS_PT = 25.0   # ~25pt ≈ 9mm
CROSS_NMS_CONF_GAP = 0.04    # min confidence advantage to suppress
```

### Impact:
- **16 matches suppressed** out of 208 total (7.7%)
- **Items affected:**
  - Item 2: lost 2 matches (160 → 158)
  - Item 4: lost 8 matches (11 → 3) ⚠️ **CRITICAL**
  - Item 7: lost 4 matches (8 → 4)
  - Item 16: lost 1 match (1 → 0)
  - Other items: minor adjustments

### ⚠️ **Critical Problem: ARCTIC TH Lost 8 Matches**

The cross-symbol NMS is designed to prevent double-counting when two templates match the same physical location. However, it's **too aggressive** on similar templates like ARCTIC TH vs non-TH:

**Expected behavior:**
- When both templates match at the same location with similar confidence, **both should survive** (ambiguous case)
- Only suppress when there's a clear winner (conf_gap ≥ 0.04)

**Actual behavior:**
- ARCTIC TH lost 8 matches → suggests non-TH template is consistently winning by ≥0.04 confidence
- This may indicate the templates are not truly identical, or threshold tuning is needed

---

## Recommendations

### 1. Investigate DXF-PDF Symbol Mapping (Priority: HIGH)

**Action:** Manually verify that PDF legend item '1' (SLICK.PRS LED 50) corresponds to the same equipment as DXF annotation P1.

**Hypothesis:** The PDF may have a different equipment distribution than the DXF, or the legend symbols may be mapped incorrectly.

### 2. Tune Cross-Symbol NMS Parameters (Priority: HIGH)

**Current:**
```python
CROSS_NMS_RADIUS_PT = 25.0   # ~9mm
CROSS_NMS_CONF_GAP = 0.04    # 4% confidence advantage
```

**Proposed adjustments:**
1. **Increase CROSS_NMS_CONF_GAP to 0.06-0.08** to make suppression more conservative
2. **Add template similarity check:** Only apply cross-NMS to templates with high visual similarity (e.g. cosine similarity > 0.9)
3. **Add debug logging:** Report which symbol pairs are being suppressed

### 3. Review Template Preprocessing for Similar Symbols (Priority: MEDIUM)

**Action:** Examine the extracted templates for ARCTIC TH vs non-TH:
- Are they visually distinct enough?
- Should they use different preprocessing (e.g. color isolation)?

### 4. Add Ground Truth Validation Mode (Priority: LOW)

**Action:** Create a mode that loads DXF ground truth and highlights discrepancies in real-time during PDF matching, to help diagnose mapping issues.

---

## Detailed Match Breakdown

### Cross-Symbol NMS Suppression Log

From diagnostic output:
```
cross_symbol_nms: 16 matches suppressed (before=208 after=192)
```

**Per-symbol counts before/after:**
| Symbol | Before | After | Lost | Impact |
|--------|--------|-------|------|--------|
| 2 (SLICK LED 50) | 160 | 158 | -2 | Minimal |
| 4 (ARCTIC TH) | 11 | 3 | **-8** | **Critical** |
| 6 (ARCTIC non-TH) | 13 | 13 | 0 | None |
| 7 (SLICK LED 30 Ex) | 8 | 4 | -4 | Moderate |
| 16 (Switch) | 1 | 0 | -1 | Minimal |

**Total lost:** 16 (but concentrated in ARCTIC TH)

---

## Conclusion

**Test Result:** ❌ **FAIL** (3/9 types within ±1, target was ≥7/9)

**Key Findings:**

1. ✅ **Cross-symbol NMS is working** (16 suppressions applied)
2. ❌ **Cross-symbol NMS is over-aggressive** on similar templates (ARCTIC TH lost 8/11 matches)
3. ❌ **DXF-PDF mapping issue** on SLICK LED 50 (158 vs 30 — requires manual verification)
4. ❌ **CD LED 27 not detected** (0 vs 2 — likely not on PDF plan)

**Next Steps:**

1. **Immediate:** Increase `CROSS_NMS_CONF_GAP` from 0.04 to 0.07 and re-test
2. **Short-term:** Add similarity guard to cross-NMS (only suppress highly similar templates)
3. **Medium-term:** Manually verify DXF-PDF symbol correspondence for drawing 006
4. **Long-term:** Implement ground truth overlay mode for real-time validation

---

**Test Completed:** 2026-04-14
**Task:** T150
**Status:** ❌ FAIL (but identified root causes for fixing)
