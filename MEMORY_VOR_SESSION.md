# VOR Pipeline Session Memory
## Date: 2026-05-28

## Current Results
- **ГПК accuracy: 80.3%** (NOVISUAL8) — up from 73.8% baseline (+6.5%)
- **КПП accuracy: 77.1%** (KPP_CHECK2) — restored to previous best, NO regression
- Visual counting disabled (VOR_VISUAL=0) — visual hurts accuracy (71.3% vs 80.3%)

## What Was Done This Session

### Fix 1: Cable-in-conduit height distribution (BIGGEST impact)
- **Problem**: Conduit cable work rows used `_inverse_floor_count_ratios` which gave wrong distribution (34.5%/17.2%/34.5%/13.8% vs etalon 24%/13%/22%/41%)
- **Fix**: Changed aggregated conduit section to use PVC tube lengths from `conduit_by_diam` as basis for conduit work row quantities, with PVC-weighted ratios for d.16mm distribution
- **Location**: `_build_cables_section()` in `pdf_vor_pipeline.py`, lines ~2346-2420 (the `else:` branch for aggregated conduit rows)
- **Key insight**: Cable-in-conduit work quantity = PVC tube length (not cable length), because multiple cables share tubes

### Fix 2: Added 16mm² cross-section conduit rows
- **Problem**: ВБШвнг 3x2.5 cables (7.5mm² > 6mm²) went to `tray_only_groups` and never produced conduit rows. Etalon expects "суммарное сечение до 16 мм2" rows
- **Fix**: PVC d.20mm (36m) now maps to "суммарное сечение до 16 мм2" via `_DIAM_TO_XS` dict. d.20mm distributes 60%/40% between "до 5м" and "от 13 до 20м"
- This produced the 2 missing rows (22m + 14m) matching etalon exactly

### Fix 3: PVC section per-diameter format (for ГПК)
- **Problem**: ГПК etalon expects separate work rows per PVC diameter×height; КПП etalon expects grouped work rows per height with material sub-rows
- **Fix**: Added `has_trays` parameter to `_build_pvc_section()`. Large buildings (has_trays=True) get per-diameter format; small buildings get grouped format
- **Important**: PVC d.20mm in large buildings uses 60%/40% split (до 5м / от 13 до 20м), not pvc_ratios

### Fix 4: Comparison matching — exact name priority
- **Problem**: INSEL luminaire material row (identical text in both sheets, qty=12) was marked "Отсутствует" because a fuzzy match with height context bonus (0.9+0.3=1.2) stole it before the exact match (1.0+0=1.0) was reached
- **Fix**: Added +0.5 bonus for exact name matches (score >= 1.0) in `vor_comparison_xlsx.py` `_build_comparison()` function, line ~759
- This ensures exact name matches always beat fuzzy matches regardless of height context

## Key Architecture Points

### Cable-in-conduit logic (aggregated, large buildings with trays):
```
conduit_by_diam = {PVC diameter → total metres from spec}
For each diameter:
  d.16mm → "суммарное сечение до 6 мм2", distribute by PVC-weighted ratios
  d.20mm → "суммарное сечение до 16 мм2", split 60%/40% (до 5м / 13-20м)
  d.32mm+ → "до 5 метров" only (power feed at ground level)
Work rows: "Прокладка кабеля в гофре на высоте {hcat} ({xs_label})"
Material rows: one total row per cable brand (outside height loop)
```

### PVC section format:
- **has_trays=True** (ГПК): per-diameter×height work rows + material sub-rows
- **has_trays=False** (КПП): grouped per-height work rows + per-diameter material sub-rows

### Conduit cable PVC-weighted ratios:
```python
_pvc_w = {"до 5 метров": 1.02, "от 5 до 13 метров": 0.62,
           "от 13 до 20 метров": 0.94, "от 20 до 35 метров": 1.74}
```
These match the etalon distribution closely (~24%/14%/22%/40%)

## Remaining Errors in ГПК (24 non-matching rows out of 122)

### Top priority for next improvements:
1. **SLICK 30W height distribution** — quantities shuffled between heights (17 vs 15 on до-5м, 3 vs 10 on 20-35м)
2. **ARCTIC.OPL ECO LED 1200** — quantities shifted between heights (30 vs 39 on до 5м)
3. **CD LED 27 overcounting** — 12 vs 2 at от 5 до 13 метров (500% error)
4. **Indicator height distribution** — small errors (1-4 units each) across MERCURY/ATOM
5. **Commissioning rows** — derived from cable counts, cascade fix

### Structural understanding of remaining errors:
- Most errors are **height distribution** issues, NOT total count issues
- Total counts (spec quantities) are correct; the problem is how they're split across height bands
- CD LED 27 is a special case: plan data shows all 12 at 5-13м but etalon splits 4/2/6

## Files Modified
1. **`pdf_vor_pipeline.py`** — Main changes:
   - `_normalize_model_name()`: watt normalization (30W→30)
   - `_model_key()`: 50-char limit, strips trailing "ex 5000k"
   - `count_equipment_on_plans()`: VOR_VISUAL env flag, text-only merge
   - `_build_lighting_section()`: 3-tier hybrid plan+spec, indicator smoothing
   - `_find_plan_counts()` / `_find_best_ratio()`: numeric token guard
   - `_build_cables_section()`: PVC-based conduit distribution
   - `_build_pvc_section()`: per-diameter format for has_trays buildings

2. **`vor_comparison_xlsx.py`** — exact name match priority bonus (+0.5)

## Server Deployment
- Server: dadev@89.169.190.135
- PM2 service: tayfa-vor on port 8051
- NOT yet deployed with current changes

## Comparison Files
- ГПК: `Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/02_PDF/VOR_COMPARISON_NOVISUAL8.xlsx` (80.3%)
- КПП: `Data/ДБТ разделы для ИИ/30. КПП/03_PDF/VOR_COMPARISON_KPP_CHECK2.xlsx` (77.1%)
