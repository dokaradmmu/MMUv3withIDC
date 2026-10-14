# Karad Postal Division MMU Hub Report — Build Spec v10

> **Status: canonical.** This is the only spec that should be given to a build (human
> or AI) from this point forward. **Retire v9 and every earlier version.**
>
> **Scope of this document:** the Excel output only. Webapp/UI layer, hosting, and
> code architecture are a separate, later phase — not covered here.
>
> **This file is standalone.** It assumes nothing from prior conversations. Given
> this file plus the 9 daily input files plus the 1 Master file, any competent
> builder — human or AI — should be able to produce the exact same workbook:
> same 19 sheets, same formulas, same colors, same borders, same wrapping, same
> freeze behavior, byte-for-byte equivalent in every way that matters. Every
> formula, every merge range, every color hex code, every border weight, and every
> row position needed to do that is written out below.
>
> **Changed in v10 (sorting + conditional formatting only — no layout, formula, or
> UI changes):**
> 1. **Excl BOs - Combined / Sub Division**: the "has data" group is now sorted by
>    Delivery % (Range) **ascending** (worst office first) instead of A→Z by name.
>    The "no delivery activity" group is unchanged — still grouped at the bottom,
>    still A→Z by name. Ties on Delivery % (Range) break alphabetically by name.
> 2. **Defaulters sheets**: all 3 tables (Delivery, DSS, COD) are now sorted by
>    that table's own Range % **ascending** (worst defaulter first) instead of
>    A→Z by name. Ties break alphabetically by name.
> 3. **Divisional Summary**: sub-division row order is now computed alphabetically
>    (Division Total always last) instead of a hardcoded list — byte-identical
>    output for Karad's 3 sub-divisions today, but no longer requires a code change
>    if sub-division names ever change.
> 4. **Conditional formatting — threshold-based red/green highlighting replaces the
>    3-color gradient** in three places (see §5.1 and the per-sheet sections below
>    for exact thresholds and ranges):
>    - Divisional Summary: entire data area (sub-division rows **and** the Division
>      Total row), replacing the old gradient there completely.
>    - Excl BOs - Combined / Sub Division: the Total/Subtotal row(s) only — these
>      previously had no conditional formatting at all. Data rows keep the
>      existing 3-color gradient, unchanged.
>    - Defaulters sheets: the Range % columns and two raw-count columns get
>      sheet-specific treatment (2-color scale, data bars, and a two-tier
>      red-highlight rule on COD) — see §5.5 and §6.4.

---

## 0. The 10 files you need, every day

| # | File | What it is |
|---|---|---|
| 1 | Office Master File (.xlsx) | Changes rarely. 288 offices, 5 relevant columns after cleanup. |
| 2 | Delivery Productivity — Range 1 (.csv) | Covers roughly the 1st–15th of the month (source system caps a single export at ~15 days). |
| 3 | Delivery Productivity — Range 2 (.csv) | Covers ~16th to the report date. Summed with Range 1, never replaces it. |
| 4 | Delivery Productivity — Single Date (.csv) | One specific day. |
| 5 | DSS Usage — Range (.csv) | 1st of month to report date, pre-aggregated one row per office. |
| 6 | DSS Usage — Single Date (.csv) | One specific day, same columns as #5. |
| 7 | Letter Box Clearance — Range (.csv) | Transaction-level, one row per office per working day. |
| 8 | Letter Box Clearance — Single Date (.csv) | One specific day, pre-aggregated. |
| 9 | COD Digital Transaction — Range (.csv) | Transaction-level, multiple rows per office. |
| 10 | COD Digital Transaction — Single Date (.csv) | One specific day, transaction-level. |

**Critical rule, non-negotiable:** the Single Date and the Range start/end dates used
anywhere in this report — in the title, the office record number, every column
header, every date label — are **whatever the person tells you they are for this
run**, never inferred by parsing a filename. File names vary (`DSS_Range_01-24_July`
one day, `DSS_Usage_28_July` the next with no "Range"/"Single Date" in the name at
all) and must never be trusted. Confirm the 3 dates (Single Date, Range Start, Range
End) with the person, or derive them from explicit instruction, before building
anything. If a file's *internal content* doesn't match what its name implies (e.g. a
file named like a Single Date file but whose totals are clearly a monthly
accumulation), trust the content, flag the mismatch, and ask.

### File schemas (exact columns, in order)

**Master** (Office Master File, before cleanup — 14 columns, only 5 matter):
```
Circle Office ID, Circle Name, Region Office ID, Region Office Name, Division Office ID,
Division Office Name, Sub Division office ID, Sub Division Name, Sub Office Name,
Office ID, Office Name, Office Type Code, PIN, Taluka
```
Keep only: `Office ID → office_id`, `Office Name → office_name`, `Sub Office Name →
sub_office`, `Sub Division Name → sub_division`, `Office Type Code → office_type`.
Expect 288 rows, 251 BPO / 36 SPO / 1 HPO, zero duplicate Office IDs (Office Names
may legitimately duplicate — Office ID is the only valid join key).

**Delivery Productivity** (Range 1, Range 2, and Single Date — identical schema):
```
office-id, office-name, postman-count, invoice-count, delivery-count,
deposit-count, return-count, redirection-count, beat-diversion-count
```
One row per office per file (Range files are already pre-aggregated for their
window, not transaction-level).

**DSS Usage** (Range and Single Date — identical schema):
```
office_id, office_name, event_date, dss_art_count, pdm_art_count, Mobile Usage %,
total_pdm_art_count, total_dss_art_count, total_dss_count
```
One row per office, pre-aggregated. Contains an `office_id = 0` summary row (exclude
it) and may contain office IDs not present in Master (exclude and log, never
silently drop). `event_date` is frequently `null` or a placeholder date — ignore it,
it is not the report date.

**Letter Box Clearance — Range**:
```
office-id, office-name, lb-location, total-letterboxes, total-letters,
avg-cleared, total-cleared, working-days, clearance-percentage
```
**One row per office PER WORKING DAY in the window** — transaction-level.
`total-letterboxes` repeats per office across its daily rows (a per-day count, not
a running total — do not treat it as cumulative).

**Letter Box Clearance — Single Date**:
```
office-id, office-name, total-letterboxes, total-letters, avg-cleared,
total-cleared, clearance-percentage
```
One row per office (37 SPO+HPO) plus one summary row with `office-id = 0` (exclude).

**COD Digital Transaction** (Range and Single Date — identical schema):
```
Circle Name, Region Name, Division Name, Sub Division Name, Office Name, Office ID,
Customer Name, Product Name, COD Cash Count, COD Cash Amount, COD Digital Count,
COD Digital Amount, Total COD Count, Total COD Amount, Booking COD Amount,
COD Collection Performance (%)
```
Transaction-level — multiple rows per office, one per COD booking.

---

## 1. Metric formulas (unchanged core logic — locked)

### 1.1 Delivery Productivity
```
Invoice Count (Range)      = Σ(invoice-count)  [Range 1 file + Range 2 file, summed]
Delivery % (Range)         = (Σ Invoice[R1+R2] − Σ Deposit[R1+R2]) / Σ Invoice[R1+R2]
Invoice Count (Single Date) = invoice-count      [from the Single Date file directly]
Delivery % (Single Date)    = (Invoice_SD − Deposit_SD) / Invoice_SD
```
Blank (not zero) if the relevant Invoice count is 0 or the office has no rows in
that file — division by zero must never render as 0%.

### 1.2 DSS Usage
```
DSS % (Range)       = total_dss_art_count / total_pdm_art_count   [Range file, direct
DSS % (Single Date)  = total_dss_art_count / total_pdm_art_count    one row/office]
```
Blank if `total_pdm_art_count = 0`. A row with `pdm=0 AND dss=0` is blank (no
activity that period); a row with `pdm>0 AND dss=0` is a real 0%.

### 1.3 COD Digital Transaction
```
Total COD Count (Range)      = Σ(Total COD Count)   [grouped by Office ID, transaction-
COD % (Range)                = Σ(COD Digital Count) / Σ(Total COD Count)   level files]
Total COD Count (Single Date) = Σ(Total COD Count)   [same, Single Date file]
COD % (Single Date)           = Σ(COD Digital Count) / Σ(Total COD Count)
```
Blank if `Total COD Count = 0` — very common for small BPOs with no COD activity
that period.

### 1.4 Letter Box Clearance (SPO + HPO only — BPOs blank everywhere, no letterbox infrastructure)
```
LB % (Range)        = Σ(total-cleared) / Σ(total-letterboxes)   [summed across every
                                                                   working-day row]
LB % (Single Date)   = total-cleared_SD / total-letterboxes_SD   [Single Date file,
                                                                   direct ratio]
```
**Total Letter Boxes** (display column): prefer the Single Date file's value; if
absent there, average the Range file's (constant-per-office) value across its rows.

### 1.5 Worked reference example — regression-check these on every run if the same offices appear
| Office | Delivery % (R) | DSS % (R) | DSS % (SD) | COD % (R) |
|---|---|---|---|---|
| Nade B.O (24107205) | 79.14% | 87.17% | **0.00%** (real zero) | 50.00% |
| Dhoroshi B.O (24107117) | 98.17% | 87.16% | **0.00%** (real zero) | 100.00% |
| Korti B.O (24107175) | 95.65% | 89.67% | **blank** (no PDM/DSS data that day) | 0.00% (real zero) |

---

## 2. Thresholds
| Type | Threshold | Basis |
|---|---|---|
| Delivery / DSS / COD defaulter | < 90% | **Range** value only |
| Exceptional defaulter | < 30% | any of 8 parameter-instances (4 metrics × Range/Single Date) |

---

## 3. Master hidden sheet — locked 5-column layout
| Col | Field |
|---|---|
| A | office_id |
| B | office_name |
| C | sub_office |
| D | sub_division |
| E | office_type |

---

## 4. Workbook architecture — 19 sheets

| # | Sheet | Visible | Tab color | Freeze panes |
|---|---|---|---|---|
| 1 | Raw Data | Yes | `#1F4E78` (navy) | **`G4`** (rows 1–3 + cols A–F locked) |
| 2 | Excl BOs - Combined | Yes | `#C00000` (red) | **None — unfrozen** |
| 3 | Excl BOs - Sub Division | Yes | `#C00000` (red) | **None — unfrozen** |
| 4 | Defaulters - ASP Karad West | Yes | `#E67E22` (orange) | **None — unfrozen** |
| 5 | Defaulters - SDIP Karad East | Yes | `#E67E22` (orange) | **None — unfrozen** |
| 6 | Defaulters - SDIP Vaduj | Yes | `#E67E22` (orange) | **None — unfrozen** |
| 7 | Divisional Summary | Yes | `#7030A0` (purple) | **None — unfrozen** |
| 8 | Exceptional Defaulters | Yes | `#C00000` (red) | **`A4`** |
| 9 | Master | Hidden | — | — |
| 10 | Src_DP_R1 | Hidden | — | — |
| 11 | Src_DP_R2 | Hidden | — | — |
| 12 | Src_DP_SD | Hidden | — | — |
| 13 | Src_DSS_R | Hidden | — | — |
| 14 | Src_DSS_SD | Hidden | — | — |
| 15 | Src_LB_R | Hidden | — | — |
| 16 | Src_LB_SD | Hidden | — | — |
| 17 | Src_CODD_R | Hidden | — | — |
| 18 | Src_CODD_SD | Hidden | — | — |
| 19 | Validation Log | Hidden | — | — |

Only **Raw Data** carries live cross-sheet formulas. Every other visible sheet is
hardcoded pre-computed values from one single computation pass — never
independently re-derived. Native gridlines are off on every visible sheet.

### 4.1 Date-label strings used throughout (compute once, reuse everywhere)
```
range_verbose = "{range_start:%d %B} to {range_end:%d %B}"    e.g. "01 July to 28 July"
range_compact = "{range_start:%d}-{range_end:%d %B}"           e.g. "01-28 July"
single_label  = "{single_date:%d %B}"                          e.g. "28 July"
```
`range_verbose` is used **only on Raw Data**. `range_compact` and `single_label` are
used on every other sheet. Dates never include the year in these labels (year
appears only in the title and record number). The Single Date sub-label is always
**plain text**, never an Excel date value.

### 4.2 Title and record number (all visible sheets except Defaulters and Exceptional Defaulters use the exact title wording below)
```
Title:          MMU Report Only S.O {single_date:%d.%m.%Y}
Record number:  No. KRDDN/G/Mails/MMU/Performance/{single_date + 1 day:%d.%m.%y}
```
Defaulters sheets use: `MMU Report Defaulters (<90%) - {sub_division} - {single_date:%d.%m.%Y}`
Exceptional Defaulters uses: `MMU Report Exceptional Defaulters (< 30%) - {single_date:%d.%m.%Y}`

**Defaulters sheets do NOT carry the record-number row at all** — removed by
request; title sits alone on row 1.

---

## 5. Global styling — exact values

| Element | Font | Size | Bold | Color | Fill | Align |
|---|---|---|---|---|---|---|
| Title | Arial | 16 | Yes | default | — | center |
| Record number | Arial | 12 | No | default | — | right |
| Marathi standing text | Nirmala UI | 12 | Yes | default (or see §5.3) | — | left (or center, see §5.4) |
| Column header | Arial | 12 | Yes | white | `#1F4E78` navy | center |
| Section label (sub-division / table name banner) | Arial | 13 | Yes | white | `#C00000` red | center |
| Body cell | Arial | 12 | No | default | — | center (left for name columns) |
| Total/subtotal row | Arial | 12 | Yes | default | `#FCE4D6` peach | center (left for name column) |
| Exceptional Defaulters — every data row | Arial | 12 | Yes | `#FF0000` red | — | center (left for name/parameter) |
| Note (e.g. "No defaulter offices…") | Arial | 12 | No, italic | default | — | left |

Percentage format: `0.00%`. Count format: `#,##0`.

### 5.1 Grid and group-segregation borders
- **Every header and body cell**: thin border, color `#808080` (mid grey), on all
  four sides — this is the baseline grid everywhere.
- **Metric-group segregation** (Excl BOs - Combined, Excl BOs - Sub Division,
  Divisional Summary only — these are the sheets with the 2-row grouped header,
  §6.2/§6.3/§6.5): a **thick** left border, color `#2C2C2C` (near-black), on the
  first column of each metric group (Delivery %, DSS Usage %, COD Digital
  Transaction %, LB Clearance %) — runs through the header rows, every data row,
  and the total/subtotal row, all the way down the column.
- **Header/data segregation** (same 3 sheets): a **thick** `#2C2C2C` border on the
  **top** edge of the group-name row, and a **thick** `#2C2C2C` border on the
  **bottom** edge of the date-sub-label row directly below it (this bottom edge is
  the actual header→data boundary). Both run the full width of the table.
- **Total/subtotal row framing** (same 3 sheets): a **thick** `#2C2C2C` border on
  both the **top** and **bottom** edges of the total/subtotal row, full width —
  visually boxing it off from the data above and the blank spacer row below.
- Raw Data, Defaulters sheets, and Exceptional Defaulters use only the thin grid —
  no group segregation borders (they don't have the 2-row grouped-header structure).

### 5.2 Wrap text — exactly these rows, nowhere else
- **Raw Data**: the single header row (all 19 columns).
- **Excl BOs - Combined**: the date-sub-label row only (not the group-name row above it).
- **Excl BOs - Sub Division**: the date-sub-label row in **each** of the 3 blocks.
- **Defaulters sheets**: the column-header row in **each** of the 3 tables (Sr. No. / Office ID / Office Name / Office Type / …).
- **Divisional Summary**: the "Karad Division Total" row in **each** of the 3 tables.
- Nowhere else gets wrap text (group-name rows, ordinary data rows, and Excl BOs
  Combined/Sub Division's own total rows stay un-wrapped).

### 5.3 Red-highlighted words (Defaulters sheets only)
Within each table's Marathi standing text, the following exact substring is
rendered in font color `#FF0000` while the rest of the paragraph stays the default
color — same font (Nirmala UI 12 bold), same everything else, just the color
differs for that one run of text:
| Table | Substring colored red |
|---|---|
| Delivery | `डिलीवरी%` |
| DSS | `DSS%` |
| COD | `COD Digital Transaction %` |

Implementation note: this requires a **rich string** (multiple font runs within one
cell) — a merged cell's format alone cannot mix two font colors. Write the merge
with an empty value first to establish the cell format across the whole range,
then write the actual rich-text content (default-color run + red run + default-color
run) into the top-left cell of that merge.

### 5.4 Centered Marathi text (2 sheets only)
- **Excl BOs - Combined** and **Excl BOs - Sub Division** (every block): the
  Marathi standing-text row is **horizontally centered** (in addition to its normal
  wrap and bold Nirmala UI styling).
- Divisional Summary's and Defaulters sheets' Marathi text rows stay **left-aligned**
  (default) — centering is specific to the two Excl BOs sheets only.

### 5.5 Threshold conditional formatting (v10, new)
Distinct from the 3-color gradient (§6.1 etc.): a **cell-value rule pair** —
green fill if strictly greater than the upper threshold, red fill if strictly
less than the lower threshold, no highlight in between. Standard Excel
"Good"/"Bad" cell-style colors throughout: green fill `#C6EFCE` / font
`#006100`; red fill `#FFC7CE` / font `#9C0006`.

| Metric | Green if > | Red if < |
|---|---|---|
| Delivery % | 95% | 90% |
| DSS Usage % | 99% | 95% |
| COD Digital Transaction % | 90% | 80% |
| Letter Box Clearance % | 99% | 95% |

Applied to:
- **Divisional Summary** (§6.5): every metric-pair column, across the 3
  sub-division rows **and** the Division Total row, per table. Replaces the
  3-color gradient entirely on this sheet. LB Clearance % is skipped for the
  "Only B.Os" table (blank for every BPO row there).
- **Excl BOs - Combined / Sub Division** (§6.2/§6.3): the Total/Subtotal
  row(s) only, per metric-pair column (excludes the Total Letter Boxes count
  column). Data rows are unaffected — they keep the 3-color gradient.

The Defaulters sheets (§6.4) use a related but distinct scheme — see that
section for the exact columns and rule types (2-color scale / data bar /
two-tier highlight).

---

## 6. Per-sheet blueprint

### 6.1 Raw Data
**19 columns (A–S).** Row 1: title (merged A1 through the 4th-from-last column) +
record number (merged across the last 3 columns, right-aligned) in the same row.
Row 2: blank spacer. Row 3: column headers, wrapped, height 30. Rows 4 onward: one
row per office (288 rows), **sorted by Office ID ascending**.

| Col | Header | Content |
|---|---|---|
| A | Sr. No. | 1–288 |
| B | Office ID | from Master |
| C | Office Name | from Master |
| D | Sub Office Name | from Master |
| E | Sub Division Name | from Master |
| F | Office Type | from Master |
| G | `Invoice Count ({range_verbose})` | live formula, count format |
| H | `Delivery % ({range_verbose})` | live formula, percentage |
| I | `Invoice-Count ({single_label})` | live formula, count format |
| J | `Delivery % ({single_label})` | live formula, percentage |
| K | `DSS % ({range_verbose})` | live formula, percentage |
| L | `DSS % ({single_label})` | live formula, percentage |
| M | `Total COD Count ({range_verbose})` | live formula, count format |
| N | `COD % ({range_verbose})` | live formula, percentage |
| O | `Total COD Count ({single_label})` | live formula, count format |
| P | `COD % ({single_label})` | live formula, percentage |
| Q | Total Letter Boxes | live formula, count format, blank if BPO |
| R | `LB % ({range_verbose})` | live formula, percentage, blank if BPO |
| S | `LB % ({single_label})` | live formula, percentage, blank if BPO |

**Live formulas** (office row's Office ID cell = `$B{row}`):
```
Invoice Count (R)   = SUMIFS(Src_DP_R1!$D:$D,Src_DP_R1!$A:$A,$B{row})
                      + SUMIFS(Src_DP_R2!$D:$D,Src_DP_R2!$A:$A,$B{row})
Delivery % (R)      = IF(InvR=0,"",(InvR − DepR)/InvR)
                      where DepR is the same SUMIFS pattern against column F of both sheets
Invoice Count (SD)  = SUMIFS(Src_DP_SD!$D:$D,Src_DP_SD!$A:$A,$B{row})
Delivery % (SD)     = IF(InvSD=0,"",(InvSD − DepSD)/InvSD)  [DepSD = same pattern, col F]
DSS % (R)           = IF(PdmR=0,"",DssR/PdmR)
                      PdmR = SUMIFS(Src_DSS_R!$G:$G,Src_DSS_R!$A:$A,$B{row})
                      DssR = SUMIFS(Src_DSS_R!$H:$H,Src_DSS_R!$A:$A,$B{row})
DSS % (SD)          = same pattern against Src_DSS_SD
Total COD Count (R) = SUMIFS(Src_CODD_R!$M:$M,Src_CODD_R!$F:$F,$B{row})
COD % (R)           = IF(TotR=0,"",DigR/TotR)
                      DigR = SUMIFS(Src_CODD_R!$K:$K,Src_CODD_R!$F:$F,$B{row})
Total COD Count (SD)/COD % (SD) = same pattern against Src_CODD_SD
Total Letter Boxes  = blank if Office Type="BPO"; else prefer SUMIFS(Src_LB_SD!$C:$C,...)
                      if that office has a row there (COUNTIFS check), else
                      AVERAGEIFS(Src_LB_R!$D:$D,Src_LB_R!$A:$A,$B{row})
LB % (R)            = blank if BPO; else IF(BoxR=0,"",ClrR/BoxR)
                      BoxR=SUMIFS(Src_LB_R!$D:$D,...), ClrR=SUMIFS(Src_LB_R!$G:$G,...)
LB % (SD)           = blank if BPO; else IF(BoxSD=0,"",ClrSD/BoxSD)
                      BoxSD=SUMIFS(Src_LB_SD!$C:$C,...), ClrSD=SUMIFS(Src_LB_SD!$F:$F,...)
```
If a data category is entirely unavailable this run (a file wasn't provided), its
formulas are skipped and the cells are left as true blanks instead.

Column widths (A→S): `8.57, 11.57, 29, 22, 22.43, 13.71, 18, 22.57, 22.57, 28,
18.43, 23.86, 23.86, 18.71, 18.71, 24.14, 21.43, 16.57, 22.14`

Conditional formatting: 3-color scale (red `#F8696B` min → yellow `#FFEB84` 50th
percentile → green `#63BE7B` max) independently on each of columns H, J, K, L, N,
P, R, S (the 8 percentage columns), across rows 4–291. Count columns (G, I, M, O,
Q) get no color scale.

### 6.2 Excl BOs - Combined
**11 columns (A–K), one flat table, 37 SPO+HPO offices.**
Row 1: title. Row 2: record number. Row 3: Marathi Text A, **centered**, wrapped,
row height 45. Row 4: group-name header row. Row 5: date-sub-label row, **wrapped**.
Rows 6 onward: data (sorted per §7), then a total row, then nothing (sheet ends).

Header structure (2 rows, merged):
| Col | Row 4 (group name) | Row 5 (sub-label) |
|---|---|---|
| A | `Sr. No.` (merged A4:A5) | — |
| B | `Office Name` (merged B4:B5) | — |
| C–D | `Delivery % (R)` (merged) | `{range_compact}` / `{single_label}` |
| E–F | `DSS Usage %` (merged) | `{range_compact}` / `{single_label}` |
| G–H | `COD Digital Transaction %` (merged) | `{range_compact}` / `{single_label}` |
| I–K | `Letter Box Clearance %` (merged) | `Total Letter Boxes` / `{range_compact}` / `{single_label}` |

Total row label: `Karad Division Total (SPO+HPO)`, volume-weighted (Σ/Σ), peach
fill, bold, boxed top+bottom in `#2C2C2C` thick (not wrapped).

Column widths: `7, 26, 13, 13, 13, 13, 13, 13, 15, 13, 13`.
Conditional formatting: 3-color scale on C, D, E, F, G, H, J, K across the data rows
only (excludes the total row). **The total row itself (v10, new)** gets the
threshold red/green highlighting from §5.5 instead — per metric-pair column
(C:D, E:F, G:H, J:K), skipping the Total Letter Boxes count column (I).

### 6.3 Excl BOs - Sub Division
**Same 11 columns as §6.2**, but split into 3 full blocks — one per sub-division, in
this fixed order: **ASP Karad West, SDIP Karad East, SDIP Vaduj**. Each block
repeats the **entire** header stack (not just a section label):

```
Row N+0: Title — "MMU Report Only S.O {date}"
Row N+1: Record number
Row N+2: Marathi Text A, CENTERED, wrapped, height 45
Row N+3: Section label — the sub-division name, red banner
Row N+4: Group-name header row
Row N+5: Date-sub-label row, WRAPPED
Row N+6 onward: that sub-division's offices (sorted per §7)
Then: subtotal row (volume-weighted, "{Sub Division} Subtotal", peach, boxed top+bottom)
Then: one blank spacer row
Then the next block starts at Row N+0 again.
```
Column widths and conditional-formatting rule: identical to §6.2, but the color
scale (and the v10 threshold highlighting on each block's Subtotal row) is
scoped **independently per block** (never across the whole sheet).

### 6.4 Defaulters sheets (×3 — ASP Karad West, SDIP Karad East, SDIP Vaduj)
**No record-number row.** Row 1: title only (`MMU Report Defaulters (<90%) -
{sub_division} - {date}`). Then 3 stacked tables, **Delivery → DSS → COD**, each:

```
Section label row (red banner)
Marathi text row — height 45, left-aligned (not centered), with the specific
  substring in that table's text colored #FF0000 per §5.3
Column header row — WRAPPED
Data rows: offices with that metric's Range value < 90%, sorted by that metric's
  own Range % ASCENDING (v10: worst defaulter first; ties break alphabetically
  by office name). If none qualify: one italic note row instead of an empty table.
Two blank spacer rows before the next table.
```

**Delivery table (7 columns, A–G):**
`Sr. No. | Office ID | Office Name | Office Type | Delivery % ({range_compact}) |
Articles in Deposit ({single_label}) | Delivery % ({single_label})`
— the "Articles in Deposit" column holds the Single Date's raw deposit-count
(undelivered articles that day) — the actual shortfall behind the percentage next
to it.

**DSS table (7 columns, A–G):**
`Sr. No. | Office ID | Office Name | Office Type | DSS Usage ({range_compact}) |
Articles Invoiced on {single_label} | DSS Usage ({single_label})`
— "Articles Invoiced" holds the Single Date's total PDM article count (DSS %'s
denominator).

**COD table (8 columns, A–H):**
`Sr. No. | Office ID | Office Name | Office Type | Total COD Articles Received
({range_compact}) | COD Digital Txn ({range_compact}) | Total COD Articles Received
({single_label}) | COD Digital Txn ({single_label})`
— "Total COD Articles Received" (both windows) holds the raw Total COD Count
(COD %'s denominator); "COD Digital Txn" holds the actual percentage despite its
count-like name (this exact wording is locked from the reference file).

Title row is merged across 8 columns (the width of the widest table, COD) even
though the narrower Delivery/DSS tables only use 7 — this keeps the title bar
visually consistent across all 3 tables on the sheet.
Column widths (A→H): `7, 11, 26, 11, 20, 22, 20, 22`.
Conditional formatting (v10, revised — no longer a uniform 3-color scale on every
percentage column; ranges computed dynamically each run from that table's actual
row count):
| Table | Range % column | Single Date % column | Count column(s) |
|---|---|---|---|
| Delivery | 2-color scale, red `#F8696B` → near-white `#FCFCFF` | 3-color scale (unchanged, §6.1-style) | "Articles in Deposit" (SD): data bar, color `#FF555A` |
| DSS | 2-color scale, same colors as Delivery | 3-color scale (unchanged) | "Articles Invoiced" (SD): no formatting |
| COD | Two-tier highlight: red text (no fill) if < 60%; red text + red fill `#FFC7CE` if exactly 0% | 3-color scale (unchanged) | "Total COD Articles Received" (Range): data bar, color `#FF555A`; (SD): no formatting |

### 6.5 Divisional Summary
**10 columns (A–J), one top-level title block (not repeated per table).**
Row 1: title. Row 2: record number. Row 3: Marathi Text A, left-aligned (not
centered), wrapped, height 45. Then 3 stacked tables in this order:

1. **Excluding B.Os (SPO + HPO only)**
2. **Only B.Os (BPO only)**
3. **Including B.Os (All 288 Offices)**

Each table:
```
Section label row (red banner)
Group-name header row: Sr. No. | Sub Division | Delivery % | DSS Usage % |
  COD Digital Txn % | LB Clearance %  (each metric spans 2 columns)
Date-sub-label row: {range_compact} / {single_label} under each metric pair
  (NOT wrapped — Divisional Summary's own header rows stay un-wrapped;
  only the total row below gets wrap, see §5.2)
4 data rows: sub-divisions sorted ALPHABETICALLY by name (v10: computed, not a
  hardcoded list — for Karad's 3 sub-divisions this is ASP Karad West, SDIP
  Karad East, SDIP Vaduj, same as before), then "Karad Division Total" —
  WRAPPED, peach fill, bold, boxed top+bottom in #2C2C2C thick
  (volume-weighted, Σ/Σ)
One blank spacer row before the next table.
```
In the "Only B.Os" table, the LB Clearance % columns are blank for every row (BPOs
have no letterbox infrastructure).
Column widths: `7, 18, 13, 13, 13, 13, 13, 13, 13, 13`.
Conditional formatting (v10, revised): the 3-color gradient is replaced entirely
by the threshold red/green highlighting from §5.5, per metric-pair column,
covering the sub-division rows **and** the Division Total row (unlike the old
gradient, which excluded the Total row). LB Clearance % is skipped for the
"Only B.Os" table.

### 6.6 Exceptional Defaulters
**7 columns (A–G), flat list, one row per office-parameter-instance below 30%.**
Row 1: title. Row 2: record number. Row 3: column headers. Row 4 onward: data,
sorted by sub-division, then office name, then parameter.

`Sr. No. | Office ID | Office Name | Sub Division | Office Type | Parameter | Value %`

**Parameter column uses real dates, matching every other sheet** — e.g. `Delivery %
({single_label})`, `COD % ({range_compact})` — never the generic "(Range)" /
"(Single Date)" placeholders.

Every data row (all 7 columns) is Arial 12 **bold, red `#FF0000`** — the entire row,
not just a flagged cell, since every row here is by definition an exceptional case.
If none qualify: one italic note row instead.
Column widths: `7, 11, 26, 16, 11, 26, 10`.
Conditional formatting: 3-color scale on the Value % column across all data rows.

### 6.7 Hidden Src_* sheets — verbatim mirror of that run's input files
One sheet per file, header row + all data rows, no filtering, no re-sorting:
`Src_DP_R1`, `Src_DP_R2`, `Src_DP_SD` (Delivery Productivity schema, §0), `Src_DSS_R`,
`Src_DSS_SD` (DSS schema), `Src_LB_R`, `Src_LB_SD` (their respective LB schemas),
`Src_CODD_R`, `Src_CODD_SD` (COD schema). If a file wasn't provided this run, the
sheet contains a single note: `(no file uploaded this run)`.

### 6.8 Hidden Validation Log
3 columns — Check, Detail, Result. Logs, every run: Master reconciliation
(row/BPO/SPO/HPO counts, duplicate-ID count), DSS Range/Single-Date excluded rows
(office_id=0 and any not-in-Master IDs, by name+ID), data availability per
parameter, and the canonical Single Date / Range Start / Range End used.

---

## 7. Sort rules
| Sheet | Rule |
|---|---|
| Raw Data | All 288 offices, Office ID ascending. |
| Master, all Src_* sheets | As-loaded from source file, never re-sorted. |
| Excl BOs - Combined | **(v10)** Offices with a genuine Delivery % (Range) value first, sorted by Delivery % (Range) ASCENDING — worst office first, ties by name; offices with no delivery activity that period (blank Delivery % R) unchanged — grouped at the bottom, A→Z by name. No separator row. |
| Excl BOs - Sub Division | Same two-key sort, applied within each sub-division block independently. |
| Defaulters tables | **(v10)** Sorted by that metric's own Range % ASCENDING — worst defaulter first, ties by office name (within that metric's <90% filter). |
| Divisional Summary | **(v10)** Computed alphabetically by sub-division name → Division Total always last. For Karad today: ASP Karad West → SDIP Karad East → SDIP Vaduj → Karad Division Total (same output as the old hardcoded order). |
| Exceptional Defaulters | Sub-division, then office name, then parameter. |
| Sub-division block order everywhere | Always ASP Karad West → SDIP Karad East → SDIP Vaduj (alphabetical for Karad's 3 — see Divisional Summary row above). |

---

## 8. Mandatory verification before delivering the file

1. **Schema check** every input file's actual columns against §0 before computing
   anything. For Letter Box Range specifically, check rows-per-office consistency;
   flag only if rows-per-office is 1 while the calendar span is more than a few
   days (a 1-day span early in the month legitimately has 1 row/office).
2. **Percentage bounds**: every percentage cell, every sheet, must be in [0%, 100%].
3. **Master reconciliation**: 288 rows, 251/36/1 BPO/SPO/HPO split, zero duplicate
   Office IDs.
4. **Formula recalculation**: if built with live formulas (Raw Data), recalculate
   the whole workbook (e.g. via a LibreOffice headless pass) and confirm zero
   formula errors.
5. **Independent cross-check**: recompute every Raw Data percentage and count value
   via a separate, independent calculation path and confirm they match the
   in-workbook values exactly (raw SUM-type count columns may legitimately show 0
   where the matching percentage column correctly shows blank — that's the
   0-denominator distinction, not a bug).
6. **Aggregation weighting**: every total/subtotal row must be Σ/Σ (volume-weighted
   from raw counts), never an average of the displayed percentages. Spot-check by
   hand.
7. **Blank-vs-zero spot check**: confirm the pattern in §1.5 still holds for
   offices with comparable data patterns.
8. **Layout fidelity**: sheet order, hidden/visible state, tab colors, all freeze
   panes, all merge ranges, wrap-text rows, centered rows, red-highlighted words,
   and border segregation all match §4–§6 exactly.
9. **Sort-order check**: confirm the delivery-first/non-delivery-last split in both
   Excl BOs sheets (v10: has-data group ordered by Delivery % ascending, not by
   name), the Defaulters tables' worst-first ordering, and the alphabetical
   sub-division ordering everywhere it appears.

If any input file is missing, degrade gracefully per parameter (blank cells, never
0/N/A/dash) rather than blocking the whole report, and note in the Validation Log
which parameters were unavailable that run.

---

## 9. Versioning discipline
- One canonical spec file exists at a time. This is v10; v9 and earlier are retired.
- Any layout or logic change, however small, bumps the version number — no silent
  parallel variants under the same version name.
- A build session should be handed this literal file, not asked to reconstruct it
  from memory of past conversations.
