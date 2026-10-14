# Karad Postal Division MMU Hub — Streamlit Webapp (v2)

Implements the MMU Hub end goal: a permanently-hosted Office Master File, an
India-Post-branded UI, 9 named input placeholders each with its own calendar
picker, graceful degradation with an explicit "Generate anyway?" confirmation,
and a live Divisional Summary preview widget. The generated workbook follows
**MMU Hub Build Spec v10** sheet-for-sheet — see `MMU_Hub_Build_Spec_v10.md` for
the complete, standalone description of exactly what the output should look like
(this is the reference to check the app against, independent of this codebase).

This is v2 of the report generator — a separate repo from the original
[`mmu-hub-karad`](https://github.com/dokaradmmu/mmu-hub-karad) so the
production app keeps running untouched while this version is rolled out.

## What's new in v2

Sorting and conditional-formatting changes only — no layout, formula, or UI
changes. Full detail (exact thresholds, columns, colors) is in
`MMU_Hub_Build_Spec_v10.md`'s changelog callout and §5.5/§7; summary:

- **Excl BOs sheets**: the "has data" group now sorts by Delivery % (Range)
  ascending (worst office first) instead of alphabetically.
- **Defaulters sheets**: all 3 tables now sort by that metric's own Range %
  ascending (worst defaulter first) instead of alphabetically. Applied
  consistently across all 3 sub-division sheets.
- **Divisional Summary**: sub-division order is now computed alphabetically
  instead of hardcoded (no visible change for Karad today).
- **Conditional formatting**: threshold-based red/green highlighting added to
  Divisional Summary (replacing its gradient) and to the Excl BOs Total/Subtotal
  rows (previously unformatted); Defaulters sheets get a 2-color scale, data
  bars, and a two-tier COD highlight on top of the existing Single-Date gradient.

Validated by reconstructing the original 9 input CSVs from a real production
run's hidden `Src_*` sheets, regenerating the workbook through this exact code,
and diffing cell-by-cell against both the original output (untouched sheets:
zero diffs) and a manually-edited reference copy (changed sheets: zero diffs).

## Files

| File | Purpose |
|---|---|
| `app.py` | Router: sidebar navigation between the two reports (`st.navigation`, needs Streamlit ≥ 1.36). Main file to deploy. |
| `ui.py` | Shared UI pieces — branded header, Master lookup (session upload wins over the saved default). |
| `idc.py` | **IDC Delivery Report** — validation, KPI computation and workbook builder (one row per delivery person). |
| `idc_report.py` | IDC page: two CSV uploads (DPR required, DSS optional) + Range From / To + Single Date pickers, validation, preview, download. |
| `tests/test_idc.py` | Synthetic tests for `idc.py`. |
| `mmu_report.py` | MMU Daily Report page — Streamlit UI — branded header, persistent Master file management, 9 named uploaders with per-file date pickers, derived report-date summary, file-status checklist, "Generate → warn → confirm" flow, live Divisional Summary widget, download button. |
| `engine.py` | Pure computation layer — reads uploaded files into DataFrames, computes all 4 KPIs × Range/Single Date with correct blank-vs-zero handling, builds every aggregated table. Also holds the Master-file / logo persistence helpers. Independently testable, no Streamlit or Excel code. |
| `workbook.py` | Builds the `.xlsx` via `xlsxwriter`, replicating Build Spec v10's layout sheet-for-sheet — grouped 2-row headers with real dates, thick border segregation between metric groups and around header/total blocks, selective wrap-text and centering, red-highlighted keywords in the Defaulters sheets' Marathi text, no record-number row on Defaulters sheets, real dates in Exceptional Defaulters' Parameter column, and (v2) threshold-based conditional formatting on Divisional Summary / Excl BOs total rows / Defaulters sheets. Raw Data's formulas are written **with cached values attached**, so the file opens with correct numbers immediately — no LibreOffice recalculation needed at runtime. |
| `requirements.txt` | Python dependencies. |

## IDC Delivery Report (a sheet in the same workbook)

**Normal use:** on the **MMU Daily Report** page, upload the IDC DPR and IDC DSS CSVs in the optional
"IDC Delivery files" section. The generated MMU workbook then gets one extra visible sheet, **IDC Report** (placed right after Divisional Summary)
(plus hidden `Src_IDC_DPR` / `Src_IDC_DSS`; the IDC checks are appended to the hidden Validation Log). It uses the
Range and Single Date already picked for the MMU files — no extra dates. Leave both uploads empty and the
workbook is exactly as before.

**IDC only:** sidebar → **IDC Delivery Report** builds a stand-alone IDC workbook. Inputs: the IDC DPR CSV (required) and IDC DSS CSV (optional), plus Range
From / To and a Single Date from the calendar pickers (never from filenames). Output: one `IDC Report` sheet
(+ 3 hidden sheets: sources and Validation Log), one row per delivery person, sorted by Delivery % (Range),
smallest first, with live SUMIFS formulas and cached values.

| KPI | Definition |
|---|---|
| Delivery % | (Invoiced − Deposited) / Invoiced — all products; Range and Single Date |
| DSS Usage % | DSS Invoiced Count / PDMS Invoiced Count; Range and Single Date |
| Parcel Delivery % | same as Delivery %, only products whose name contains "parcel" (Range) |
| RTS % / Redirected % | Returned / Invoiced, Redirected / Invoiced (Range) |

All are Σ/Σ (volume-weighted); blank, never 0, when the denominator is 0. Formatting: colour scale on Delivery %
and Parcel Delivery %, data bar on Parcels Invoiced, DSS < 100% red text (bold on the Range column), RTS /
Redirected > 5% dark-red text on light-red fill; the total row uses text colours (Delivery % red < 93,
amber 93–<95, green ≥ 95; RTS / Redirected red ≥ 5%, else green). Names are matched across the two files ignoring
case and extra spaces. Switching between the two pages clears the other page's uploaded files (Streamlit behaviour).

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

On first run, upload the Office Master File once via "Upload / Replace Master File" —
it's saved to `./data/master_file.xlsx` and auto-loads on every subsequent run without
re-uploading. An optional header logo can be persisted the same way at
`./data/logo.*`, overriding the default India Post logo shown in the banner.

## Deploy to Streamlit Community Cloud

1. Push this folder to a GitHub repo.
2. On [share.streamlit.io](https://share.streamlit.io), point a new app at `app.py`.
3. No `packages.txt` (system packages) needed — pure Python throughout.
4. ⚠️ **Persistence caveat**: Streamlit Community Cloud's filesystem persists while
   the app instance stays running, but resets on redeploy or a cold restart. For
   guaranteed permanence of the Master File across redeploys, either self-host on a
   server with a real persistent disk, or swap `engine.py`'s `save_master_file` /
   `load_persisted_master` for a small cloud-storage read/write (S3, Google Drive,
   etc.) — the rest of the app doesn't need to change.

## What matches the literal spec

- **9 named placeholders, each with its own calendar picker(s)**: Delivery
  Productivity Range-1 (from/to), Range-2 (from/to), Single Date; DSS Usage Range
  (from/to), Single Date; Letter Box Clearance Range (from/to), Single Date; COD
  Digital Transaction Range (from/to), Single Date. **Dates are only ever taken
  from these pickers — never inferred from uploaded filenames** (this matters:
  real daily files have arrived with inconsistent, sometimes misleading names).
- **Delivery Productivity Range-2** exists because the source system caps a single
  export at ~15 days — Range-1 and Range-2 are summed together per office, kept as
  two separate hidden mirror sheets (`Src_DP_R1`, `Src_DP_R2`) with Raw Data's
  formulas summing across both.
- **Any of the 9 files may be missing.** Clicking Generate with files missing shows
  a reminder listing exactly which ones, and asks "Still generate the report?" —
  only on explicit "Yes" does it proceed. Missing parameters render as true blanks
  everywhere, and defaulter tables for a missing parameter note "data not provided"
  instead of falsely implying zero defaulters.
- **A permanently hosted Master File**, with a "replace" affordance in the UI.
- **A live Divisional Summary widget** on the page itself, showing all 4 KPIs ×
  Range and Single Date, refreshed the moment files are uploaded.
- **India Post branding**: a maroon/gold GIGW-style header, National Emblem +
  India Post logo, bilingual subtitle text, replaceable logo slot.
- **v10 formatting** (see the spec for full detail): grouped 2-row headers with
  literal dates everywhere except Raw Data and Defaulters (which use flat headers
  with literal dates instead); thick dark border segregation between metric
  groups and around header/total blocks on the 3 sheets that use grouped headers;
  wrapped text on specific header/total rows only; centered Marathi message text
  on the two Excl BOs sheets; red-highlighted keywords (डिलीवरी%, DSS%, COD Digital
  Transaction %) within the Defaulters sheets' Marathi paragraphs; no
  record-number row on Defaulters sheets; real dates (not generic "(Range)"/
  "(Single Date)") in Exceptional Defaulters' Parameter column; all panes
  unfrozen except Raw Data (`G4`) and Exceptional Defaulters (`A4`).

## Validation performed before delivery (original v1 build — still applies; see "What's new in v2" above for the v2-specific checks)

- Engine output cross-checked against the three worked reference offices in Build
  Spec v9 §1.5 (Nade B.O, Dhoroshi B.O, Korti B.O) — exact match, including the
  blank-vs-real-zero distinction.
- Generated workbook's Raw Data values cross-checked against an independent
  pandas computation across all 288 offices × 8 percentage columns (2,304 cells)
  — zero mismatches.
- Same file's formulas independently recalculated via LibreOffice — 2,991
  formulas, zero errors.
- Structural spot-checks: all 19 sheets present; freeze panes correct (`G4` on Raw
  Data, `A4` on Exceptional Defaulters, unfrozen everywhere else); thick group
  borders present on Excl BOs - Combined's sub-label row; wrap-text confirmed on
  that same row and on Defaulters' header row; centered alignment confirmed on
  the Excl BOs Combined Marathi row; Defaulters sheets confirmed to have no
  record-number row; Exceptional Defaulters confirmed showing real dates in the
  Parameter column.
- App boot-tested end-to-end (`streamlit run app.py` against a pre-seeded
  persisted Master file) — served HTTP 200 with no exceptions in the server log.
- Confirmed `app.py`'s call site into `build_workbook()` still matches the
  current function signature exactly (no code changes needed there despite
  `workbook.py` having grown substantially since `app.py` was first written).
