# K12 SWP Grant Tracking — Operational Runbook

## Project Overview

This project tracks **K-12 Strong Workforce Program (SWP)** grant fiscal reporting for the **Central Mother Lode** region. Data originates from the **NOVA system** (nova.cccco.edu) and is maintained in an Excel spreadsheet.

- **Spreadsheet:** `K12 Reporting Status - Central Mother Lode.xlsx`
- **Structure docs:** See `SPREADSHEET_CONTEXT.md` for detailed workbook structure
- **Update cadence:** A couple times per month to keep fiscal data current

The spreadsheet contains Round 1-8 tabs (one per funding round), Look tabs (bulk NOVA exports for Rounds 1-6), and supporting tabs.

---

## Repeatable Process — "Update Grant Data from NOVA"

### Phase 1: Prerequisites & Authentication

1. User must have **NOVA** (nova.cccco.edu) open and authenticated in Chrome
2. Verify auth by calling `tabs_context_mcp` and confirming a NOVA tab exists
3. Identify which rounds need updating — active rounds with final report deadlines not yet passed
4. Confirm with the user which rounds to update before proceeding

**Round Reference Table:**

| Round | Grant Award FY | NOVA Year Code | Final Report Due | Quarterly Column FYs |
|-------|---------------|----------------|-----------------|---------------------|
| R5 | 2022-23 | `2023004` | 09/30/2025 | 22-23, 23-24, 24-25 |
| R6 | 2023-24 | `2024004` | 09/30/2026 | 23-24, 24-25, 25-26 |
| R7 | 2024-25 | `2025004` | 09/30/2027 | 24-25, 25-26, 26-27 |
| R8 | 2025-26 | `2026004` | 09/30/2028 | 25-26, 26-27, 27-28 |

### Phase 2: Dashboard Scan — Get Grant Listing

The dashboard is the **authoritative source** for grant listings. Never rely solely on hardcoded plan IDs in `scrape_round.py`.

1. Navigate to the NOVA Fiscal Reporting Dashboard: `nova.cccco.edu/swpk/fiscal-reports/plans`
2. Filter by **region: Central/Mother Lode**
3. Filter by **fiscal year** matching the round's Grant Award FY (see Round Reference Table above):
   - R5 → filter by 2022-23
   - R6 → filter by 2023-24
   - R7 → filter by 2024-25
   - R8 → filter by 2025-26
4. Extract from the dashboard listing for each grant:
   - Plan ID
   - Grant Name
   - Lead Institution
   - Submitted count (X/Y format — Y = number of reporting institutions)
   - Approval Status
5. Record the **total grant count** per round — this is the **expected row count** used for validation
6. **Compare against `scrape_round.py`** plan IDs:
   - New grants not in `scrape_round.py` → add them, scrape them
   - Grants in `scrape_round.py` but missing from dashboard → warn, investigate
   - Update `scrape_round.py` with any new plan IDs after each run

**Dashboard filter limitation:** `form_input` can set text values but doesn't trigger Angular state changes for combobox dropdowns. Instead, use this approach:

1. Navigate to `nova.cccco.edu/swpk/fiscal-reports/plans` (wait 5s for SSO)
2. Call `get_page_text` to extract the full dashboard listing (all regions, all rounds)
3. The dashboard text contains repeating blocks per grant, each with:
   - Plan ID (numeric, in the URL or as text)
   - Lead Agency name
   - Project Title
   - Fiscal Year
   - Submitted count (e.g., "1/1 Submitted" or "0/3 Submitted")
   - Approval Status (Certified, Submitted, Awaiting Submittal, etc.)
4. Filter to Central/Mother Lode grants by matching Lead Agency against the 68-institution list in `SPREADSHEET_CONTEXT.md`
5. Filter to the target round's fiscal year
6. Parse out Plan IDs, institution counts, and approval statuses

**Note:** The dashboard may paginate. If fewer grants appear than expected, check for pagination controls or scroll triggers. The `get_page_text` output may need multiple calls if the page lazy-loads.

### Phase 3: Individual Grant Scraping

For each grant from the dashboard listing:

1. Navigate to: `nova.cccco.edu/swpk/fiscal-reports/plans/{planId}?duration={yearCode}`
2. **Year codes** (see Round Reference Table in Phase 1 for per-round mapping):
   - FY2022-23 = `2023004`
   - FY2023-24 = `2024004`
   - FY2024-25 = `2025004`
   - FY2025-26 = `2026004`
   - FY2026-27 = `2027004`
   - FY2027-28 = `2028004`
3. Wait **5 seconds** after navigation — NOVA SSO may redirect before auto-authenticating
4. Extract data via JavaScript on `document.body.innerText`:

```javascript
// Plan ID
const planId = window.location.href.match(/plans\/(\d+)/)?.[1];

// Grant Name
const grantLink = document.querySelector('a[href*="/swpk/plans/"][href*="/preview"]');
const grantName = grantLink?.textContent?.trim();

// Approval Status
const statusMatch = document.body.innerText.match(/Status\n(Certified|Approved|Pending Approval|Draft|Submitted|Awaiting Submittal|Not Started)/);
const approvalStatus = statusMatch?.[1] || '';

// Financial data — find ALL lines starting with "Totals\t"
const lines = document.body.innerText.split('\n');
const totalsLines = lines.filter(l => l.startsWith('Totals\t'));

// Filter to 7-column Expenditure Report rows (skip 5-column Financial Match rows)
const expenditureRows = totalsLines.filter(l => l.split('\t').length >= 7);

// Sum across all institutions (multi-institution grants have multiple rows)
let ptdExpenditure = 0, projectBudget = 0, budgetRemaining = 0;
for (const row of expenditureRows) {
  const cols = row.split('\t');
  // Column mapping: Totals | PTD Exp | PTD Forecast | % Forecast | Project Budget | % Budget | Budget Remaining
  ptdExpenditure += parseFloat(cols[1].replace(/[$,]/g, '')) || 0;
  projectBudget += parseFloat(cols[4].replace(/[$,]/g, '')) || 0;
  budgetRemaining += parseFloat(cols[6].replace(/[$,]/g, '')) || 0;
}

const institutionCount = expenditureRows.length;
```

5. **Extract quarterly submission status** — each grant spans 3 fiscal years. Visit each FY's page to determine which quarters have been submitted.

   **Key optimization:** The primary page from step 4 (`?duration={grantYearCode}`) IS the first FY page. Extract Q4 status from it immediately after the financial data — no extra navigation needed. Only navigate to the 2nd and 3rd FY pages separately.

   **Skip future FY pages:** Only scan FY pages where at least one quarterly deadline has passed. If both Q2 and Q4 deadlines are in the future, set both to FALSE without navigating. This saves significant time — e.g., for R7 in early 2026, only the first FY page (FY24-25) needs scanning; FY25-26 Q2 deadline isn't until 02/28/26.

   **Per-FY extraction — use this JavaScript on each FY page:**

   ```javascript
   const text = document.body.innerText;
   const lines = text.split('\n');
   const r = {};
   for (let i = 0; i < lines.length; i++) {
     const t = lines[i].trim();
     if (t === 'Q2' && i + 1 < lines.length) r.q2 = lines[i + 1].trim();
     if (t === 'Q4' && i + 1 < lines.length) r.q4 = lines[i + 1].trim();
     if (t === 'Final Report' && i + 1 < lines.length) r.final = lines[i + 1].trim();
   }
   r.planId = window.location.href.match(/plans\/(\d+)/)?.[1];
   JSON.stringify(r);
   ```

   - Status values: `"Complete"` = submitted (TRUE), `"Incomplete"` = not submitted (blank)
   - Q2 and Q4 are **independent** — a grant can have Q2=Complete but Q4=Incomplete on the same FY page
   - The primary page `approvalStatus` (Certified/Submitted/etc.) is NOT a reliable proxy for individual quarter status

   **FY page structure:**
   - First FY (grant award year): shows Q4 only — **extract on same page load as step 4**
   - Middle FY: shows Q2 and Q4
   - Last FY: shows Q2, Q4, and Final Report

   **Time expectation:** With optimization, ~2 extra pages per grant (not 3). For 38 R7 grants with 1 skippable FY = ~76 extra page loads at ~6s each = ~8 minutes per round.

   **Map to quarterly columns:**
   - Col E = Grant FY Q4 → extract from primary page (step 4) → Q4 status
   - Col F = Grant FY+1 Q2 → navigate to `?duration={nextYearCode}` → Q2 status
   - Col G = Grant FY+1 Q4 → same page as Col F → Q4 status
   - Col H = Grant FY+2 Q2 → navigate to `?duration={nextNextYearCode}` → Q2 status
   - Col I = Grant FY+2 Q4 → same page as Col H → Q4 status
   - Col J = Final Report → same page as Col H/I → Final Report status

6. Save all results to `scraped_data.json` with this structure per grant:

```json
{
  "planId": "29751",
  "grantName": "Agriculture Pathway Expansion",
  "leadInstitution": "Newman-Crows Landing Unified",
  "approvalStatus": "Certified",
  "ptdExpenditure": 79554,
  "projectBudget": 418567,
  "budgetRemaining": 339013,
  "dashboardInstitutions": 1,
  "dashboardApproval": "Approved",
  "quarterlyStatus": {
    "FY24-25_Q4": true,
    "FY25-26_Q2": false,
    "FY25-26_Q4": false,
    "FY26-27_Q2": false,
    "FY26-27_Q4": false
  },
  "finalReport": false
}
```

### Phase 4: Spreadsheet Update (openpyxl)

**Step 0 — Backup before modifying:**

Before making any changes to the spreadsheet, create a timestamped backup:
```bash
mkdir -p backups
cp "K12 Reporting Status - Central Mother Lode.xlsx" "backups/K12 Reporting Status - Central Mother Lode - $(date +%Y-%m-%d).xlsx"
```

**Step 1 — Run the regeneration script:**
```bash
python3 regenerate_spreadsheet.py
```

This script handles the full rebuild. See the file for implementation details. If the script needs modification (e.g., new rounds, changed column layout), edit `regenerate_spreadsheet.py` rather than writing a one-off script.

**Strategy (implemented in `regenerate_spreadsheet.py`):**
- For **new round tabs** not yet in the spreadsheet: Create new sheet
- For **existing round tabs** already in the spreadsheet: **Merge and recreate**:
  1. Read existing data keyed by Proposal ID (column B)
  2. Preserve values from columns P (Notes) and Q (Unexpended) into a lookup dict
  3. Delete old sheet
  4. Recreate with fresh scraped data
  5. Restore preserved P and Q values by matching Proposal ID
  6. For grants that existed before but are no longer in the scrape: log a warning (grant may have been removed from NOVA)

**Column layout (17 columns — matches all existing Round tabs):**

| Col | Header | Format | Data Source |
|-----|--------|--------|-------------|
| A | Lead Institution | Text | `leadInstitution` from scrape |
| B | Proposal ID | Text | `planId` from scrape |
| C | # Reporting Insitutions | Number (keep original typo) | `dashboardInstitutions` from dashboard |
| D | Grant Name | Text | `grantName` from scrape |
| E-I | 5 quarterly columns (see Quarterly Column Mapping) | Boolean | `quarterlyStatus` from scrape: `TRUE` if submitted, blank if not |
| J | Final Report (Due MM/DD/YYYY) | Boolean | `finalReport` from scrape: `TRUE` if final report submitted (Complete), blank if not |
| K | Report Waiting Approval | Text | Derived: if `dashboardApproval` is "Pending Approval" or "Submitted" → set to that status; otherwise blank |
| L | Grant Amount | Number, `#,##0` | `projectBudget` from scrape |
| M | Total Reported Expenditures | Number, `#,##0` | `ptdExpenditure` from scrape |
| N | Total Reported Expenditures Approved | Number, `#,##0` | **Approximation:** if `dashboardApproval` is "Approved" or "Certified" → same as M; if "Pending Approval" → 0 |
| O | % Spent | Formula `=M{row}/L{row}`, format `0.0%` | Calculated |
| P | Notes | Text | **Preserved from previous tab** via Proposal ID merge; blank for new grants |
| Q | Unexpended according to last fiscal report | Text | **Preserved from previous tab** via Proposal ID merge; blank for new grants |

**Formatting:**
- Font: Arial throughout
- Header row: Bold, white text on blue fill (`#4472C4`), wrap text enabled
- Data rows sorted alphabetically by Lead Institution (column A)
- Last row: "Central Mother Lode Region" total row with:
  - `SUM` formulas for Grant Amount (L), Total Reported Expenditures (M), Expenditures Approved (N)
  - `COUNTIF` for quarterly submission columns (E-I) AND Final Report (J)
  - `% Spent` formula referencing the totals

**Column widths:** A=28.38, B=10.88, C=10.88, D=59.0, E-J=7.75, K=8.75, L-N=11.75, O=8.63, P=32.38, Q=15.0

### Phase 5: Update SPREADSHEET_CONTEXT.md

After updating the spreadsheet:
1. Update the round summary table with current grant counts and budget totals
2. Update the institution list if new institutions appeared
3. Update the "Current Status & Known Issues" section with latest status

---

## Validation Checklist

All checks must pass before saving. Run these in order:

1. **Grant count match:** Scraped grant count per round must match dashboard listing count
2. **Plan ID verification:** Every scraped Plan ID must appear in the dashboard listing
3. **Budget sanity:** All `projectBudget` values > $0
4. **Expenditure bounds:** `0 <= ptdExpenditure <= projectBudget` (flag but don't reject if over 100%)
5. **Budget remaining:** `budgetRemaining` approximately equals `projectBudget - ptdExpenditure` (within $1 rounding)
6. **Spreadsheet row count:** Data rows per tab must match expected grant count
7. **Column header match:** Compare headers against the Quarterly Column Mapping table to ensure correct layout (do not compare against another Round tab, since all tabs may be recreated)
8. **Sum validation:** Print total budget and total expenditure per round; compare against previous run if available
9. **Spot check:** Re-visit 3 random grant pages in browser and compare all fields against scraped values
10. **Formula check:** Verify `% Spent` formula evaluates correctly for 5 random rows
11. **Data preservation check:** After recreating a tab, verify all previously-existing Proposal IDs still have their Notes (P) and Unexpended (Q) values restored. Print any grants that existed before but are missing from the new scrape.
12. **Quarterly status consistency:** For each grant, if `approvalStatus` is Certified/Approved/Submitted on the primary FY page, then Col E (grant FY Q4) should generally be TRUE. Flag any mismatches for manual review (note: this is a heuristic — Q2/Q4 status from the FY page is authoritative).
13. **Quarterly monotonicity:** Quarters should not skip — if a later quarter is TRUE, all earlier quarters should also be TRUE. Flag if e.g., Col G=TRUE but Col F=FALSE.
14. **Dashboard grant count vs scrape_round.py:** If dashboard returns more grants than `scrape_round.py` lists, flag as "new grants found — add to scrape_round.py". If fewer, flag as "grants removed from NOVA".
15. **Final report status consistency:** If all quarterly columns (E-I) are TRUE and the final report due date has passed, Col J should be TRUE. Flag if blank.
16. **Cross-round deduplication:** Verify no plan ID appears in multiple rounds (sanity check).
17. **Institution count sanity:** For multi-institution grants (instCount > 1), verify the institution count hasn't changed from the previous scrape. Flag significant changes.
18. **Stale data detection:** For grants with past-due quarterly deadlines where the quarterly column is FALSE, flag as "potentially stale — may need manual verification on NOVA".

---

## Known Edge Cases

- **NOVA SSO redirects:** Every navigation may redirect to `/login?returnUrl=...` — wait 5 seconds, the page auto-authenticates via SSO. Do not treat this as a failure.
- **Multi-institution grants:** Have multiple Expenditure Report tables (one per institution). Sum all 7-column Totals rows. The dashboard "X/Y Submitted" count gives the reliable institution count (Y value).
- **5-column vs 7-column Totals rows:** Some institutions render only 5 columns when budget columns are blank. Use the dashboard institution count as the reliable source, not the scraped row count.
- **$0 expenditure grants:** Normal for new/early-lifecycle grants. Do not flag as errors.
- **"Awaiting Submittal" grants:** Have no approval status text on the page — `approvalStatus` will be an empty string. This is expected.
- **Page not loaded:** If extraction returns an empty `planId` or `grantName`, the page hasn't finished loading. Wait 3 more seconds and retry once.
- **Rounds 5-6 transitioning from Look-tab formulas:** These rounds previously used QUERY formulas referencing Look tabs (R5 Look, R6 Look). The scrape-based approach replaces all formulas with hardcoded values. The Look tabs are preserved as historical archives but are no longer referenced by the Round tabs after the first scrape-based update.
- **#REF! errors in R5/R6:** The existing "# Reporting Institutions" column had #REF! errors from broken formula references. The scrape-based approach fixes this by using `dashboardInstitutions` as a hardcoded value.
- **Column N approximation:** "Total Reported Expenditures Approved" is approximated — the scrape cannot distinguish approved vs. unapproved expenditure line items. If the overall grant status is "Approved" or "Certified", column N equals column M. If "Pending Approval", column N is 0. This may slightly differ from the granular per-line approval data in the Look tabs.
- **Quarterly text patterns:** The NOVA page does NOT show "Q2 Approved" or "Q4 Submitted" as standalone text. Instead, each FY page shows "Q2" on one line followed by "Complete" or "Incomplete" on the next line, and similarly for "Q4" and "Final Report". Use the JavaScript extraction pattern in Phase 3 step 5.
- **Q2/Q4 are independent:** A grant can have Q2=Complete but Q4=Incomplete on the same FY page (e.g., grant 25760 on FY24-25). The primary page `approvalStatus` (Certified/Submitted) cannot be used as a proxy for individual quarter status.
- **"Unsubmitted" vs "Awaiting Submittal":** FY pages that haven't had reports submitted show "Status\nUnsubmitted" — this is distinct from "Awaiting Submittal" which appears on the primary page for grants that haven't filed their first report.
- **Dashboard filter limitation:** `form_input` can set text values but doesn't trigger Angular state changes for combobox dropdowns. Use `get_page_text` to extract the full listing and filter in Python.
- **Tab management:** Chrome extension tabs can disappear between sessions. Always call `tabs_context_mcp` before navigating and reuse existing authenticated tabs when available.
- **R7 Col E proxy values (2026-03-12):** R7 FY24-25_Q4 was set to TRUE for 33 grants based on `approvalStatus` (Certified/Submitted) as a shortcut — NOT from actual FY page Q4 status. This may contain errors since approvalStatus is not a reliable proxy (see "Q2/Q4 are independent" above). On the next R7 update, do a proper per-page scan of `?duration=2025004` for all 38 grants to get authoritative Q4 values.

---

## Quarterly Column Mapping

Pattern for determining the 5 quarterly columns (E-I) and Final Report column (J) per round:

| Round | Grant FY | Col E | Col F | Col G | Col H | Col I | Col J (Final) |
|-------|----------|-------|-------|-------|-------|-------|---------------|
| R5 | 2022-23 | FY22-23 Q4 (08/31/23) | FY23-24 Q2 (02/28/24) | FY23-24 Q4 (08/31/24) | FY24-25 Q2 (02/28/25) | FY24-25 Q4 (08/31/25) | 09/30/25 |
| R6 | 2023-24 | FY23-24 Q4 (08/31/24) | FY24-25 Q2 (02/28/25) | FY24-25 Q4 (08/31/25) | FY25-26 Q2 (02/28/26) | FY25-26 Q4 (08/31/26) | 09/30/26 |
| R7 | 2024-25 | FY24-25 Q4 (08/31/25) | FY25-26 Q2 (02/28/26) | FY25-26 Q4 (08/31/26) | FY26-27 Q2 (02/28/27) | FY26-27 Q4 (08/31/27) | 09/30/27 |
| R8 | 2025-26 | FY25-26 Q4 (08/31/26) | FY26-27 Q2 (02/28/27) | FY26-27 Q4 (08/31/27) | FY27-28 Q2 (02/28/28) | FY27-28 Q4 (08/31/28) | 09/30/28 |

**General pattern for Round N with Grant FY XX-YY:**
- Col E: FY XX-YY Q4 (08/31/YY)
- Col F: FY (XX+1)-(YY+1) Q2 (02/28/(YY+1))
- Col G: FY (XX+1)-(YY+1) Q4 (08/31/(YY+1))
- Col H: FY (XX+2)-(YY+2) Q2 (02/28/(YY+2))
- Col I: FY (XX+2)-(YY+2) Q4 (08/31/(YY+2))
- Col J: Final Report Due 09/30/(YY+3)

**NOVA URL year code:** `{end_year}004` (e.g., FY2024-25 = `2025004`)

---

## Context Window Management

A full update run involves 100+ browser navigations and can exceed the context window. Follow these practices to avoid losing progress mid-run:

1. **Save intermediate results to files.** After completing each round's scraping, write results to `scraped_data.json` immediately — don't accumulate everything in memory.
2. **Batch by round.** Complete one round fully (dashboard scan → per-grant scrape → update scraped_data.json) before starting the next.
3. **Use `regenerate_spreadsheet.py`** to rebuild the spreadsheet from `scraped_data.json` — this can run independently after all scraping is done, even in a new session.
4. **If compacted mid-run:** Read `scraped_data.json` to see what's already been saved. Check which grants in the current round still have placeholder/old quarterly values to determine where to resume.
5. **Combine extractions per page.** On each page load, extract ALL needed data (financial + quarterly status) in a single JavaScript call to minimize back-and-forth.

---

## File Inventory

| File | Purpose |
|------|---------|
| `K12 Reporting Status - Central Mother Lode.xlsx` | The spreadsheet (source of truth for grant tracking) |
| `SPREADSHEET_CONTEXT.md` | Structural documentation of the spreadsheet |
| `scraped_data.json` | Raw data from the last NOVA scrape (overwritten each run) |
| `CLAUDE.md` | This file (operational runbook for the update process) |
| `scrape_round.py` | Helper script with hardcoded plan IDs per round (R5-R8). Compare against dashboard each run. |
| `regenerate_spreadsheet.py` | Rebuilds Round 5-8 tabs from `scraped_data.json`. Preserves Notes/Unexpended columns. Run after scraping is complete. |
| `.gitignore` | Git ignore rules (excludes `backups/`, `.playwright-mcp/`, `__pycache__/`) |
| `backups/` | Timestamped spreadsheet backups created before each update (gitignored) |
