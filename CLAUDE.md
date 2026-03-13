# K12 SWP Grant Tracking — Operational Runbook

## Global Overrides

This is a Python data pipeline (browser scraping + openpyxl), not a frontend or API project. The following global CLAUDE.md workflows do **not apply**:

- **Frontend & Design** — no HTML/UI in this project
- **Security Workflow** — skip auth, input validation, and injection checks (no API, no auth, no user input). Only the "sensitive data exposure" check applies.
- **Testing Workflow** — no test framework configured. `validate.py` serves as the validation step instead.

Use **context7** for openpyxl API verification when modifying `regenerate_spreadsheet.py`.

---

## Project Overview

This project tracks **K-12 Strong Workforce Program (SWP)** grant fiscal reporting for the **Central Mother Lode** region. Data originates from the **NOVA system** (nova.cccco.edu) and is maintained in an Excel spreadsheet.

- **Spreadsheet:** `K12 Reporting Status - Central Mother Lode.xlsx`
- **Structure docs:** See `SPREADSHEET_CONTEXT.md` for detailed workbook structure
- **Update cadence:** A couple times per month to keep fiscal data current

The spreadsheet contains Round 1-8 tabs (one per funding round), Look tabs (bulk NOVA exports for Rounds 1-6), and supporting tabs.

---

## Browser Automation — Tool Detection & Fallback

This process supports two browser automation backends. **Claude-in-Chrome is preferred** (uses existing Chrome session with SSO). Playwright MCP is the fallback.

### Detection (run at start of every session)

1. Try `tabs_context_mcp`. If it succeeds → use **Claude-in-Chrome** for this session.
2. If `tabs_context_mcp` is not available or errors → try `browser_tabs(action: "list")`.
   - If that succeeds → use **Playwright MCP** for this session.
   - If that also fails → neither backend is configured. Stop and ask the user to set one up.
3. Set a mental flag for which backend is active. All subsequent tool calls use that backend's names.

### Tool Mapping

| Purpose | Claude-in-Chrome | Playwright MCP | Notes |
|---------|-----------------|----------------|-------|
| Detect availability | `tabs_context_mcp` | `browser_tabs(action: "list")` | Try chrome first, fall back to playwright |
| Navigate | `navigate(url, tabId)` | `browser_navigate(url)` | Playwright has no tabId param — operates on active tab |
| Get page text | `get_page_text(tabId)` | `browser_evaluate(() => document.body.innerText)` | No direct equivalent in Playwright |
| Execute JS | `javascript_tool(text, tabId)` | `browser_evaluate(function)` | Playwright requires arrow fn wrapper: `() => { ... }` |
| Form input | `form_input(ref, value, tabId)` | `browser_fill_form(fields)` | Different ref format (playwright uses snapshot refs) |
| Read page structure | `read_page(tabId)` | `browser_snapshot()` | Playwright returns accessibility tree |
| Create tab | `tabs_create_mcp` | `browser_tabs(action: "new")` | |
| Wait | `computer(action: "wait", duration, tabId)` | `browser_wait_for(...)` | |
| Install browser | N/A | `browser_install()` | Playwright-only; needed if browser binary missing |

### JavaScript Extraction Syntax

The JS extraction blocks in Phase 3 use raw code strings. Adapt per backend:

**Claude-in-Chrome** (`javascript_tool`):
- Pass JS as a plain string in the `text` parameter
- The result of the last expression is returned automatically
- Example: `"document.body.innerText"`

**Playwright MCP** (`browser_evaluate`):
- Wrap JS in an arrow function in the `function` parameter
- Must explicitly return the value
- Example: `"() => { return document.body.innerText; }"`

For the multi-line extraction blocks (financial data, quarterly status), wrap the entire block:
- Chrome: pass as-is in `text`
- Playwright: wrap in `"() => { <existing code>; return JSON.stringify(r); }"`

### SSO Authentication Difference

- **Chrome:** User logs into NOVA in their normal Chrome browser before starting. SSO cookies persist.
- **Playwright:** Launches a separate browser window. On first navigation to NOVA, the user must manually log in within that Playwright browser window. The session persists for the duration of the scraping run but is lost when the browser closes.

**Playwright first-run extra step:** After detection picks Playwright, navigate to `nova.cccco.edu` and wait. If a login page appears, tell the user: "Please log into NOVA in the Playwright browser window that just opened. Let me know when you're logged in." Then proceed with the pre-flight grant page check.

---

## Repeatable Process — "Update Grant Data from NOVA"

### Phase 1: Prerequisites & Authentication

1. User must have **NOVA** (nova.cccco.edu) accessible — either logged in via Chrome or ready to log in via Playwright.
2. **Detect browser backend** — follow the "Browser Automation — Tool Detection & Fallback" section above. Record which backend (Chrome or Playwright) is active for this session.
3. **Pre-flight SSO check:**
   - **Chrome path:** Call `tabs_context_mcp`, confirm a NOVA tab exists.
   - **Playwright path:** Call `browser_navigate` to `nova.cccco.edu`. If a login page appears, ask the user to authenticate in the Playwright browser window. Wait for confirmation.
4. **Pre-flight grant page check:** Navigate to one known grant page (e.g., `nova.cccco.edu/swpk/fiscal-reports/plans/29751?duration=2025004`) and extract `planId` via JavaScript. If empty after 5 seconds, SSO is not active — ask the user to log in manually and retry.
5. Verify `scraped_data.json` exists (to preserve previous round data during partial updates). If missing, create an empty JSON structure `{}` so the script can start fresh.
6. Identify which rounds need updating — active rounds with final report deadlines not yet passed. For R7 and later rounds, `/swpk/plans` is used as the primary discovery source (see Phase 2b).
7. Confirm with the user which rounds to update before proceeding

**Round Reference Table:**

| Round | Grant Award FY | NOVA Year Code | Final Report Due | Quarterly Column FYs |
|-------|---------------|----------------|-----------------|---------------------|
| R5 | 2022-23 | `2023004` | 09/30/2025 | 22-23, 23-24, 24-25 |
| R6 | 2023-24 | `2024004` | 09/30/2026 | 23-24, 24-25, 25-26 |
| R7 | 2024-25 | `2025004` | 09/30/2027 | 24-25, 25-26, 26-27 |
| R8 | 2025-26 | `2026004` | 09/30/2028 | 25-26, 26-27, 27-28 |

### Phase 2: Dashboard Scan — Get Grant Listing

The dashboard is the **authoritative source** for grant listings with fiscal data (R5-R6, and grants that have started reporting in R7+). For R7 and later rounds, `/swpk/plans` is the primary discovery source (see Phase 2b). Never rely solely on hardcoded plan IDs in `scrape_round.py`.

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
4. Filter to Central/Mother Lode grants by matching Lead Agency against the institution list in `SPREADSHEET_CONTEXT.md`
5. Filter to the target round's fiscal year
6. Parse out Plan IDs, institution counts, and approval statuses

**Note:** The dashboard may paginate. If fewer grants appear than expected, check for pagination controls or scroll triggers. The `get_page_text` output may need multiple calls if the page lazy-loads.

**Plan Status & Lead LEA extraction (R6-R8 only):**

After the dashboard scan, extract Plan Status and Lead LEA from the NOVA plans listing:

1. Navigate to `nova.cccco.edu/swpk/plans`
2. Extract full page text via `get_page_text` (or `browser_evaluate(() => document.body.innerText)` for Playwright)
3. The plans listing contains blocks per plan with Plan Status (Certified/Submitted/blank) and Lead LEA
4. For each Plan ID from the dashboard scan, match and extract:
   - `planStatus`: the plan-level certification status
   - `leadLEA`: the lead Local Education Agency name
5. Store as `planStatus` and `leadLEA` fields in each grant's scraped data
6. These fields are optional — R5 grants will not have them (`.get()` with empty string default handles this)

### Phase 2b: Plans Listing Scan (R7+ only)

The fiscal reports dashboard only shows grants that have started reporting. For R7 and later rounds, many plans exist in Submitted/Draft state before fiscal reporting begins. Use `/swpk/plans` as the primary discovery source.

1. Navigate to `nova.cccco.edu/swpk/plans`
2. The page shows all plans (paginated, 100 per page). Extract plan data by parsing the tab-separated page text:
   - Each row: `ID\tPathway Improvement\tLead LEA\tRegion\tAllocation Year\tStatus\tActions`
   - Filter for `Region = "Central/Mother Lode"` AND `Allocation Year` matching the round's grant FY
3. Paginate through all pages using the `#qa_pagination_next` button (via JS `.click()`) to collect all matching plans
4. This is the **authoritative plan list** — use it instead of hardcoded `scrape_round.py` plan IDs
5. Cross-reference with fiscal reports dashboard:
   - Plans on both → full fiscal data + plans metadata
   - Plans only on `/swpk/plans` → attempt fiscal page, fall back to zero-value entry with $0 budget, $0 expenditure, blank quarterly status
6. Update `scrape_round.py` planIds to match discovered list

**Pagination approach:** The Angular filter dropdowns are unreliable via browser automation (po-select components cause tab detachment). Instead, extract the full unfiltered listing page by page and filter in code. Use `document.body.innerText` to get tab-separated rows, parse with regex, and collect Central/Mother Lode entries for the target allocation years. Auto-paginate by calling `document.querySelector('#qa_pagination_next').click()` via JavaScript with 2-3 second delays between pages.

### Phase 3: Individual Grant Scraping

For each grant from the combined listing (dashboard + Phase 2b discoveries for R7+):

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
     // Use first-match-wins to avoid multi-institution pages overwriting with wrong value
     if (t === 'Q2' && i + 1 < lines.length && !r.q2) r.q2 = lines[i + 1].trim();
     if (t === 'Q4' && i + 1 < lines.length && !r.q4) r.q4 = lines[i + 1].trim();
     if (t === 'Final Report' && i + 1 < lines.length && !r.final) r.final = lines[i + 1].trim();
   }
   r.planId = window.location.href.match(/plans\/(\d+)/)?.[1];
   // Validate extracted values are expected
   if (r.q2 && r.q2 !== 'Complete' && r.q2 !== 'Incomplete') r.q2_warning = r.q2;
   if (r.q4 && r.q4 !== 'Complete' && r.q4 !== 'Incomplete') r.q4_warning = r.q4;
   if (r.final && r.final !== 'Complete' && r.final !== 'Incomplete') r.final_warning = r.final;
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

6. **Error handling per grant:**
   - After JavaScript extraction, verify `planId` is not empty. If empty, the page hasn't loaded — wait 3 seconds and retry extraction once.
   - If retry still returns empty `planId`, log the grant's plan ID to `failed_grants.txt` and continue to the next grant.
   - After completing all grants in a round, report the `failed_grants.txt` count. Re-attempt failed grants before moving to the next round.
   - If a grant fails twice, skip it and flag it in the final validation output for manual review.

7. Save all results to `scraped_data.json` with this structure per grant:

```json
{
  "planId": "29751",
  "grantName": "Agriculture Pathway Expansion",
  "leadInstitution": "Newman-Crows Landing Unified",
  "approvalStatus": "Certified",
  "planStatus": "Certified",
  "leadLEA": "Newman-Crows Landing Unified",
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
  2. Preserve Notes and Unexpended values into a lookup dict (found by header text, not fixed column index)
  3. Delete old sheet
  4. Recreate with fresh scraped data
  5. Restore preserved Notes and Unexpended values by matching Proposal ID
  6. For grants that existed before but are no longer in the scrape: log a warning (grant may have been removed from NOVA)

**Column layout — R5 has 17 columns (A-Q); R6-R8 have 19 columns (A-S) with Plan Status and Lead LEA.**

Columns A-K are identical across all rounds:

| Col | Header | Format | Data Source |
|-----|--------|--------|-------------|
| A | Lead Institution | Text | `leadInstitution` from scrape |
| B | Proposal ID | Text | `planId` from scrape |
| C | # Reporting Insitutions | Number (keep original typo) | `dashboardInstitutions` from dashboard |
| D | Grant Name | Text | `grantName` from scrape |
| E-I | 5 quarterly columns (see Quarterly Column Mapping) | Boolean | `quarterlyStatus` from scrape: `TRUE` if submitted, blank if not |
| J | Final Report (Due MM/DD/YYYY) | Boolean | `finalReport` from scrape: `TRUE` if final report submitted (Complete), blank if not |
| K | Report Waiting Approval | Text | Derived: if `dashboardApproval` is "Pending Approval" or "Submitted" → set to that status; otherwise blank |

R6-R8 only (not present in R5):

| Col | Header | Format | Data Source |
|-----|--------|--------|-------------|
| L | Plan Status | Text | `planStatus` from `/swpk/plans` listing (Certified/Submitted/blank) |
| M | Lead LEA | Text | `leadLEA` from `/swpk/plans` listing |

Remaining columns (letter shifts for R6-R8 vs R5):

| R5 Col | R6-R8 Col | Header | Format | Data Source |
|--------|-----------|--------|--------|-------------|
| L | N | Grant Amount | Number, `#,##0` | `projectBudget` from scrape |
| M | O | Total Reported Expenditures | Number, `#,##0` | `ptdExpenditure` from scrape |
| N | P | Total Reported Expenditures Approved | Number, `#,##0` | **Approximation:** if `dashboardApproval` is "Approved" or "Certified" → same as Total Exp; if "Pending Approval" → 0 |
| O | Q | % Spent | Formula, format `0.0%` | R5: `=IF(L{row}=0,"",M{row}/L{row})`; R6-R8: `=IF(N{row}=0,"",O{row}/N{row})` |
| P | R | Notes | Text | **Preserved from previous tab** via Proposal ID merge; blank for new grants |
| Q | S | Unexpended according to last fiscal report | Text | **Preserved from previous tab** via Proposal ID merge; blank for new grants |

**Formatting:**
- Font: Arial throughout
- Header row: Bold, white text on blue fill (`#4472C4`), wrap text enabled
- Data rows sorted alphabetically by Lead Institution (column A)
- Last row: "Central Mother Lode Region" total row with:
  - `SUM` formulas for Grant Amount, Total Reported Expenditures, Expenditures Approved
  - `COUNTIF` for quarterly submission columns (E-I) AND Final Report (J)
  - `% Spent` formula referencing the totals

**Column widths (R5):** A=28.38, B=10.88, C=10.88, D=59.0, E-J=7.75, K=8.75, L-N=11.75, O=8.63, P=32.38, Q=15.0
**Column widths (R6-R8):** A=28.38, B=10.88, C=10.88, D=59.0, E-J=7.75, K=8.75, L=8.75, M=28.38, N-P=11.75, Q=8.63, R=32.38, S=15.0

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
11. **Data preservation check:** After recreating a tab, verify all previously-existing Proposal IDs still have their Notes and Unexpended values restored (column letters differ per layout: P/Q for R5, R/S for R6-R8). Print any grants that existed before but are missing from the new scrape.
12. **Quarterly status consistency:** For each grant, if `approvalStatus` is Certified/Approved/Submitted on the primary FY page, then Col E (grant FY Q4) should generally be TRUE. Flag any mismatches for manual review (note: this is a heuristic — Q2/Q4 status from the FY page is authoritative).
13. **Quarterly monotonicity (cross-FY only):** If any quarter in a later FY is TRUE, all quarters in earlier FYs should also be TRUE. Flag if e.g., FY25-26 has a TRUE but FY24-25 has a FALSE. **Within a single FY, Q2 and Q4 are independent** — do NOT flag Q2=TRUE/Q4=FALSE on the same FY as an error (see "Q2/Q4 are independent" edge case).
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
- **Expenditures Approved approximation:** "Total Reported Expenditures Approved" is approximated — the scrape cannot distinguish approved vs. unapproved expenditure line items. If the overall grant status is "Approved" or "Certified", Approved equals Total Expenditures. If "Pending Approval", Approved is 0. This may slightly differ from the granular per-line approval data in the Look tabs. (Col N in R5, Col P in R6-R8.)
- **Quarterly text patterns:** The NOVA page does NOT show "Q2 Approved" or "Q4 Submitted" as standalone text. Instead, each FY page shows "Q2" on one line followed by "Complete" or "Incomplete" on the next line, and similarly for "Q4" and "Final Report". Use the JavaScript extraction pattern in Phase 3 step 5.
- **Q2/Q4 are independent:** A grant can have Q2=Complete but Q4=Incomplete on the same FY page (e.g., grant 25760 on FY24-25). The primary page `approvalStatus` (Certified/Submitted) cannot be used as a proxy for individual quarter status.
- **"Unsubmitted" vs "Awaiting Submittal":** FY pages that haven't had reports submitted show "Status\nUnsubmitted" — this is distinct from "Awaiting Submittal" which appears on the primary page for grants that haven't filed their first report.
- **Dashboard filter limitation:** `form_input` can set text values but doesn't trigger Angular state changes for combobox dropdowns. Use `get_page_text` to extract the full listing and filter in Python.
- **Tab management:** Chrome extension tabs can disappear between sessions. Always call `tabs_context_mcp` before navigating and reuse existing authenticated tabs when available.
- **R7 Col E proxy values (2026-03-12):** R7 FY24-25_Q4 was set to TRUE for 33 grants based on `approvalStatus` (Certified/Submitted) as a shortcut — NOT from actual FY page Q4 status. This may contain errors since approvalStatus is not a reliable proxy (see "Q2/Q4 are independent" above). On the next R7 update, do a proper per-page scan of `?duration=2025004` for all 38 grants to get authoritative Q4 values.
- **Playwright vs Chrome tool differences:** Playwright's `browser_evaluate` requires JS wrapped in an arrow function (`() => { ... }`), unlike Chrome's `javascript_tool` which takes raw code. Playwright also has no `get_page_text` — use `browser_evaluate(() => document.body.innerText)` instead. The detection step in "Browser Automation" section determines which syntax to use for the session.
- **Playwright SSO timeout:** Playwright browser sessions can time out during long scraping runs (100+ pages). If extraction starts returning empty results or login pages, the SSO session has expired. Ask the user to re-authenticate in the Playwright window, then resume from `failed_grants.txt` or the last saved grant in `scraped_data.json`.
- **Plans-only grants (no fiscal data):** Grants discovered from `/swpk/plans` but not on the fiscal reports dashboard have $0 budget, $0 expenditure, blank quarterly status, and blank approval status. `leadInstitution` uses Lead LEA from plans listing. `grantName` uses Pathway Improvement from plans listing. These grants appear in the spreadsheet with empty/zero values — the `% Spent` formula handles division by zero via `=IF(N{row}=0,"",...)`.
- **Plans listing pagination:** The `/swpk/plans` page uses Angular po-select dropdowns for filters, which don't respond reliably to programmatic interaction (tab detachment issue). Instead, scrape the full unfiltered listing (all 1800+ plans) page by page and filter in code. The pagination button `#qa_pagination_next` is clickable via JavaScript `.click()`. Use 2-3 second delays between pages. The table data is tab-separated in `document.body.innerText`.

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
2. **Incremental saves within a round.** After scraping every 10 grants, write intermediate results to `scraped_data.json`. Use Python to maintain a running dict — read existing JSON, update the current round's grants list, write back. This ensures progress survives context compaction.
3. **Batch by round.** Complete one round fully (dashboard scan → per-grant scrape → update scraped_data.json) before starting the next.
4. **Use `regenerate_spreadsheet.py`** to rebuild the spreadsheet from `scraped_data.json` — this can run independently after all scraping is done, even in a new session.
5. **If compacted mid-run:** Read `scraped_data.json` to see what's already been saved. For the current round, check which grants have `quarterlyStatus` with all-false values AND `ptdExpenditure` of 0 — these likely haven't been scraped yet (unless they're genuinely new/empty grants). Cross-reference against the dashboard listing to identify which grants still need scraping. Resume from the first unscraped grant.
6. **Combine extractions per page.** On each page load, extract ALL needed data (financial + quarterly status) in a single JavaScript call to minimize back-and-forth.

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
| `validate.py` | Automated validation checks on scraped data and spreadsheet. Run after regeneration. |
| `requirements.txt` | Python dependencies (openpyxl). Install with `pip install -r requirements.txt`. |
| `README.md` | Setup and usage guide for new users. |
| `.gitignore` | Git ignore rules (excludes `backups/`, `.playwright-mcp/`, `__pycache__/`) |
| `backups/` | Timestamped spreadsheet backups created before each update (gitignored) |
