# K12 Reporting Status - Central Mother Lode.xlsx

## Purpose

This workbook tracks **K-12 Strong Workforce Program (SWP)** grant reporting status for the **Central Mother Lode** region. It monitors quarterly fiscal report submissions, expenditures, and approval status for grants awarded across 8 funding rounds spanning FY2019-20 through FY2027-28. The data originates from the **NOVA system** (nova.cccco.edu), the California Community Colleges Chancellor's Office reporting platform.

## Workbook Structure

The workbook has **16 sheets** organized into 3 categories:

### 1. Round Tabs (8 sheets) — The Core Tracking Sheets

Each "Round" tab is a tracker for one funding round's grants **filtered to Central Mother Lode institutions only**. These are the primary working sheets.

| Sheet | Grants | Grant Total | FY Span | Final Report Due |
|-------|--------|-------------|---------|-----------------|
| Round 1 | 30 | $44,742,940 | FY19-20 Q4 → FY21-22 Q2 | 03/31/2022 |
| Round 2 | 26 | $41,273,830 | FY20-21 Q2 → FY22-23 Q2 | 03/31/2023 |
| Round 3 | 31 | $41,693,920 | FY20-21 Q4 → FY22-23 Q4 | 09/30/2023 |
| Round 4 | 22 | $36,011,642 | FY21-22 Q4 → FY23-24 Q4 | 09/30/2024 |
| Round 5 | 28 | $136,156,838 | FY22-23 Q4 → FY24-25 Q4 | 09/30/2025 |
| Round 6 | 30 | $42,715,216 | FY23-24 Q4 → FY25-26 Q4 | 09/30/2026 |
| Round 7 | 38 | $18,786,467 | FY24-25 Q4 → FY26-27 Q4 | 09/30/2027 |
| Round 8 | 4 | $3,265,760 | FY25-26 Q4 → FY27-28 Q4 | 09/30/2028 |

**Column layout (consistent across all Round tabs):**

| Column | Description |
|--------|-------------|
| Lead Institution | K-12 school district or ROP leading the grant |
| Proposal ID | Unique NOVA proposal identifier |
| # Reporting Institutions | Count of institutions filing reports (has #REF! errors in Rounds 5 & 6) |
| Grant Name | Full project title |
| FY Quarterly columns (varies) | Report submission status — `TRUE` = submitted, blank = not submitted |
| Final Report | Final report submission status |
| Report Waiting Approval | Flags reports pending approval |
| Grant Amount | Total grant award in dollars |
| Total Reported Expenditures | Sum of all expenditures reported to date |
| Total Reported Expenditures Approved | Sum of approved expenditures only |
| % Spent | Ratio of expenditures to grant amount |
| Notes | Free-text notes |
| Unexpended according to last fiscal report | Remaining unspent funds |

**Last row** in each Round tab contains a **region total** row (labeled "Central / Mother Lode Region" or "Central Mother Lode Region").

### 2. Look Tabs (6 sheets: R1 Look through R6 Look) — Raw NOVA Data

These are **bulk data exports from NOVA** containing fiscal report data for **ALL regions statewide** (not just Central Mother Lode). They serve as the lookup/source data that feeds the Round tabs.

**Each Look tab has 10 columns:**

| Column | Description |
|--------|-------------|
| Plan Details Proposal ID | Matches to Round tab Proposal ID |
| Plan Details Lead Agency | Institution name |
| Plan Details Project Title | Full grant title |
| Fiscal Reports Reporting Period | e.g., "2024-25 Quarter 4", "Final" |
| Fiscal Reports Fiscal Report Approve | "Approved" or "Pending Approval" |
| Fiscal Reports Fiscal Report Submit | "Submitted" or "Draft" |
| Budgets & Expenditures Grant Funds Expenditure | Dollar amount reported for that period |
| Fiscal Reports Count Submitted | Number of sub-reports submitted |
| Fiscal Reports Percent Submitted | Percentage of expected sub-reports submitted |
| Allocations Allocations | Allocation amount (often 0) |

**Row counts per Look tab:**

| Sheet | Total Rows | Unique Agencies | Unique Proposals |
|-------|-----------|----------------|-----------------|
| R1 Look | 2,617 | 169 | 241 |
| R2 Look | 2,106 | 177 | 234 |
| R3 Look | 1,846 | 202 | 263 |
| R4 Look | 1,478 | 174 | 211 |
| R5 Look | 1,575 | 189 | 224 |
| R6 Look | 1,636 | 194 | 232 |

**Key relationship:** Every Proposal ID in a Round tab exists in its corresponding Look tab. The Round tabs filter the Look data down to ~25-30 Central Mother Lode grants from the 200+ statewide proposals.

### 3. Supporting Tabs

**Reporting Deadlines:** Links to key resources (CCCCO K12 SWP page, CDE CTEIG page, NOVA login, Cal-PASS PLUS) plus a timeline of quarterly reporting deadlines and CTEIG milestones from Jan 2021 through late 2023.

**All Regions:** Simple list of the 8 SWP regions:
- Bay Area
- Central Mother Lode
- Inland Empire
- Los Angeles
- North Far North
- Orange County
- San Diego Imperial
- South Central Coast

## How the Data Flows

1. **NOVA** (nova.cccco.edu) is the source system where grantees submit fiscal reports
2. **Look tabs** are bulk exports from NOVA — one per round, all regions, all reporting periods
3. **Round tabs** are manually curated trackers that reference the Look data for Central Mother Lode grants only
4. The Round tabs track whether each quarterly report has been submitted (TRUE/blank) and summarize financials (grant amount, expenditures, % spent)

## 68 Central Mother Lode Institutions (across all rounds)

Amador County ROP, Amador County Unified, Aspire Vanguard College Preparatory Academy, be.tech, Burton Elementary, Calaveras Unified, Career Technical Education Charter, Ceres Unified, Chawanakee Unified, Clovis Unified, Coalinga-Huron Unified, Corcoran Joint Unified, Cutler-Orosi Joint Unified, Delhi Unified, El Nido Elementary, El Tejon Unified, Escalon Unified, Exeter Unified, Firebaugh-Las Deltas Unified, Fowler Unified, Fresno Unified, Fruitvale Elementary, Galt Joint Union High, Golden Plains Unified, Hanford Joint Union High, Hilmar Unified, Inyo Co. Office of Education, Kern Co. Office of Education, Kern High, Kern High ROC, Kings County ROP, Le Grand Union High, Lemoore Union High, Lincoln Unified, Linden Unified, Lodi Unified, Madera Unified, Manteca Unified, Mariposa County Unified, McFarland Unified, Merced City Elementary, Merced City School District, Merced Co. Office of Education, Merced County ROP, Merced Union High, Modesto City High, Newman-Crows Landing Unified, Oakdale Joint Unified, Patterson Joint Unified, Porterville Unified, Reef-Sunset Unified, Ripon Unified, Riverdale Joint Unified, San Joaquin Co. Office of Education, Sanger Unified, Sonora Union High, Stockton Unified, Taft Union High, Tulare Co. Office of Education, Tuolumne County Superintendent of Schools, Valley ROP, Visalia Technical Early College, Visalia Unified, Wasco Union High, Waterford Unified, Wonderful College Prep Academy, Yosemite ROP, Yosemite Unified

## Current Status & Known Issues

- **Rounds 1-3:** Fully closed. Final reports due and mostly submitted. High spend rates (92-97% overall).
- **Round 4:** Final reports were due 09/30/2024. ~47% spent overall — some grants lagging.
- **Round 5:** Active. Reports through FY24-25 Q2 due. ~16% spent — still early in grant lifecycle.
- **Round 6:** Most current. FY25-26 reports in progress. ~24% spent. Many reports still pending approval.
- **#REF! errors** in "# Reporting Institutions" column for Rounds 5 and 6 — likely broken formula references.
- **R6 Look** has 722 "Pending Approval" entries (44% of rows) — significant backlog of unapproved reports.
- **Round 7:** Added 2026-03-11. 38 grants, $18.8M total budget, 2.45% spent. Data scraped directly from NOVA individual grant fiscal report pages. Most grants have $0 expenditure — early in lifecycle. 5 grants are "Awaiting Submittal", 2 are "Pending Approval", 31 are "Approved/Certified". No Look tab — data populated directly from NOVA scrape.
- **Round 8:** Added 2026-03-11. 4 grants, $3.3M total budget, $0 spent. All grants are "Awaiting Submittal" — very early in grant lifecycle. No Look tab — data populated directly from NOVA scrape.

## Data Sources for Round 7 & 8

Rounds 7 and 8 were populated differently from Rounds 1-6:
- **Rounds 1-6:** Use Look tabs (bulk NOVA exports) as source data, with Google Sheets QUERY formulas in the Round tabs
- **Rounds 7-8:** Data scraped directly from individual NOVA grant fiscal report pages (`nova.cccco.edu/swpk/fiscal-reports/plans/{planId}?duration={yearCode}`) on 2026-03-11. Values are hardcoded (not formula-driven). Quarterly submission columns are blank pending future data pulls. The "# Reporting Institutions" column uses the count from the NOVA dashboard listing rather than formula references.
