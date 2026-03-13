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
| Round 5 | 28 | $23,283,728 | FY22-23 Q4 → FY24-25 Q4 | 09/30/2025 |
| Round 6 | 30 | $20,632,608 | FY23-24 Q4 → FY25-26 Q4 | 09/30/2026 |
| Round 7 | 38 | $21,549,421 | FY24-25 Q4 → FY26-27 Q4 | 09/30/2027 |
| Round 8 | 24 | $3,265,760 | FY25-26 Q4 → FY27-28 Q4 | 09/30/2028 |

**Column layout:**

R5 uses 17 columns (A-Q). R6-R8 use 19 columns (A-S) with two additional columns: Plan Status (L) and Lead LEA (M). Columns A-K are identical across all rounds.

| Column | R5 | R6-R8 | Description |
|--------|-----|-------|-------------|
| Lead Institution | A | A | K-12 school district or ROP leading the grant |
| Proposal ID | B | B | Unique NOVA proposal identifier |
| # Reporting Institutions | C | C | Count of institutions filing reports |
| Grant Name | D | D | Full project title |
| FY Quarterly columns (varies) | E-I | E-I | Report submission status — `TRUE` = submitted, blank = not submitted |
| Final Report | J | J | Final report submission status |
| Report Waiting Approval | K | K | Flags reports pending approval |
| Plan Status | — | L | Plan-level certification status from `/swpk/plans` (R6-R8 only) |
| Lead LEA | — | M | Lead Local Education Agency from `/swpk/plans` (R6-R8 only) |
| Grant Amount | L | N | Total grant award in dollars |
| Total Reported Expenditures | M | O | Sum of all expenditures reported to date |
| Total Reported Expenditures Approved | N | P | Sum of approved expenditures only |
| % Spent | O | Q | Ratio of expenditures to grant amount |
| Notes | P | R | Free-text notes |
| Unexpended according to last fiscal report | Q | S | Remaining unspent funds |

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

## 73 Central Mother Lode Institutions (across all rounds)

Amador County ROP, Amador County Unified, Aspire Vanguard College Preparatory Academy, be.tech, Burton Elementary, Calaveras Unified, Career Technical Education Charter, Ceres Unified, Chawanakee Unified, Clovis Unified, Coalinga-Huron Unified, Corcoran Joint Unified, Cutler-Orosi Joint Unified, Delhi Unified, El Nido Elementary, El Tejon Unified, Escalon Unified, Exeter Unified, Firebaugh-Las Deltas Unified, Fowler Unified, Fresno ROP, Fresno Unified, Fruitvale Elementary, Galt Joint Union High, Golden Plains Unified, Grow Public Schools, Hanford Joint Union High, Hilmar Unified, Inyo Co. Office of Education, Kern Co. Office of Education, Kern High, Kern High ROC, Kings County ROP, Kingsburg Joint Union High, Le Grand Union High, Lemoore Union High, Lincoln Unified, Linden Unified, Lodi Unified, Madera Unified, Manteca Unified, Mariposa County Unified, McFarland Unified, Merced City Elementary, Merced City School District, Merced Co. Office of Education, Merced County ROP, Merced Union High, Modesto City High, Newman-Crows Landing Unified, Oakdale Joint Unified, Patterson Joint Unified, Porterville Unified, Reef-Sunset Unified, Richland Union Elementary, Ripon Unified, Riverdale Joint Unified, San Joaquin Co. Office of Education, Sanger Unified, Sonora Union High, Stockton Unified, Taft Union High, Tehachapi Unified, Tulare Co. Office of Education, Tuolumne County Superintendent of Schools, Valley ROP, Visalia Technical Early College, Visalia Unified, Wasco Union High, Waterford Unified, Wonderful College Prep Academy, Yosemite ROP, Yosemite Unified

## Current Status & Known Issues

- **Rounds 1-3:** Fully closed. Final reports due and mostly submitted. High spend rates (92-97% overall).
- **Round 4:** Final reports were due 09/30/2024. ~47% spent overall — some grants lagging.
- **Round 5:** Final report due 09/30/2025 (passed). All 28 grants Certified. $23.3M budget, $461K spent (2.0%). All quarterly reports submitted. 25/28 Final Reports Complete (3 Incomplete: 20672, 20635, 20588). Scraped from NOVA 2026-03-11.
- **Round 6:** Active. $20.6M budget, $1.4M spent (6.8%). 25 of 30 grants Certified for FY23-24. FY24-25: 23/30 Q2 Complete, 18/30 Q4 Complete. FY25-26 quarterly reports not yet submitted. Scraped from NOVA 2026-03-11.
- **Round 7:** Active. $21.5M budget, $461K spent (2.1%). 33/38 FY24-25 Q4 Complete, 13/38 FY25-26 Q2 Complete. FY25-26 Q4 and FY26-27 not yet submitted. Scraped from NOVA 2026-03-13 (authoritative per-page Q4 values, replacing 2026-03-12 proxy values).
- **Round 8:** Early lifecycle. 24 grants (expanded from 4 via `/swpk/plans` discovery on 2026-03-12). 4 grants have fiscal data ($3.3M budget), 20 are plans-only (Submitted/Draft, $0 budget). Re-scraped 2026-03-13.
- **R6 Look** has 722 "Pending Approval" entries (44% of rows) — significant backlog of unapproved reports.

## Data Sources for Rounds 5-8

Rounds 5-8 were last updated on 2026-03-13 via direct NOVA scraping (R5/R6 data from 2026-03-11, R7/R8 re-scraped 2026-03-13):
- Data scraped from individual NOVA grant fiscal report pages (`nova.cccco.edu/swpk/fiscal-reports/plans/{planId}?duration={yearCode}`)
- Values are hardcoded (not formula-driven), except % Spent which uses a formula (R5: `=M{row}/L{row}`, R6-R8: `=O{row}/N{row}`)
- Quarterly submission status determined by visiting each FY page and extracting Q2/Q4/Final Report → Complete/Incomplete status
- Q2 and Q4 are independent per FY page — cannot use primary page approval status as a proxy
- "# Reporting Institutions" uses the dashboard X/Y submitted count (Y value)
- Final Report status (Col J) tracks per-row TRUE/blank based on Final Report → Complete/Incomplete from the last FY page
- Previous #REF! errors in R5/R6 "# Reporting Institutions" column have been fixed
- Rounds 1-4 remain unchanged (still use Look tab formula references)
