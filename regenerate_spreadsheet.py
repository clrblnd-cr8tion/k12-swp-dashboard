#!/usr/bin/env python3
"""
Regenerate Round 5-8 tabs in the K12 Reporting Status spreadsheet
from scraped_data.json. Preserves Notes (P) and Unexpended (Q) columns
by matching on Proposal ID.

Usage: python3 regenerate_spreadsheet.py
"""
import json
import os
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from utils import SPREADSHEET, SCRAPED_DATA, is_data_row

# Column indices (1-based, openpyxl convention)
COL_QUARTERLY_START = 5   # E — first quarterly column
COL_FINAL_REPORT = 10     # J
COL_WAITING = 11          # K
COL_GRANT_AMOUNT = 12     # L
COL_TOTAL_EXP = 13        # M
COL_APPROVED_EXP = 14     # N
COL_PCT_SPENT = 15        # O
COL_NOTES = 16            # P
COL_UNEXPENDED = 17       # Q

# Round configurations: sheet name, quarterly column headers, quarterly keys, final report header
ROUND_CONFIGS = {
    'R5': {
        'sheet_name': 'Round 5',
        'quarterly_labels': [
            'FY22-23 Q4\n(08/31/23)', 'FY23-24 Q2\n(02/28/24)',
            'FY23-24 Q4\n(08/31/24)', 'FY24-25 Q2\n(02/28/25)',
            'FY24-25 Q4\n(08/31/25)'
        ],
        'quarterly_keys': [
            'FY22-23_Q4', 'FY23-24_Q2', 'FY23-24_Q4',
            'FY24-25_Q2', 'FY24-25_Q4'
        ],
        'final_report_header': 'Final Report\n(Due 09/30/25)',
    },
    'R6': {
        'sheet_name': 'Round 6',
        'quarterly_labels': [
            'FY23-24 Q4\n(08/31/24)', 'FY24-25 Q2\n(02/28/25)',
            'FY24-25 Q4\n(08/31/25)', 'FY25-26 Q2\n(02/28/26)',
            'FY25-26 Q4\n(08/31/26)'
        ],
        'quarterly_keys': [
            'FY23-24_Q4', 'FY24-25_Q2', 'FY24-25_Q4',
            'FY25-26_Q2', 'FY25-26_Q4'
        ],
        'final_report_header': 'Final Report\n(Due 09/30/26)',
    },
    'R7': {
        'sheet_name': 'Round 7',
        'quarterly_labels': [
            'FY24-25 Q4\n(08/31/25)', 'FY25-26 Q2\n(02/28/26)',
            'FY25-26 Q4\n(08/31/26)', 'FY26-27 Q2\n(02/28/27)',
            'FY26-27 Q4\n(08/31/27)'
        ],
        'quarterly_keys': [
            'FY24-25_Q4', 'FY25-26_Q2', 'FY25-26_Q4',
            'FY26-27_Q2', 'FY26-27_Q4'
        ],
        'final_report_header': 'Final Report\n(Due 09/30/27)',
    },
    'R8': {
        'sheet_name': 'Round 8',
        'quarterly_labels': [
            'FY25-26 Q4\n(08/31/26)', 'FY26-27 Q2\n(02/28/27)',
            'FY26-27 Q4\n(08/31/27)', 'FY27-28 Q2\n(02/28/28)',
            'FY27-28 Q4\n(08/31/28)'
        ],
        'quarterly_keys': [
            'FY25-26_Q4', 'FY26-27_Q2', 'FY26-27_Q4',
            'FY27-28_Q2', 'FY27-28_Q4'
        ],
        'final_report_header': 'Final Report\n(Due 09/30/28)',
    },
}

# Fixed headers for columns A-D and K-Q (E-J are per-round)
BASE_HEADERS = [
    'Lead Institution', 'Proposal ID', '# Reporting Insitutions', 'Grant Name',
    None, None, None, None, None,  # E-I: quarterly (filled per round)
    None,  # J: Final Report (filled per round)
    'Report Waiting Approval', 'Grant Amount', 'Total Reported Expenditures',
    'Total Reported Expenditures Approved', '% Spent', 'Notes',
    'Unexpended according to last fiscal report'
]

BLUE_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HEADER_FONT = Font(name='Arial', bold=True, color='FFFFFF')
DATA_FONT = Font(name='Arial')
BOLD_FONT = Font(name='Arial', bold=True)
COL_WIDTHS = [28.38, 10.88, 10.88, 59.0, 7.75, 7.75, 7.75, 7.75, 7.75, 7.75,
              8.75, 11.75, 11.75, 11.75, 8.63, 32.38, 15.0]


def regenerate():
    if not os.path.exists(SCRAPED_DATA):
        print(f"ERROR: {SCRAPED_DATA} not found. Run the NOVA scrape first.")
        sys.exit(1)
    if not os.path.exists(SPREADSHEET):
        print(f"ERROR: {SPREADSHEET} not found. Ensure the spreadsheet is in this directory.")
        sys.exit(1)

    with open(SCRAPED_DATA, 'r') as f:
        data = json.load(f)

    wb = load_workbook(SPREADSHEET)

    for round_key in ['R5', 'R6', 'R7', 'R8']:
        if round_key not in data:
            print(f"WARNING: {round_key} not found in {SCRAPED_DATA}, skipping.")
            continue
        config = ROUND_CONFIGS[round_key]
        sheet_name = config['sheet_name']
        round_data = data[round_key]
        grants = round_data['grants']

        # Preserve Notes (P) and Unexpended (Q) from existing sheet
        preserved = {}
        if sheet_name in wb.sheetnames:
            old_ws = wb[sheet_name]
            for row in old_ws.iter_rows(min_row=2, max_col=17, values_only=False):
                pid = row[1].value
                if is_data_row(pid):
                    preserved[str(pid)] = {
                        'notes': row[15].value if len(row) > 15 else None,
                        'unexpended': row[16].value if len(row) > 16 else None,
                    }
            idx = wb.sheetnames.index(sheet_name)
            del wb[sheet_name]
            ws = wb.create_sheet(sheet_name, idx)
        else:
            ws = wb.create_sheet(sheet_name)

        # Sort grants alphabetically by lead institution
        grants.sort(key=lambda g: g['leadInstitution'])

        # Set column widths
        for i, w in enumerate(COL_WIDTHS):
            ws.column_dimensions[get_column_letter(i + 1)].width = w

        # Build headers
        headers = list(BASE_HEADERS)
        for i, label in enumerate(config['quarterly_labels']):
            headers[4 + i] = label
        headers[9] = config['final_report_header']

        # Write header row
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = HEADER_FONT
            cell.fill = BLUE_FILL
            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

        # Write data rows
        for row_idx, grant in enumerate(grants, 2):
            pid = grant['planId']

            ws.cell(row=row_idx, column=1, value=grant['leadInstitution']).font = DATA_FONT
            ws.cell(row=row_idx, column=2, value=pid).font = DATA_FONT
            ws.cell(row=row_idx, column=3, value=grant['dashboardInstitutions']).font = DATA_FONT
            ws.cell(row=row_idx, column=4, value=grant['grantName']).font = DATA_FONT

            # E-I: Quarterly status
            for qi, qkey in enumerate(config['quarterly_keys']):
                val = grant['quarterlyStatus'].get(qkey, False)
                ws.cell(row=row_idx, column=COL_QUARTERLY_START + qi, value=True if val else None).font = DATA_FONT

            # J: Final Report
            final = grant.get('finalReport', False)
            ws.cell(row=row_idx, column=COL_FINAL_REPORT, value=True if final else None).font = DATA_FONT

            # K: Report Waiting Approval (case-insensitive comparison)
            da = grant.get('dashboardApproval', '')
            waiting = da if da.lower() in ('pending approval', 'submitted') else None
            ws.cell(row=row_idx, column=COL_WAITING, value=waiting).font = DATA_FONT

            # L: Grant Amount
            cell = ws.cell(row=row_idx, column=COL_GRANT_AMOUNT, value=grant['projectBudget'])
            cell.font = DATA_FONT
            cell.number_format = '#,##0'

            # M: Total Reported Expenditures
            cell = ws.cell(row=row_idx, column=COL_TOTAL_EXP, value=grant['ptdExpenditure'])
            cell.font = DATA_FONT
            cell.number_format = '#,##0'

            # N: Total Reported Expenditures Approved (case-insensitive comparison)
            da_lower = da.lower()
            approved_exp = grant['ptdExpenditure'] if da_lower in ('approved', 'certified') else 0
            cell = ws.cell(row=row_idx, column=COL_APPROVED_EXP, value=approved_exp)
            cell.font = DATA_FONT
            cell.number_format = '#,##0'

            # O: % Spent (formula, guarded against division by zero)
            cell = ws.cell(row=row_idx, column=COL_PCT_SPENT, value=f'=IF(L{row_idx}=0,"",M{row_idx}/L{row_idx})')
            cell.font = DATA_FONT
            cell.number_format = '0.0%'

            # P: Notes (preserved)
            prev = preserved.get(pid, {})
            ws.cell(row=row_idx, column=COL_NOTES, value=prev.get('notes')).font = DATA_FONT

            # Q: Unexpended (preserved)
            ws.cell(row=row_idx, column=COL_UNEXPENDED, value=prev.get('unexpended')).font = DATA_FONT

        # Total row
        total_row = len(grants) + 2
        ws.cell(row=total_row, column=1, value='Central Mother Lode Region').font = BOLD_FONT

        # COUNTIF for quarterly columns E-I and Final Report J
        for col in range(COL_QUARTERLY_START, COL_FINAL_REPORT + 1):
            cl = get_column_letter(col)
            ws.cell(row=total_row, column=col,
                    value=f'=COUNTIF({cl}2:{cl}{total_row - 1},TRUE)').font = BOLD_FONT

        # SUM for L, M, N
        for col in [COL_GRANT_AMOUNT, COL_TOTAL_EXP, COL_APPROVED_EXP]:
            cl = get_column_letter(col)
            cell = ws.cell(row=total_row, column=col,
                           value=f'=SUM({cl}2:{cl}{total_row - 1})')
            cell.font = BOLD_FONT
            cell.number_format = '#,##0'

        # % Spent for total row (guarded against division by zero)
        cell = ws.cell(row=total_row, column=COL_PCT_SPENT, value=f'=IF(L{total_row}=0,"",M{total_row}/L{total_row})')
        cell.font = BOLD_FONT
        cell.number_format = '0.0%'

        # Report
        restored = sum(1 for g in grants if g['planId'] in preserved)
        print(f"{sheet_name}: {len(grants)} grants, total row at {total_row}, "
              f"preserved {restored}/{len(preserved)} Notes/Unexpended")

    wb.save(SPREADSHEET)
    print(f"\nSaved {SPREADSHEET}")


if __name__ == '__main__':
    regenerate()
