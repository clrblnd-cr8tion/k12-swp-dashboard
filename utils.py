#!/usr/bin/env python3
"""Shared constants and helpers for K12 SWP grant tracking scripts."""

SPREADSHEET = 'K12 Reporting Status - Central Mother Lode.xlsx'
SCRAPED_DATA = 'scraped_data.json'


def is_data_row(pid):
    """True if pid is a grant data row, not the header or total row."""
    return pid and 'Central' not in str(pid) and 'Region' not in str(pid)
