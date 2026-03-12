# K12 SWP Grant Tracking — Central Mother Lode

Tracks K-12 Strong Workforce Program (SWP) grant fiscal reporting for the Central Mother Lode region. Data is scraped from the NOVA system (nova.cccco.edu) and maintained in an Excel spreadsheet.

## Prerequisites

- **Python 3.8+**
- **NOVA account** with SSO access to nova.cccco.edu
- **Claude Code** with **one** of these browser automation MCP servers configured:
  - **Option A (recommended):** `claude-in-chrome` — uses your existing Chrome session. Log into NOVA in Chrome before starting.
  - **Option B (fallback):** `playwright` MCP server — launches a separate browser. You'll be prompted to log into NOVA in the Playwright browser window on first run.

Install Python dependencies:

```bash
pip install -r requirements.txt
```

## How to Run an Update

1. Open Chrome and navigate to [nova.cccco.edu](https://nova.cccco.edu). Verify you are logged in (SSO should auto-authenticate).
2. Start Claude Code in this project directory.
3. Tell Claude which rounds to update:
   ```
   Update grant data from NOVA for rounds R6 and R7
   ```
4. Claude follows the operational runbook in `CLAUDE.md` — scraping NOVA, updating `scraped_data.json`, and regenerating the spreadsheet.
5. After scraping completes, the spreadsheet is rebuilt automatically via:
   ```bash
   python3 regenerate_spreadsheet.py
   ```
6. Run validation to check data integrity:
   ```bash
   python3 validate.py
   ```

## Files

| File | Purpose |
|------|---------|
| `K12 Reporting Status - Central Mother Lode.xlsx` | The tracking spreadsheet (source of truth) |
| `scraped_data.json` | Raw data from the last NOVA scrape |
| `regenerate_spreadsheet.py` | Rebuilds Round 5-8 tabs from scraped data |
| `validate.py` | Automated validation checks on scraped data and spreadsheet |
| `scrape_round.py` | Hardcoded plan IDs per round (R5-R8) |
| `CLAUDE.md` | Operational runbook for the update process |
| `SPREADSHEET_CONTEXT.md` | Structural documentation of the spreadsheet |
| `backups/` | Timestamped spreadsheet backups (created before each update) |

## Troubleshooting

**`ModuleNotFoundError: No module named 'openpyxl'`**
Run `pip install -r requirements.txt`.

**Chrome extension not found / no NOVA tab detected**
Ensure the `claude-in-chrome` MCP server is configured in Claude Code settings and Chrome is running with the extension active. If Chrome isn't available, configure the Playwright MCP server instead — Claude will auto-detect which backend to use.

**SSO timeout / login redirect loop**
NOVA uses SSO that auto-authenticates. If pages keep redirecting to `/login`, your session may have expired. Log in manually at nova.cccco.edu (in Chrome or the Playwright browser window), then retry.

**Playwright browser not installed**
If you get an error about the browser not being installed, Claude will call `browser_install` automatically. This downloads Chromium (~150MB) on first use.

**Playwright SSO session lost mid-run**
If NOVA pages start showing login screens during a long run, the Playwright session may have timed out. Log in again in the Playwright browser window and tell Claude to continue.

**"scraped_data.json not found"**
The scraping phase must complete before running `regenerate_spreadsheet.py`. Start a NOVA update first, or ensure a previous `scraped_data.json` exists in this directory.

**Spreadsheet not found**
Ensure `K12 Reporting Status - Central Mother Lode.xlsx` is in the project root directory.

For detailed process documentation, see `CLAUDE.md`.
