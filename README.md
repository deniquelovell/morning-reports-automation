# Morning Reports Automation
Automated ingestion and processing of daily prime broker and treasury reports using Excel VBA.

## Overview
This project automates the daily import of operational reports from multiple brokers and internal systems into a standardized Excel reporting workbook.

## Business Problem
Daily treasury and middle-office reporting often requires manually locating files, opening inconsistent broker exports, clearing stale worksheet data, and pasting results into the correct tabs. This process is repetitive, time-sensitive, and prone to error.

## Solution
This VBA workflow:

- locates the correct daily report folder using `yymmdd` naming
- falls back to the most recent valid folder if today's folder is unavailable
- imports multiple broker and treasury files into standardized destination sheets
- handles malformed broker files misnamed as Excel files
- appends all tabs into one normalized output table
- preserves formatting for selected reporting sheets
- supports missing optional margin reports and switches treasury templates accordingly
- clears workbook filters before processing
- restores Excel state safely after execution

## Key Features
- Dynamic folder discovery
- Robust file matching by partial name and date token
- Defensive workbook opening for HTML/CSV-disguised Excel files
- Fast array-based paste operations
- Optional-file handling for margin workflows
- Automatic date standardization to mmddyy
- Safe cleanup and Excel state restoration

## Tech Stack
- Excel VBA
- Workbook automation
- File system handling
- Operational reporting workflows

## Repository Structure
- `src/modMain.bas` — orchestration entry point
- `src/modConfig.bas` — configuration constants and report definitions
- `src/modExcelHelpers.bas` — shared Excel utility functions
- `src/modPathHelpers.bas` — folder and path utilities
- `src/modFileDiscovery.bas` — file search and matching helpers
- `src/modWorkbookOpen.bas` — robust workbook-open logic
- `src/modClearHelpers.bas` — clearing and cleanup helpers
- `src/modImportEngine.bas` — general import logic
- `src/modNormalization.bas` — normalization logic
- `src/modTreasurySheets.bas` — treasury tab naming and template toggling
