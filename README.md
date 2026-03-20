# Morning Reports Automation

Automated ingestion and processing of daily broker and treasury reports using Excel VBA.

## Overview
This project automates the daily import of operational reports into a standardized Excel reporting workbook. It eliminates manual file handling, reduces errors, and ensures consistency across recurring reporting workflows.

## Problem
Daily reporting required:
- manually locating files across dated folders
- opening inconsistent report formats
- clearing and repopulating Excel sheets
- handling missing or misnamed files

This process was time-consuming, repetitive, and prone to error.

## Solution
This VBA-based automation:

- dynamically locates report folders using date-based logic
- identifies and selects the latest valid files using flexible name matching
- handles malformed or misnamed files (e.g., HTML/CSV disguised as Excel)
- standardizes data import into predefined reporting sheets
- preserves formatting while removing stale data
- restores Excel settings safely after execution

## Key Features
- Dynamic folder discovery (`yymmdd` structure)
- Robust file matching (prefix + date tokens)
- Fallback handling for missing daily folders
- Fast array-based data transfers
- Modular architecture (separate logic by function)
- Defensive workbook opening logic

## Tech Stack
- Excel VBA
- File system handling
- Workbook automation
- Operational reporting workflows

## Repository Structure
- `src/` — VBA modules
- `docs/` — workflow and structure documentation

## Impact

- Reduced manual report preparation time by ~60–80%
- Eliminated repetitive daily data-copy tasks across multiple reports
- Improved consistency and reduced risk of human error in reporting
- Standardized ingestion process across multiple report types
- Enabled scalable workflow that can be reused across similar reporting structures
- Designed with modular architecture to support future extensions and additional report types from any prime broker
