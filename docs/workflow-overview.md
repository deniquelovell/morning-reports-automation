# Workflow Overview

## Purpose
This project automates the daily ingestion of broker and treasury reports into a standardized Excel reporting workbook.

## Process Flow
1. Load the base report directory from the workbook path or saved named range.
2. Validate that the selected directory contains daily folders in `yymmdd` format.
3. Locate today's report folder, or fall back to a valid recent folder.
4. Search for the newest matching report file based on:
   - report prefix
   - optional date token
   - supported file extensions
5. Open the source workbook using fallback handling for malformed or misnamed files.
6. Select the correct worksheet for import.
7. Clear stale content from the destination range while preserving formatting where needed.
8. Paste the new data into the target worksheet.
9. Restore Excel state and notify the user when processing is complete.

## Design Goals
- Reduce manual daily reporting steps
- Improve consistency across recurring imports
- Handle messy real-world file naming
- Support reusable workflows across similar report packages
