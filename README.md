# KBFC DB Data Formatting

This repository contains a small workflow for transforming raw MKWC detection export data into an Excel workbook that is formatted for KBFC database upload.

## Purpose

The script in [mkwc/scripts/mkwc_detection_formatting.py](mkwc/scripts/mkwc_detection_formatting.py) reads a source Excel file exported from the remote detection system, reformats the date and time fields, maps reader and site identifiers, adds the required detection fields, and writes a new Excel file ready for import.

## Project Structure

- [mkwc/scripts/mkwc_detection_formatting.py](mkwc/scripts/mkwc_detection_formatting.py) - Main formatting script
- [mkwc/templates/KBFC_Detection.xlsm](mkwc/templates/KBFC_Detection.xlsm) - Excel template used as the base workbook
- [README.md](README.md) - Project overview and usage notes

## Requirements

- Python 3.x
- pandas
- openpyxl

## Setup

1. Update the file paths in the script for:
   - the detection template workbook
   - the source data workbook
   - the output folder
2. Make sure the source workbook contains a sheet named "Downloaded Tag IDs".
3. Run the script from the repository root or from the scripts folder:

```bash
py mkwc/scripts/mkwc_detection_formatting.py
```

## Output

The script writes an Excel file named in the format:

```text
mkwc_detection_formatted_YYYYMMDD.xlsx
```

The output is saved to the directory specified in the script.

## Notes

The current workflow includes hard-coded mappings for reader deployment IDs, site IDs, and antenna ID values. These may need to be adjusted for different deployments or source files.

Use the KBFC database to retrieve the appropriate ImportIDs and SiteIDs. 

This script does not populate the FileDate or ReaderID columns in the output file, as those are not required to upload the detection data to the KBFC database. 