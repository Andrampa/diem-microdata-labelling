# diem-microdata-labelling

Automatically apply official DIEM codebook labels to household survey microdata.

## Overview

This repository provides Python and R scripts to automatically apply value labels
to DIEM household survey microdata using the official DIEM codebooks.

The tool is designed for analysts and data managers working with DIEM microdata
who need a clean, labelled version of survey datasets for analysis, validation,
or sharing.

## ⚠️ Important warning

This script is designed to work **only with DIEM microdata as downloaded directly
from the DIEM Hub**, without any structural modification.

In particular:

- The script assumes that field names, coding schemes, and variable structures
  exactly match the official DIEM codebooks
- If microdata files have been manually edited, renamed, restructured, or partially
  transformed, the labelling may fail or produce incorrect results
- If microdata from **different infrastructure versions (v1.0 and v2.0)** are
  merged into a single file, the script **may not work correctly**, as the automatic
  infrastructure detection relies on mutually exclusive field structures

Always ensure that:
- The input file comes from a **single DIEM infrastructure version**
- The file structure has not been altered after download

## What the script does

Given a DIEM microdata file (`.csv`, `.xls`, or `.xlsx`), the script will:

- Automatically detect the DIEM infrastructure version  
  - Version 1.0: pre–December 2022 questionnaires  
  - Version 2.0: post–December 2022 questionnaires (identified by the `qc_step0_date` field)
- Download the appropriate DIEM codebook from the DIEM Hub if not already present
- Apply coded-value labels using:
  - One dictionary sheet per variable
  - A dedicated list of derived yes/no fields
- Replace coded values with human-readable labels (or optionally add label columns)
- Preserve the original file format in the output
- Print a detailed summary to the console, including:
  - Labelled fields
  - Skipped fields
  - Fields with low match rates
  - Unmapped codes requiring review

## Input requirements

### Microdata file

- Supported formats: `.csv`, `.xls`, `.xlsx`
- Must be DIEM household survey microdata downloaded from the DIEM Hub
- Must come from a single infrastructure version
- The presence of the `qc_step0_date` field is used to detect infrastructure version 2.0


### Codebooks

The script automatically downloads the official DIEM codebooks from ArcGIS Online:

- DIEM codebook v1.0 (pre–2022)
- DIEM codebook v2.0 (post–2022)

Downloaded codebooks are saved in the same folder as the input microdata file and
reused in subsequent runs.

## How to use

1. Open the script and set the path to your microdata file:

```python
MICRODATA_FILE = r"path\to\your\DIEM_microdata_file.xlsx"

## Run the script
```bash
python diem_microdata_labelling.py
```

The labelled output file will be created in the same folder as the input file,
with `_LABELLED` appended to the filename.

## Output

- Output format always matches the input format (`.csv`, `.xls`, or `.xlsx`)

### Output filename examples

- `DIEM_household_surveys_microdata.xlsx` → `DIEM_household_surveys_microdata_LABELLED.xlsx`

## Configuration options

### Replace values or add label columns
```python
ADD_LABEL_COLUMNS = False
```

- `False`: coded values are replaced by labels
- `True`: labels are added as separate columns

### Minimum acceptable match rate
```python
MIN_OK_MATCH_RATE = 0.95
```

Fields with lower match rates are flagged in the summary output.

## Typical use cases

- Data quality checks
- Rapid validation of survey outputs
- Preparing labelled datasets for analysts or partners
- Reviewing questionnaire changes across infrastructure versions

## Dependencies

### Python version
- Python 3.8 or higher
- `pandas`
- `openpyxl` (for Excel input and output)

### R version
- R 4.0 or higher
- readr
- readxl
- writexl
- stringr


## Notes

- Fields with no values are skipped automatically
- Fields with unmapped codes are clearly reported
- The script does not modify fields unless a valid dictionary is available
- The approach is intentionally conservative to avoid unintended data changes

## Maintainer

**DIEM Data Team**  
FAO – Office of Emergencies and Resilience (OER)
