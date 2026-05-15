# Supplier Matching Tool

## Overview

This Python script automates supplier matching between:

* **System 1 supplier exports**
* **System 2 contact exports**

The goal is to help prepare supplier mappings for a purchase order approval workflow.

The script performs:

1. **Exact matching** using:

   * System 1 supplier `Code`
   * System 2 `AccountNumber`

2. **Fuzzy matching** using supplier names when exact code matches are not found.

The output is an Excel workbook containing:

* Exact matches
* Unmatched suppliers
* Suggested fuzzy matches

---

# Requirements

## Python Version

Python 3.10+ recommended.

---

# Required Packages

Install dependencies:

```bash id="xxx"
pip install pandas openpyxl rapidfuzz
```

---

# Input Files

Place these files in the same folder as the script:

```text id="xxx"
system1_suppliers.xlsx
system2_contacts.xlsx
supplier_matcher.py
```

---

# Expected Columns

## System 1

Required columns:

* `Code`
* `Name`

## System 2

Required columns:

* `AccountNumber`
* `ContactName`

---

# How It Works

## 1. Load Excel Files

The script loads:

* System 1 supplier export
* System 2 contact export

using pandas.

---

## 2. Clean Column Headers

The script automatically removes:

* hidden spaces
* BOM characters
* formatting inconsistencies

from Excel column names.

---

## 3. Clean Supplier Codes

Supplier codes are normalized by:

* trimming spaces
* converting to uppercase

Example:

```text id="xxx"
abc123 → ABC123
```

This improves exact matching reliability.

---

## 4. Normalize Supplier Names

Supplier names are cleaned for fuzzy matching by:

* converting to uppercase
* removing punctuation
* removing company suffixes such as:

  * LTD
  * LIMITED
  * PTY
  * INC
  * LLC
  * COMPANY
  * CO

Example:

```text id="xxx"
ABC Paper Pty Ltd → ABC PAPER
```

This improves fuzzy name comparison accuracy.

---

## 5. Exact Matching

The script first attempts to match suppliers using:

```text id="xxx"
System 1 Code == System 2 AccountNumber
```

Exact matches are stored separately.

---

## 6. Fuzzy Matching

For suppliers that do not match by code:

* supplier names are compared using RapidFuzz
* fuzzy similarity scoring is applied
* only matches with a score >= 80 are included

Example:

```text id="xxx"
ABC PAPER
ABC PAPER SUPPLIES
```

may still match successfully.

---

# Output File

The script generates:

```text id="xxx"
supplier_matching_output.xlsx
```

Containing 3 worksheets:

---

## Exact Matches

Suppliers successfully matched by code.

---

## Unmatched

Suppliers with no exact match found.

---

## Fuzzy Suggestions

Possible supplier matches based on name similarity.

Includes:

* System 1 supplier code
* System 1 supplier name
* Suggested System 2 account number
* Suggested System 2 contact name
* Match confidence score

---

# Running the Script

Run from terminal:

```bash id="xxx"
python supplier_matcher.py
```

---

# Example Workflow

1. Export suppliers from System 1
2. Export contacts from System 2
3. Save both Excel files into the project folder
4. Run the script
5. Review the generated matching workbook
6. Use results for supplier mapping

---

# Notes

* Exact code matches are considered highest confidence.
* Fuzzy matches should always be reviewed manually.
* The fuzzy threshold can be adjusted in the script:

```python id="xxx"
if score >= 80:
```

Higher values = stricter matching.

---

# Future Improvements

Potential enhancements:

* confidence scoring (Green / Amber / Red)
* duplicate supplier detection
* automatic import file generation
* drag-and-drop interface
* GUI version for non-technical users

---

## Author

[Tyler Mc Luckie](https://tylermcluckie.github.io) · [ntmdigital.co](https://ntmdigital.co)
