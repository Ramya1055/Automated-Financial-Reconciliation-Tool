# Automated-Financial-Reconciliation-Tool

# NEFT, RTGS, and NPCI Reconciliation 

This project automates the reconciliation of NEFT, RTGS, and NPCI financial transactions under the **Data Analytics and IT Department of APCOB**. It is a Python-based solution that replaces manual processes with a fast, accurate, and reliable system.

---

## Overview

- Automates reconciliation across:
  - Core Banking System (CBS) text reports
  - SFMS (Structured Financial Messaging System) Excel files
  - NPCI RTGS CSV files
- Extracts and matches transaction details like:
  - Trace Numbers
  - Amounts
  - Narratives
  - End-to-End IDs (NEFT)
  - Transaction End IDs (RTGS)
  - NPCI Reference Numbers (from CBS narratives)
- Validates transactions based on:
  - Date
  - Amount
  - Transaction IDs
- Tolerates inconsistencies in:
  - File formats
  - Date formats
  - Numeric values

---

## Features

- Regex-based parsing of unstructured and semi-structured files
- Fuzzy matching logic for inconsistent identifiers
- Automated segregation and handling of CAMT (ISO 20022) records
- Excel report generation with:
  - Matched records
  - Unmatched CBS/SFMS/NPCI entries
  - CAMT-specific sheets
  - Reconciliation Summary Dashboard
- Duplicate detection
- Error-tolerant processing of changing formats and layouts

---

## Tools & Libraries

- **Python 3.x**
- `pandas`
- `openpyxl`
- `re` (Regular Expressions)

---

## Business Impact

- Reduced reconciliation time from **hours to minutes**
- Improved accuracy and reduced manual errors
- Ensures auditability and compliance
- Adaptable to:
  - NEFT (Inward / Outward)
  - RTGS (Inward / Outward)
  - NPCI RTGS workflows

 ---

## Folder Structure 

├── sample input files/
│ ├── cbs_reports/
│ ├── sfms_files/
│ └── npci_csvs/
├── sample output/
│ └── Sample output/
├── main.py
└── README.md
