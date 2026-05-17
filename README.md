# Python PDF Contract Generator

A Python automation script that generates bilingual (English/Thai) real estate contract documents from a structured Excel data file. It populates two separate `.docx` Word templates — one owner management contract and one tenant lease agreement — then converts both to PDF, saving hours of manual copy-paste work per contract batch.

---

## Table of Contents

- [Overview](#overview)
- [How It Works](#how-it-works)
- [Project Structure](#project-structure)
- [Input: Excel Data File](#input-excel-data-file)
  - [Column Reference](#column-reference)
- [Output Files](#output-files)
- [Automatic Calculations](#automatic-calculations)
- [Template Placeholder System](#template-placeholder-system)
  - [Lease Agreement Placeholders](#lease-agreement-placeholders)
  - [Owner Contract Placeholders](#owner-contract-placeholders)
  - [Superscript Ordinal Suffixes](#superscript-ordinal-suffixes)
  - [Conditional Blocks (Second Party)](#conditional-blocks-second-party)
- [Getting Started](#getting-started)
  - [For End Users (No Python Required)](#for-end-users-no-python-required)
  - [For Developers](#for-developers)
- [Dependencies](#dependencies)
- [Building the Executable (.exe)](#building-the-executable-exe)
- [Error Handling](#error-handling)
- [Limitations & Known Issues](#limitations--known-issues)

---

## Overview

Manually filling in contracts from a spreadsheet is slow and error-prone. This tool eliminates that work entirely. You fill in `data_input.xlsx`, run the script (or the compiled `.exe`), and receive polished, ready-to-print `.docx` and `.pdf` files in the `output/` folder — one set per row of contract data.

**Key capabilities:**

- Bilingual output: every field is written in both English and Thai.
- Dual document type: generates an owner management contract and a tenant lease agreement per row.
- Auto-conversion to PDF directly from the generated Word files.
- Smart calculations: deposit, late fees, lease duration, payment deadlines, and Buddhist Era years are all derived automatically.
- Handles optional second owners and second tenants — their clauses are injected only when present.
- Ordinal suffixes (1st, 2nd, 3rd…) are rendered as proper superscripts in the Word document.
- Nationality-aware ID detection: Thai nationals get "ID card / บัตรประชาชน"; all others get "passport / หนังสือเดินทาง".
- Duplicate-safe output: if a file already exists, the script appends `_1`, `_2`, etc. rather than overwriting.

---

## How It Works

```
data_input.xlsx
      │
      ▼
pandas (read & transpose)
      │
      ▼
Data processing pipeline
  ├─ Date parsing & formatting (EN + TH, Buddhist Era)
  ├─ Lease period calculation (years / months / days)
  ├─ Rent → deposit (×2), late fee, number-to-words (EN + TH baht text)
  ├─ Payment deadline (start day + 4, edge-case handling for month-end)
  ├─ Nationality → ID type determination
  └─ Name fallback (use English name if Thai name is blank)
      │
      ▼ (per row)
┌─────────────────────────────────────────┐
│  owner-contract-public.docx             │
│  placeholder → value substitution       │
│  → 代租管合約-{project} {room}.docx     │
│  → 代租管合約-{project} {room}.pdf      │
└─────────────────────────────────────────┘
┌─────────────────────────────────────────┐
│  lease-contract-public.docx             │
│  placeholder → value substitution       │
│  superscript suffix rendering           │
│  footer signature block population      │
│  → Lease Agreement-{project} {room}.docx│
│  → Lease Agreement-{project} {room}.pdf │
└─────────────────────────────────────────┘
```

The script reads the Excel file once, runs the full processing pipeline, then loops over every row to produce the document pair.

---

## Project Structure

```
python-pdf-contract-generator/
├── contract_maker.py           # Main script — all logic lives here
├── data_input.xlsx             # Input data (you fill this in)
├── lease-contract-public.docx  # Tenant lease agreement template
├── owner-contract-public.docx  # Owner management contract template
├── requirements.txt            # Runtime dependencies
├── requirements-dev.txt        # Type-stub dependencies (mypy)
├── output/                     # Generated files appear here (auto-created)
└── README.md
```

---

## Input: Excel Data File

`data_input.xlsx` uses a **transposed** layout: the first column is named `Attributes` and holds the field names (one per row); each additional column holds the values for one contract. This means you add contracts by adding columns, not rows.

Example layout:

| Attributes              | Contract A        | Contract B        |
|-------------------------|-------------------|-------------------|
| project_name            | The Grand Condo   | Sky Tower         |
| room_number             | 12/34             | 505               |
| rent_price              | 25000             | 45000             |
| rent_start_date         | 2026-06-01        | 2026-07-15        |
| …                       | …                 | …                 |

### Column Reference

Fields marked **Required** will cause an error if left blank. Optional fields produce empty strings or are omitted from the document if not filled.

#### Property

| Field | Type | Required | Description |
|---|---|---|---|
| `project_name` | Text | Yes | Condominium / project name |
| `room_number` | Text | Yes | Room number (e.g. `12/34` or `505`) |
| `building_no` | Text | No | Building number within the project |
| `room_floor_number` | Integer | No | Floor number (rendered with ordinal suffix) |
| `room_area` | Number | No | Unit area in square meters |
| `project_address` | Text | No | Full address in English |
| `project_address_th` | Text | No | Full address in Thai |

#### Lease Dates

| Field | Type | Required | Description |
|---|---|---|---|
| `rent_start_date` | Date (YYYY-MM-DD) | Yes | First day of the lease |
| `rent_end_date` | Date (YYYY-MM-DD) | Yes | Last day of the lease |

#### Financials

| Field | Type | Required | Description |
|---|---|---|---|
| `rent_price` | Number | Yes | Monthly rent in THB |
| `water_meter_no` | Text | No | Water meter number |
| `electric_meter_no` | Text | No | Electricity meter number |

#### Owner (Party 1 — Landlord)

| Field | Type | Required | Description |
|---|---|---|---|
| `owner_name` | Text | Yes | Full name in English |
| `owner_name_th` | Text | No | Full name in Thai (falls back to English if blank) |
| `owner_passport` | Text | Yes | Passport or ID card number |
| `owner_nationality` | Text | Yes | Nationality (use `Thailand` or `Thai` for Thai nationals) |
| `owner_nationality_th` | Text | No | Nationality in Thai |
| `owner_passport_expire_date` | Date | Yes | Document expiry date |
| `owner_bank` | Text | No | Bank name in English |
| `owner_bank_th` | Text | No | Bank name in Thai |
| `owner_bank_branch` | Text | No | Bank branch in English |
| `owner_bank_branch_th` | Text | No | Bank branch in Thai |
| `owner_bank_account_no` | Text | No | Bank account number |
| `owner_bank_account_name` | Text | No | Account holder name |

#### Second Owner (Optional)

Leave all `*_2` owner fields blank if there is only one owner. Their entire clause block is suppressed automatically.

| Field | Type | Description |
|---|---|---|
| `owner_name_2` | Text | Second owner full name in English |
| `owner_name_2_th` | Text | Second owner name in Thai |
| `owner_passport_2` | Text | Second owner passport / ID number |
| `owner_nationality_2` | Text | Second owner nationality |
| `owner_nationality_2_th` | Text | Second owner nationality in Thai |
| `owner_passport_expire_date_2` | Date | Second owner document expiry |

#### Tenant (Party 2 — Lessee)

| Field | Type | Required | Description |
|---|---|---|---|
| `tenant_name` | Text | Yes | Full name in English |
| `tenant_name_th` | Text | No | Full name in Thai (falls back to English if blank) |
| `tenant_passport` | Text | Yes | Passport or ID card number |
| `tenant_nationality` | Text | Yes | Nationality (use `Thailand` or `Thai` for Thai nationals) |
| `tenant_nationality_th` | Text | No | Nationality in Thai |
| `tenant_passport_expire_date` | Date | Yes | Document expiry date |

#### Second Tenant (Optional)

| Field | Type | Description |
|---|---|---|
| `tenant_name_2` | Text | Second tenant full name in English |
| `tenant_name_2_th` | Text | Second tenant name in Thai |
| `tenant_passport_2` | Text | Second tenant passport / ID number |
| `tenant_nationality_2` | Text | Second tenant nationality |
| `tenant_nationality_2_th` | Text | Second tenant nationality in Thai |
| `tenant_passport_expire_date_2` | Date | Second tenant document expiry |

#### Witnesses

| Field | Type | Description |
|---|---|---|
| `witness_name` | Text | First witness name |
| `witness_name_2` | Text | Second witness name (optional; signature line suppressed if blank) |

---

## Output Files

For each data column in the Excel file the script produces four files in the `output/` folder:

| File | Description |
|---|---|
| `代租管合約-{project_name} {room}.docx` | Owner management contract (Word) |
| `代租管合約-{project_name} {room}.pdf` | Owner management contract (PDF) |
| `Lease Agreement-{project_name} {room}.docx` | Tenant lease agreement (Word) |
| `Lease Agreement-{project_name} {room}.pdf` | Tenant lease agreement (PDF) |

Room numbers containing `/` are reformatted: `12/34` becomes `(12-34)` in the filename. If a file already exists the script adds a numeric suffix (`_1`, `_2`, …) to avoid overwriting.

---

## Automatic Calculations

The following values are **derived from your input** and do not need to be entered manually:

| Calculated Value | Logic |
|---|---|
| **Security deposit** | `rent_price × 2` |
| **Late payment fee** | 500 THB if rent < 30,000 THB; 1,000 THB if rent ≥ 30,000 THB |
| **Payment deadline** | Start day + 4 (e.g. start on the 1st → must pay by 5th). Special strings are used for days 27–31 to handle month-end edge cases (e.g. "31st or 1st"). |
| **Lease duration** | Calculated precisely in years, months, and days using `dateutil.relativedelta`. |
| **Number-to-words (English)** | Rent, deposit, and late fee converted to English words via `num2words`. |
| **Number-to-words (Thai)** | Rent, deposit, and late fee converted to Thai baht text via `pythainlp.bahttext`. |
| **Buddhist Era year** | CE year + 543 (e.g. 2026 → พ.ศ. 2569). |
| **Thai month names** | Month number mapped to Thai (e.g. 6 → มิถุนายน). |
| **ID type label** | `Thailand` / `Thai` nationality → "ID card / บัตรประชาชน"; all others → "passport / หนังสือเดินทาง". |
| **Comma-formatted numbers** | Rent, deposit, late fee formatted with thousands separator (e.g. 25,000). |
| **Floor ordinal suffix** | Floor number rendered with correct English ordinal (1st, 2nd, 3rd, 4th…). |

---

## Template Placeholder System

The Word templates contain marker strings (placeholders) that the script searches for and replaces. Placeholders respect the case of the original: an ALL-CAPS placeholder produces an ALL-CAPS replacement; a Title-case placeholder produces a Title-case replacement.

### Lease Agreement Placeholders

Place any of these strings inside your `lease-contract-public.docx` template:

| Placeholder | Replaced With |
|---|---|
| `startdayplahor` | Lease start day (integer) with ordinal superscript |
| `Startmmyyplahor` | Start month and year in English (e.g. `June 2026`) |
| `Startdatethplahor` | Full start date in Thai |
| `enddayplahor` | Lease end day with ordinal superscript |
| `Endmmyyplahor` | End month and year in English |
| `Enddatethplahor` | Full end date in Thai |
| `OWNERNAMEPLAHOR` | Owner full name (English, UPPERCASE) |
| `OWNERNAMETHPLAHOR` | Owner full name (Thai, UPPERCASE) |
| `OWNERPASSPLAHOR` | Owner passport / ID number |
| `Ownernatplahor` | Owner nationality (English) |
| `Ownernatthplahor` | Owner nationality (Thai) |
| `Ownerpassexpplahor` | Owner document expiry date (English) |
| `Ownerpassexpthplahor` | Owner document expiry date (Thai) |
| `ow1idpenplahor` | "ID card" or "passport" for owner 1 (English, no case change) |
| `ow1idpthplahor` | "บัตรประชาชน" or "หนังสือเดินทาง" for owner 1 (Thai, no case change) |
| `Projectnameplahor` | Project name |
| `Roomnumberplahor` | Room number |
| `Buildingnumberplahor` | Building number |
| `Projectaddressplahor` | Project address (English) |
| `Projectaddressthplahor` | Project address (Thai) |
| `TENANTNAMEPLAHOR` | Tenant full name (English, UPPERCASE) |
| `TENANTNAMETHPLAHOR` | Tenant full name (Thai, UPPERCASE) |
| `TENANTPASSPLAHOR` | Tenant passport / ID number |
| `Tenantnatplahor` | Tenant nationality (English) |
| `Tenantnatthplahor` | Tenant nationality (Thai) |
| `Tenantpassexpplahor` | Tenant document expiry date (English) |
| `Tenantpassexpthplahor` | Tenant document expiry date (Thai) |
| `te1idpenplahor` | "ID card" or "passport" for tenant 1 (English, no case change) |
| `te1idpthplahor` | "บัตรประชาชน" or "หนังสือเดินทาง" for tenant 1 (Thai, no case change) |
| `floorplahor` | Floor number (integer) |
| `floorsufplahor` | Floor number with ordinal suffix (e.g. `3rd`) |
| `Areaplahor` | Room area |
| `rentstartdatedayplahor` | Numeric start day (for Thai payment clause) |
| `lastdaybeforefeeplahor` | Payment deadline in English (e.g. `5thsuffixplahor`) |
| `lastdaybeforefeethplahor` | Payment deadline in Thai (e.g. `5`) |
| `Rentnoplahor` | Rent amount with comma (e.g. `25,000`) |
| `Rentenplahor` | Rent amount in English words |
| `Rentthplahor` | Rent amount in Thai baht text |
| `Latefeenumplahor` | Late fee amount with comma |
| `Latefeetextenplahor` | Late fee in English words |
| `Latefeetextthplahor` | Late fee in Thai baht text |
| `Leaseperiodenplahor` | Lease duration in English (e.g. `1 Year, 6 Months`) |
| `Leaseperiodthplahor` | Lease duration in Thai (e.g. `1 ปี 6 เดือน`) |
| `Depositplahor` | Security deposit amount with comma |
| `Deposittextplahor` | Deposit in English words |
| `Deposittextthplahor` | Deposit in Thai baht text |
| `OWNBANAPLAHOR` | Owner bank name (English) |
| `OWNBANATHPLAHOR` | Owner bank name (Thai) |
| `Ownbabrplahor` | Bank branch (English) |
| `Ownbabrthplahor` | Bank branch (Thai) |
| `Ownbaaccnoplahor` | Bank account number |
| `OWNERBANKACCOUNTNAMEPLAHOR` | Bank account holder name |
| `Wameplahor` | Water meter number |
| `Elmeplahor` | Electricity meter number |

**Footer signature block** (inside footer tables):

| Placeholder | Replaced With |
|---|---|
| `OWNERNAMEPLAHOR` | Owner name |
| `TENANTNAMEPLAHOR` | Tenant name |
| `WITNESSNAMEPLAHOR` | Witness name |
| `OWNERNAME2FTPLAHOR` | Second owner name wrapped in parentheses, e.g. `(John Doe)` |
| `TENANTNAME2FTPLAHOR` | Second tenant name in parentheses |
| `WITNESSNAME2FTPLAHOR` | Second witness name in parentheses |
| `Iflandlord2plahor` | Signature line label for second landlord (blank if no second owner) |
| `Iftenant2plahor` | Signature line label for second tenant (blank if no second tenant) |
| `Ifwitness2plahor` | Signature line label for second witness (blank if no second witness) |

### Owner Contract Placeholders

Place these inside `owner-contract-public.docx`. These are paragraph-level replacements:

| Placeholder | Replaced With |
|---|---|
| `PROJECTNAMEHOLDER` | Project name |
| `UNITNUMBERHOLDER` | Room number |
| `STYAHD` | Lease start year |
| `STMDHD` | Lease start month |
| `STDYHD` | Lease start day |
| `ENYAHD` | Lease end year |
| `ENMDHD` | Lease end month |
| `ENDYHD` | Lease end day |

Table-level replacements:

| Placeholder | Replaced With |
|---|---|
| `NAME1HOLDER` | Owner name |
| `NAME2HOLDER` | `/ {owner_name_2}` (with slash prefix, blank if no second owner) |
| `NAME2NOSLASHHOLDER` | Second owner name without slash prefix |
| `PSPT1HO` | Owner passport / ID number |
| `PSPT2HO` | `/ {owner_passport_2}` (with slash prefix, blank if no second owner) |

### Superscript Ordinal Suffixes

Within any placeholder value that contains an ordinal number, the script looks for these special marker strings and converts them to properly formatted superscript text in the Word document:

| Marker | Superscript |
|---|---|
| `stsuffixplahor` | st |
| `ndsuffixplahor` | nd |
| `rdsuffixplahor` | rd |
| `thsuffixplahor` | th |

Example: the value `5thsuffixplahor` becomes "5" followed by a superscript "th" in the final document.

The superscript runs use the `Cordia New (Body CS)` font at 15pt to match the Thai-compatible body font.

### Conditional Blocks (Second Party)

For second owners and second tenants, the template contains special conditional placeholders. When the relevant name field is blank in the Excel file, the entire clause is replaced with an empty string:

| Placeholder in Template | Injected Text (when present) |
|---|---|
| `Ifowner2dicenplahor` | `, {name}, holding {id type} no. {passport} of {nationality} Nationality Expiry date {expiry date}` |
| `Ifowner2dicthplahor` | `, {name_th} ถือ{id type_th}เลขที่ {passport} สัญชาติ {nationality_th} หมดอายุวันที่ {expiry_th}` |
| `Iftenant2dicenplahor` | Same structure for second tenant (English) |
| `Iftenant2dicthplahor` | Same structure for second tenant (Thai) |

---

## Getting Started

### For End Users (No Python Required)

**Step 1 — Download the project files**

1. Go to the repository page on GitHub.
2. Click the green **Code** button → **Download ZIP**.
3. Unzip the archive. The folder contains `data_input.xlsx` and the two `.docx` templates.

**Step 2 — Download the executable**

1. Download `contract_maker.exe` from the [Google Drive link](https://drive.google.com/file/d/1GVAsSgw9n7JNU0VFs6RpfVaz7MFmlksw/view?usp=sharing).
2. Windows SmartScreen or your browser may show a warning for unsigned `.exe` files — this is expected.

**Step 3 — Prepare your data**

1. Open `data_input.xlsx`.
2. Fill in all required fields (see [Column Reference](#column-reference) above). Each contract is a separate column.
3. Save and close the file.

**Step 4 — Run**

1. Place `contract_maker.exe` in the same folder as `data_input.xlsx` and the two `.docx` templates.
2. Double-click `contract_maker.exe`.
3. A console window will appear. Wait for it to print the output file paths and close.
4. Open the `output/` folder to find your generated `.docx` and `.pdf` files.

> **Note:** Microsoft Word must be installed for the PDF conversion to work. `docx2pdf` uses Word's built-in PDF export on Windows.

---

### For Developers

**Prerequisites**

- Python 3.8 or newer
- pip
- Microsoft Word (required at runtime by `docx2pdf` for PDF conversion)

**Setup**

```bash
# Clone the repository
git clone https://github.com/your-username/python-pdf-contract-generator.git
cd python-pdf-contract-generator

# Create and activate a virtual environment (recommended)
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS / Linux

# Install runtime dependencies
pip install -r requirements.txt

# Optional: install type stubs for static analysis
pip install -r requirements-dev.txt
```

**Run**

```bash
python contract_maker.py
```

Generated files will appear in the `output/` folder.

**Type checking (optional)**

```bash
mypy contract_maker.py
```

---

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| `python-docx` | 1.2.0 | Read and write `.docx` Word files; paragraph, run, and table manipulation |
| `pandas` | 2.2.3 | Read Excel data; transposed DataFrame as the data model |
| `numpy` | 2.1.3 | Conditional column construction (`np.select`) |
| `python-dateutil` | 2.9.0 | Precise lease duration via `relativedelta` |
| `pythainlp` | 5.1.2 | Thai baht text conversion (`bahttext`) |
| `num2words` | 0.5.14 | English number-to-words conversion |
| `docx2pdf` | 0.1.8 | Convert generated `.docx` files to PDF via Microsoft Word |

Development-only (type stubs for mypy):

| Package | Version |
|---|---|
| `pandas-stubs` | 3.0.0.260204 |
| `types-python-dateutil` | 2.9.0.20260508 |

Install all runtime dependencies:

```bash
pip install -r requirements.txt
```

---

## Building the Executable (.exe)

To compile the script into a standalone `.exe` so end users do not need Python installed:

```bash
pip install pyinstaller
pyinstaller --onefile contract_maker.py
```

The compiled file appears at `dist/contract_maker.exe`. Copy it into the project folder alongside `data_input.xlsx` and the two `.docx` templates before distributing.

> **Important:** The `.exe` does not bundle the template files or `data_input.xlsx`. Users still need those files in the same folder.

---

## Error Handling

The script raises descriptive errors and stops immediately in these cases:

**Missing or invalid dates** — any date field that is blank or cannot be parsed raises:
```
ERROR: Column 'rent_start_date' has 1 missing or invalid date(s).
Affected row(s): [0]
ALL date fields must be filled in correctly before running.
Partial data is NOT acceptable — fix the Excel file and try again.
```

**Missing rent price** — a blank `rent_price` cell raises:
```
ERROR: 'rent_price' contains a missing value.
ALL rent price fields must be filled in correctly before running.
Partial data is NOT acceptable — fix the Excel file and try again.
```

Fix the Excel file and re-run. No partial output is generated for the failed row.

---

## Limitations & Known Issues

- **Hard-coded templates:** The placeholder names are specific to the included `.docx` templates. Adapting the script to a completely different contract format requires editing the placeholder dictionaries in `contract_maker.py`.
- **Microsoft Word required:** PDF conversion via `docx2pdf` on Windows invokes the locally installed Microsoft Word. If Word is not installed, `.docx` files are still generated but PDF conversion will fail.
- **Procedural architecture:** The script was written in a linear, procedural style — processing and document generation are not encapsulated in classes. This makes it straightforward to read top-to-bottom but harder to unit test or extend.
- **Single-file processing:** All rows in the Excel file are processed in one run with no progress bar or recovery from mid-batch failure. If the script crashes on row 5 of 10, rows 1–4 may or may not be written to disk depending on where the failure occurred.
- **Thai font dependency:** Superscript ordinal runs are hardcoded to `Cordia New (Body CS)` at 15pt. If this font is not available on the target machine, Word will substitute the nearest available font.
