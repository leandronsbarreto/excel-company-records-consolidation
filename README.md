# Excel Company Records Consolidation

This repository contains a Python script for consolidating company records from multiple Excel files into one output file.

The script reads all `.xlsx` files in the current folder, checks whether they have the expected structure and combines the valid files into a single Excel workbook.

## Sample files

The repository includes two sample input files:

```text
File_1.xlsx
File_2.xlsx
```

Each file contains 70 company records, not counting the header row.

Both files use the same column structure:

```text
Company Name
ID Number
Commercial Register Number
Street and Number
Postal Code
City
Phone
Email Address
```

## What the script does

The script:

```text
reads all .xlsx files in the current folder
ignores the final output file if it already exists
uses the first worksheet of each Excel file
skips the first row as the header row
removes completely empty rows
checks whether each file has the expected number of columns
combines all valid files into one DataFrame
exports the result to a new Excel file
```

The output file is:

```text
consolidated.xlsx
```

## Input requirements

Each input file must contain exactly 8 columns.

Expected columns:

```text
Company Name
ID Number
Commercial Register Number
Street and Number
Postal Code
City
Phone
Email Address
```

Files with a different number of columns are skipped and shown in the terminal output.

## Requirements

The script uses:

```text
pandas
openpyxl
```

Install the required packages with:

```bash
pip install pandas openpyxl
```

## How to run

Place the script in the same folder as the Excel files.

Example folder structure:

```text
project-folder/
│
├── consolidate.py
├── File_1.xlsx
└── File_2.xlsx
```

Then run:

```bash
python consolidate.py
```

After the script finishes, it creates:

```text
consolidated.xlsx
```

## Terminal output

The script prints information about the process, including:

```text
current folder
number of Excel files found
file names
number of columns and rows per file
files skipped because of incorrect structure
number of valid DataFrames
total number of consolidated records
output file creation
```

With the provided sample files, the expected number of consolidated records is:

```text
140
```

## Notes

Keep only the input Excel files in the folder, or make sure that other `.xlsx` files do not have the same structure unless they should also be included.

The script excludes `consolidated.xlsx` automatically if it already exists, so running the script again does not duplicate the previous output file.

## Development note

I developed and adapted this script with support from ChatGPT. The logic, testing and final adjustments were reviewed by me and tailored to this specific Excel comparison use case. The script reflects my own workflow requirements.

## Data privacy

The sample files are intended for demonstration purposes. They do not contain confidential company data, personal data or sensitive information.
