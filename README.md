# bank-parse

Parse Handelsbanken (Swedish bank) xlsx export files into a unified
transactions CSV.

```
python3 bank-parse.py finance/data/*.xlsx --out transactions.csv
```

## Why

Bank exports come in bank-specific formats that are painful to work with.
This converts Handelsbanken xlsx exports into one clean, standard CSV
(columns: date, description, amount, balance, category-ready) that you can
load into any spreadsheet or finance tool.

## Requirements

- Python 3.8+
- `openpyxl` (`pip install openpyxl`)

## Output

CSV with unified columns. Multiple input files are merged and sorted
by date.
