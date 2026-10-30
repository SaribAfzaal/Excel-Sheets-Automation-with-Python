# Excel & Sheets Automation with Python

Merge every Excel/CSV file in a folder into one clean, formatted workbook with a summary sheet.

## Features

- Merges `.xlsx`, `.xlsm`, `.xls` and `.csv` files (every sheet of multi-sheet workbooks)
- Handles mismatched columns: headers are normalized (`order id` / `Order ID` -> `Order Id`), and columns missing from a file are left blank
- Removes duplicate rows (ignoring which file they came from)
- Missing values: `drop`, `fill` with a default, or `flag` in a "Missing Values" column
  (columns a file never had are not counted as missing)
- Formatting: bold header, auto-fit widths, alternating row colors, frozen header, filters
- Conditional formatting on a chosen column: green above a threshold, red below
- `Summary` sheet: count/total/average/min/max for numeric columns (ID columns skipped) plus cleaning stats

## Setup

```bash
pip install -r requirements.txt
```

## Try it with the sample data

```bash
python generate_sample_data.py          # (re)creates ./sample_input
python excel_automation.py --config config.ini
```

Open `merged_output.xlsx`. The sample files have different column order/case, an extra column,
a duplicate order (1002), missing values, a multi-sheet workbook and text-formatted numbers (`$1,100`).

## Usage

Command line:

```bash
python excel_automation.py --input-folder sample_input --output result.xlsx \
    --cf-column Sales --threshold 1000 --missing fill
```

Config file (`config.ini`):

```ini
[settings]
input_folder = sample_input
output_file = merged_output.xlsx
missing_strategy = flag      ; drop | fill | flag
fill_value = N/A             ; used by 'fill' for text columns
numeric_fill_value = 0       ; used by 'fill' for numeric columns
cf_column = Sales
threshold = 1000
add_source_column = true
```

Command-line arguments override the config file:

```bash
python excel_automation.py --config config.ini --missing drop --threshold 1500
```

| Argument | Meaning |
|---|---|
| `--config` | path to the `.ini` file |
| `--input-folder` | folder containing the files to merge |
| `--output` | output `.xlsx` path |
| `--missing {drop,fill,flag}` | missing-value strategy (default `flag`) |
| `--fill-value`, `--numeric-fill-value` | defaults used by `fill` |
| `--cf-column`, `--threshold` | column and threshold for conditional formatting (both required) |
| `--no-source-column` | omit the "Source File" column |

## Notes

- Close the output file in Excel before re-running, or the write will fail.
- Conditional formatting uses native Excel rules, so it updates live if you edit values.
- Files that can't be read are skipped with a warning; the rest are still merged.
