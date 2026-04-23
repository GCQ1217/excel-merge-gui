# Excel Merge Tool

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

[简体中文](README.zh-CN.md) | English

A lightweight desktop GUI tool to merge multiple Excel / CSV files into one by appending rows.

## Features

- Merge .xlsx, .xls, and .csv files by row
- Select specific sheet for each Excel file
- Set global start row (skip headers)
- Set per-file row range
- Auto-detect CSV encoding (UTF-8 / GBK / GB18030 / Latin1 ...)
- Auto-detect CSV delimiter
- Convert CSV to Excel

## Supported Formats

| Format | Extension | Read | Write |
|--------|-----------|------|-------|
| Excel (new) | .xlsx | Yes | Yes |
| Excel (legacy) | .xls | Yes | — |
| CSV | .csv | Yes | Yes |

## Quick Start

### Run from source

```bash
pip install -r requirements.txt
python excel_merge_gui.py
```

### Run as exe

Download the pre-built exe or build it yourself:

```bash
scripts\build_pyinstaller.bat
scripts\build_nuitka.bat
```

## Project Structure

```
├── README.md
├── README.zh-CN.md
├── requirements.txt
├── excel_merge_gui.py
├── testdata/
│   ├── employees_1.xlsx
│   ├── employees_2.xlsx
│   ├── multi_sheet.xlsx
│   ├── products.xls
│   ├── cities_utf8.csv
│   └── cities_gbk.csv
└── scripts/
    ├── build_nuitka.bat
    └── build_pyinstaller.bat
```

## Tech Stack

- **GUI**: tkinter
- **Data**: pandas
- **Excel I/O**: openpyxl (.xlsx), xlrd (.xls)

## License

[MIT](LICENSE)
