# Excel Division App

A simple tool to divide an Excel file into multiple smaller files based on a specified column.

## Usage

1. Run the executable file `excel_division_by_unique_values.exe` or run `excel_division_app.py` from the command line.

2. Select the Excel file you want to divide, and define the worksheet name.

3. Select the column which you want to use to divide the Excel file.

4. Enter the constant part of the new Excel file names (e.g. '_Student' for '1_Student.xlsx', '2_Student.xlsx', etc.)

5. Select the folder you want the files to be saved to.

6. Click the "Submit" button to divide the Excel file. You are done.

## Requirements

- Python 3.6+
- openpyxl
- tkinter
- Pillow
- requests

## Installation

- Install Python from the official website: https://www.python.org/downloads/
- Install the required packages by running the following command in the command line:

```bash
pip install openpyxl
```

## Author

- Mate Mihalovits, PhD - mmateo1120@gmail.com

## Version

- 1.0.0
