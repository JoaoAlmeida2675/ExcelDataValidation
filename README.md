# Excel Data Validation Automation

Automate the validation of Excel data by checking for blank or empty cells in a specified range of lines and columns using Python with the openpyxl library. This script streamlines the process, making it efficient and easy to identify and manage empty cells in your Excel files.

## Introduction

Managing and validating data in Excel spreadsheets can be a time-consuming task, especially when dealing with large datasets. This Python script simplifies the process by automating the detection of empty cells within a specified range. It utilizes the openpyxl library to interact with Excel files, providing a seamless solution for Excel data validation.

Check [INFO.md](INFO.md) for more detailed information about the code.

## How It Works

### 1. Function `get_column_index_from_letter`:

This function converts a letter or set of letters to uppercase and calculates the corresponding column index. It ensures consistency in identifying columns, where A=1, B=2, etc. By receiving a string `letter` as a parameter, it iterates over each character, calculates the corresponding index in the alphabet, and returns the column index.

### 2. Open Excel and Load the File:

Uses the openpyxl library to open and load an Excel workbook from the specified path. If successful, it proceeds to the next steps.

### 3. Select Sheet and Specify Columns to Check:

Selects a specific sheet in the workbook and specifies a set of columns to be checked. The script allows customization of the sheet name, columns, rows, and the row containing column names.

### 4. Check Empty Cells:

Iterates over the specified columns and rows to check if the cells are empty. It identifies empty cells, records the column names and messages for empty columns in the `empty_column_names` list.

### 5. Create the `.txt` File with Messages for Empty Columns:

Creates a text file named "MissingDataExcel.txt" with messages for empty columns. The script uses Python's file handling to write the messages to the file.

### 6. Close Excel:

Closes the Excel workbook, ensuring proper resource management after the validation process.

## Usage

1. Install the openpyxl library:

   ```bash
   pip install openpyxl
   ```

2. Customize the script:

   - Set the Excel file path:

     ```python
     excel_file_path = "Your\\Path\\To\\Excel\\File.xlsx"
     ```

   - Specify the sheet name:

     ```python
     sheet_name = "Your\\Sheet\\Name"
     ```

   - Define the range of columns to check, rows to check, and the row containing column names.

   - Set the .txt file path:

     ```python
     file_name = "Your\\Save\\Path\\File.txt"
     ```
     
4. Run the script:

   ```bash
   python ValidateExcelData.py
   ```


Feel free to modify the script according to your specific Excel file and validation requirements.
