# CSV to Excel Converter with CustomTkinter

A modern and user-friendly application to convert CSV files into Excel files with additional transformations and summaries. This tool is designed for employees to easily review call data, check answered and missed calls, and ensure proper follow-ups.

---

## Features

- **CSV to Excel Conversion**: Converts CSV files into Excel format with added transformations.
- **CustomTkinter Interface**: A modern and polished graphical user interface.
- **Data Transformations**:
  - Adds a new column `Time + 1h` with the time incremented by 1 hour.
  - Converts `Answered` to "Yes" or "No" for better readability.
  - Adds an `Outgoing Calls` column based on the `Inbound` column.
- **Summary Sheet**: Automatically generates a summary sheet in the Excel file with metrics like:
  - Total Calls
  - Incoming Calls
  - Outgoing Calls
  - Answered Calls
  - Missed Calls
- **Portable Executable**: Easily run the program on any Windows machine without needing Python installed.

---

## Download and Run

You can download the pre-built executable directly from this repository and run it on your Windows machine:

1. Go to the **[Releases Section](https://github.com/DennisEckerskorn/CSV_Converter_To_Excel/releases/download/CSV_Converter/CSVConversion.exe)** of this repository.
2. Download the file `CSVConversion.exe`.
3. Double-click the `.exe` file to launch the application.

Alternatively, you can download the `.exe` file directly from the repository [here](./CSVConversion.exe).

---

## Usage

1. Open the application by double-clicking the `.exe` file.
2. Click **"Select File"** to choose the CSV file you want to convert.
3. Click **"Select Path"** to choose where to save the Excel file.
4. Click **"Convert File"** to generate the Excel file.
5. Open the generated Excel file to review the data and summary.

---

## Requirements

No additional requirements are needed to run the `.exe` file. Simply download and run it on any Windows machine.

If you want to run the Python script directly, you will need:
- **Python 3.8+**
- **Dependencies**:
  - `pandas`
  - `openpyxl`
  - `customtkinter`

Install dependencies using:
```bash
pip install -r requirements.txt
