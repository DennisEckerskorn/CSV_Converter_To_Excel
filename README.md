# CSV to Excel Converter with CustomTkinter

A modern and user-friendly application to convert CSV files into Excel files with additional transformations, summaries, and follow-up call analysis.  
This tool is designed for employees to easily review call data, check answered and missed calls, and ensure proper follow-ups.

---

## ✨ Features

### ✅ CSV to Excel Conversion
- Converts `.csv` files (tab-separated) into Excel `.xlsx` files.
- Adds a polished graphical interface using `CustomTkinter`.

### 🔄 Data Transformations
- Adds a new column: **Time + 1h** – time incremented by 1 hour.
- Converts boolean values to **"Yes"/"No"** for readability.
- Adds an **Outgoing Calls** column based on the `Inbound` field.

### 📊 Summary Sheet
Automatically generates a `Summary` sheet in the Excel file with the following metrics:
- Total Calls
- Incoming Calls
- Outgoing Calls
- Answered Calls
- Missed Calls
- Average Callback Delay (minutes)

### 📞 Callback Analysis – NEW!
Introduced in the latest version:
- New sheet: **Callbacks**
- Detects missed calls (not answered) and checks if a follow-up call was made.
- Calculates the **time it took to return each missed call**.
- Shows:
  - Date of missed call
  - Callback time
  - Delay in minutes
  - Caller number and user name

🟥 **Conditional Formatting**:  
If the delay exceeds **40 minutes**, the corresponding cell in "Delay (minutes)" is automatically highlighted in red.

---

## 🖥️ Portable Executable

- The app is bundled as `CSVConversion.exe`.
- Runs on any Windows machine **without needing to install Python**.

---

## 🚀 Download and Run

1. Go to the [Releases](https://github.com/<your-username>/<your-repo>/releases) section of this repository.
2. Download the file **`CSVConversion.exe`**.
3. Double-click the file to launch the application.

---

## 🧭 Usage

1. Open the app (`CSVConversion.exe`).
2. Click **"Select File"** to choose your `.csv` file (tab-delimited).
3. Click **"Select Path"** to choose where to save the `.xlsx`.
4. Click **"Convert File"**.
5. Open the resulting Excel file to review:

- The `Calls` sheet with full detail.
- The `Summary` sheet with metrics.
- The `Callbacks` sheet with follow-up analysis.

---

## 🧰 Requirements (Only for script version)

If you want to run the Python script manually:

### 🔧 Requirements
- Python 3.8+
- Install dependencies with:

```bash
pip install -r requirements.txt
