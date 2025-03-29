import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox


def add_hour_to_time_column(df):
    """Adds an hour to the 'time' column"""
    df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'], errors='coerce')  # Combine and convert to datetime
    if df['datetime'].isna().any():
        raise ValueError("Some rows have invalid 'Date' or 'Time' values.")
    df['datetime + 1h'] = df['datetime'] + pd.Timedelta(hours=1)  # Add 1 hour
    return df


def add_outbound_column(df):
    """Creates the 'outbound' column based on the 'Inbound' column"""
    df['Inbound'] = df['Inbound'].astype(bool)  # Ensure the column is boolean
    df['outbound'] = ~df['Inbound']
    df['Inbound'] = df['Inbound'].map({True: 'Yes', False: 'No'})
    df['outbound'] = df['outbound'].map({True: 'Yes', False: 'No'})
    return df


def reorder_and_select_columns(df):
    """Reorders and selects only the relevant columns"""
    # Convert 'Answered' to 'Yes'/'No'
    df['Answered'] = df['Answered'].map({True: 'Yes', False: 'No'})

    # Select and reorder columns
    columns = [
        'UserName', 'UserEmail', 'UserPhone', 'Source', 'SourceDetail',
        'Date', 'Time', 'datetime + 1h', 'Duration', 'Answered',
        'Inbound', 'outbound', 'Number', 'PhonebookName'
    ]
    df = df[columns].copy()  # Create a copy to avoid SettingWithCopyWarning

    # Rename columns for clarity
    df = df.rename(columns={
        'UserName': 'User Name',
        'UserEmail': 'User Email',
        'UserPhone': 'User Phone',
        'Source': 'Source',
        'SourceDetail': 'Source Detail',
        'Date': 'Date',
        'Time': 'Time',
        'datetime + 1h': 'Time + 1h',
        'Duration': 'Duration (s)',
        'Answered': 'Answered',
        'Inbound': 'Incoming Calls',
        'outbound': 'Outgoing Calls',
        'Number': 'Number',
        'PhonebookName': 'Phonebook Name'
    })
    return df


def create_summary(df, writer):
    """Creates a summary sheet in the Excel file"""
    total_calls = len(df)
    total_incoming = len(df[df['Incoming Calls'] == 'Yes'])
    total_outgoing = len(df[df['Outgoing Calls'] == 'Yes'])
    total_answered = len(df[df['Answered'] == 'Yes'])
    total_missed = total_calls - total_answered

    summary_data = {
        'Metric': ['Total Calls', 'Incoming Calls', 'Outgoing Calls', 'Answered Calls', 'Missed Calls'],
        'Count': [total_calls, total_incoming, total_outgoing, total_answered, total_missed]
    }
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)


def process_csv(input_path, output_path):
    """Reads the CSV file, applies transformations, and saves it as an Excel file"""
    try:
        df = pd.read_csv(input_path, sep='\t')
        required_columns = ['UserName', 'UserEmail', 'UserPhone', 'Source', 'SourceDetail',
                            'Date', 'Time', 'Duration', 'Answered', 'Inbound', 'Number', 'PhonebookName']
        if not all(col in df.columns for col in required_columns):
            raise ValueError("The CSV file does not contain the required columns.")

        df = add_hour_to_time_column(df)
        df = add_outbound_column(df)
        df = reorder_and_select_columns(df)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Calls', index=False)
            create_summary(df, writer)

        messagebox.showinfo("Success", f"The Excel file has been saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error during conversion: {e}")


def select_file():
    """Opens a dialog to select a CSV file"""
    file_path = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*"))
    )
    csv_path.set(file_path)


def select_save_path():
    """Opens a dialog to select the path to save the Excel file"""
    file_path = filedialog.asksaveasfilename(
        title="Save Excel file",
        defaultextension=".xlsx",
        filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*"))
    )
    excel_path.set(file_path)


def convert_file():
    """Converts the CSV file into an Excel file"""
    input_path = csv_path.get()
    output_path = excel_path.get()

    if not input_path or not output_path:
        messagebox.showwarning("Warning", "Please select a CSV file and a save path.")
        return

    process_csv(input_path, output_path)


# Create the main window using CustomTkinter
ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"

root = ctk.CTk()
root.title("CSV to Excel Converter")
root.geometry("500x400")

# Global variables for input and output paths
csv_path = ctk.StringVar()
excel_path = ctk.StringVar()

# Labels and input fields
ctk.CTkLabel(root, text="CSV File:", font=("Arial", 14)).pack(pady=10)
ctk.CTkEntry(root, textvariable=csv_path, width=400).pack(pady=5)
ctk.CTkButton(root, text="Select File", command=select_file).pack(pady=5)

ctk.CTkLabel(root, text="Save as (Excel):", font=("Arial", 14)).pack(pady=10)
ctk.CTkEntry(root, textvariable=excel_path, width=400).pack(pady=5)
ctk.CTkButton(root, text="Select Path", command=select_save_path).pack(pady=5)

# Button to convert the file
ctk.CTkButton(root, text="Convert File", command=convert_file, fg_color="green").pack(pady=20)

# Start the application
root.mainloop()
