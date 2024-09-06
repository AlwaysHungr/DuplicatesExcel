import pandas as pd
from collections import defaultdict
from tkinter import Tk, filedialog, messagebox, Button, Label
import os

# Function to read multiple Excel files and store them in a list of DataFrames
def read_excel_files(file_paths):
    dataframes = []
    for file in file_paths:
        df = pd.read_excel(file)
        dataframes.append(df)
    return dataframes

# Function to compare columns across DataFrames and remove duplicates
def merge_unique_columns(dfs):
    combined_df = pd.DataFrame()
    col_tracker = defaultdict(list)  # Track unique columns

    for df in dfs:
        for col in df.columns:
            col_data = df[col].dropna()  # Drop NaN values to avoid false matches on empty columns

            # Check if the column already exists in combined_df by comparing content
            is_duplicate = False
            for existing_col in combined_df.columns:
                if col_data.equals(combined_df[existing_col].dropna()):
                    is_duplicate = True
                    break

            if not is_duplicate:
                combined_df[col] = df[col]  # Add column if it is unique

    return combined_df

# Function to save the final combined DataFrame to a new Excel file
def save_to_excel(df, output_file):
    df.to_excel(output_file, index=False)

# Main function to process the files
def process_excel_files(input_files, output_file):
    try:
        # Step 1: Read the input Excel files
        dfs = read_excel_files(input_files)

        # Step 2: Merge and remove duplicate columns
        merged_df = merge_unique_columns(dfs)

        # Step 3: Save the merged DataFrame to a new Excel file
        save_to_excel(merged_df, output_file)
        messagebox.showinfo("Success", f"Merged Excel file saved to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI to select files
def select_input_files():
    file_paths = filedialog.askopenfilenames(title="Select Excel Files",
                                             filetypes=[("Excel files", "*.xlsx")])
    if file_paths:
        input_label.config(text=f"Selected {len(file_paths)} file(s)")
        return file_paths
    return []

def select_output_file():
    output_file = filedialog.asksaveasfilename(title="Save Merged Excel File",
                                               defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")])
    if output_file:
        output_label.config(text=f"Output file: {os.path.basename(output_file)}")
        return output_file
    return ""

def merge_files():
    input_files = select_input_files()
    if not input_files:
        return

    output_file = select_output_file()
    if not output_file:
        return

    process_excel_files(input_files, output_file)

# Initialize the GUI
root = Tk()
root.title("Excel Merger Tool")

# Input files label and button
input_label = Label(root, text="No input files selected", padx=10, pady=10)
input_label.pack()

output_label = Label(root, text="No output file selected", padx=10, pady=10)
output_label.pack()

# Merge button
merge_button = Button(root, text="Select Files and Merge", command=merge_files, padx=10, pady=10)
merge_button.pack()

# Run the GUI loop
root.geometry("400x200")
root.mainloop()
