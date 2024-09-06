import pandas as pd
from collections import defaultdict

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
    # Step 1: Read the input Excel files
    dfs = read_excel_files(input_files)

    # Step 2: Merge and remove duplicate columns
    merged_df = merge_unique_columns(dfs)

    # Step 3: Save the merged DataFrame to a new Excel file
    save_to_excel(merged_df, output_file)
    print(f"Merged Excel file saved to {output_file}")

# Example usage
if __name__ == "__main__":
    input_files = [r"C:\Users\g.fotopoulos\τεστ.xlsx", 
                   r"C:\Users\g.fotopoulos\τεστ1.xlsx"]  # Replace with your file paths
    output_file = r"C:\Users\g.fotopoulos\merged_output.xlsx"
    process_excel_files(input_files, output_file)
