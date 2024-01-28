import pandas as pd
import sys
from excel_utils import col_letter_to_index

def remove_rows_without_keyword(file_path, column_letter, keyword):
    """Remove rows that do not contain the keyword in the specified column."""
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Convert column letter to index
    column_index = col_letter_to_index(column_letter)

    # Check if column index is valid
    if column_index >= len(df.columns):
        raise ValueError("Column letter is out of range.")

    # Column name
    column_name = df.columns[column_index]

    # Filter rows
    df_filtered = df[df[column_name].astype(str).str.lower().str.strip().str.contains(keyword.lower().strip(), na=False)]

    return df_filtered

if __name__ == "__main__":
    # Check if necessary arguments are provided
    if len(sys.argv) != 4:
        print("Error: Please provide the Excel file path, column letter, and keyword.")
        sys.exit(1)

    # Arguments
    file_path = sys.argv[1]
    column_letter = sys.argv[2]
    keyword = sys.argv[3]

    # Remove rows and save the modified file
    df_modified = remove_rows_without_keyword(file_path, column_letter, keyword)
    modified_file_path = file_path.replace('.xlsx', '_rows_filtered.xlsx')
    df_modified.to_excel(modified_file_path, index=False)

    print(f"Modified file saved as {modified_file_path}")
