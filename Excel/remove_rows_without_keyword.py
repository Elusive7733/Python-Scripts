import pandas as pd
import sys
from excel_utils import col_letter_to_index  # Import the utility function

def remove_rows_without_keywords(file_path, column_letter, keywords):
    """Remove rows that do not contain any of the keywords in the specified column."""
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Convert column letter to zero-based index
    column_index = col_letter_to_index(column_letter)

    # Check if column index is valid
    if column_index >= len(df.columns):
        raise ValueError("Column letter is out of range.")

    # Column name
    column_name = df.columns[column_index]

    # Prepare keywords for comparison (lowercase and stripped)
    keywords = [k.lower().strip() for k in keywords]

    # Filter rows
    df_filtered = df[df[column_name].astype(str).str.lower().str.strip().apply(lambda x: any(k in x for k in keywords))]

    return df_filtered

if __name__ == "__main__":
    # Check if necessary arguments are provided
    if len(sys.argv) < 4:
        print("Error: Please provide the Excel file path, column letter, and at least one keyword.")
        sys.exit(1)

    # Arguments
    file_path = sys.argv[1]
    column_letter = sys.argv[2]
    keywords = sys.argv[3:]

    # Remove rows and save the modified file
    df_modified = remove_rows_without_keywords(file_path, column_letter, keywords)
    modified_file_path = file_path.replace('.xlsx', '_rows_filtered.xlsx')
    df_modified.to_excel(modified_file_path, index=False)

    print(f"Modified file saved as {modified_file_path}")


# Example usage:
    # python remove_rows_without_key.py "path_to_excel.xlsx" "D" "Karachi" "Lahore" "Islamabad"
