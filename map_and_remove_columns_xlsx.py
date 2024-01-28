import pandas as pd
import sys

def col_index_to_letter(index):
    """Convert a column index to an Excel column letter."""
    letter = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

def remove_columns_from_excel(file_path, columns_to_remove):
    """Remove specified columns from an Excel file."""
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Map column indices to letters
    column_index_to_letter_map = {i: col_index_to_letter(i) for i in range(1, len(df.columns) + 1)}

    # Find the corresponding column names to remove
    columns_to_remove_names = [name for index, name in enumerate(df.columns) 
                               if column_index_to_letter_map[index + 1] in columns_to_remove]

    # Remove the specified columns
    df_modified = df.drop(columns=columns_to_remove_names)

    return df_modified

if __name__ == "__main__":
    # Check if columns to remove are provided as arguments
    if len(sys.argv) < 3:
        print("Error: Please provide the Excel file path and the columns to remove.")
        sys.exit(1)

    # Excel file path
    file_path = sys.argv[1]

    # Columns to remove
    columns_to_remove = sys.argv[2:]

    # Remove columns and save the modified file
    df_modified = remove_columns_from_excel(file_path, columns_to_remove)
    modified_file_path = file_path.replace('.xlsx', '_modified.xlsx')
    df_modified.to_excel(modified_file_path, index=False)

    print(f"Modified file saved as {modified_file_path}")


# Example usage:
# python map_and_remove_columns_xlsx.py path_to_excel.xlsx A H J AB AC AG AH AI AJ AK AL AM AN AO AP
