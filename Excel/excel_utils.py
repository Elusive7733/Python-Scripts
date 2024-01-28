# excel_utils.py
def col_letter_to_index(letter):
    """
    Convert an Excel column letter (e.g., 'A', 'Z', 'AA') to a zero-based column index.
    """
    expn = 0  # Initialize exponent for base-26 calculation
    column_index = 0  # Initialize the column index accumulator

    # Process each character in reverse (to handle multi-letter columns correctly)
    for char in reversed(letter):
        # Convert character to uppercase and calculate its base-26 positional value
        # 'A' -> 1, 'B' -> 2, ..., 'Z' -> 26
        base26_val = ord(char.upper()) - 65 + 1

        # Add the character's contribution to the column index
        # For single-letter, expn is 0 (26^0 = 1), for 'AA', 'A' contributes as 26^1
        column_index += base26_val * (26 ** expn)

        # Increment exponent for next character (moving left in the column string)
        expn += 1

    # Adjust for zero-based indexing (Python) from one-based (Excel)
    return column_index - 1
