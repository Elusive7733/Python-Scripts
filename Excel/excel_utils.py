# excel_utils.py
def col_letter_to_index(letter):
    """Convert an Excel column letter to a zero-based column index."""
    expn = 0
    column_index = 0
    for char in reversed(letter):
        column_index += (ord(char.upper()) - 65 + 1) * (26 ** expn)
        expn += 1
    return column_index - 1
