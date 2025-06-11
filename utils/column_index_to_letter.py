def column_index_to_letter(index):
    """Преобразует 1 → 'A', 2 → 'B', ..., 27 → 'AA'"""
    result = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result
