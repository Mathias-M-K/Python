def get_column_pos(workbook, text_to_search_for):
    for y in range(1, workbook.max_column + 1):

        if workbook.cell(row=1, column=y).value == text_to_search_for:
            return y