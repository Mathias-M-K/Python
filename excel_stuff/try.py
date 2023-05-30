import openpyxl

from openpyxl.workbook import Workbook

column = 1


def is_pattern_one(current_line, prev_line):
    return current_line == prev_line + 1


def is_pattern_two(current_line, prev_line):
    return current_line == prev_line - 1


def write_to_output(workbook, data):
    global column
    i = 1
    for dataEntry in data:
        index_cell = workbook.cell(row=i, column=column)
        value_cell = workbook.cell(row=i, column=column+1)

        index_cell.value = dataEntry[0]
        value_cell.value = dataEntry[1]
        i += 1
        
    column += 3


def save_workbook(workbook):
    workbook.save("output_my.xlsx")


def do_stuff():
    # opening workbook and getting the active sheet. I guess it's always number one? No idea why
    workbook = openpyxl.load_workbook("Book2.xlsx")
    sheet = workbook.active

    workbook_output = Workbook()
    sheet_out = workbook_output.active

    # Creating a list to store our pattern
    pattern_content = []

    # To recognize a pattern, we need to know the content of previous row and if we are already working a pattern
    prev_row = None
    pattern_active = False
    for i in range(1, sheet.max_row + 2):

        # if it's the first row, we set prev_row and skip
        if i == 1:
            prev_row = (sheet.cell(row=i, column=1).value, sheet.cell(row=i, column=2).value)
            continue

        index = sheet.cell(row=i, column=1).value
        value = sheet.cell(row=i, column=2).value

        # If current and prev values match any of our patterns, we save the value in our list
        if is_pattern_one(index, prev_row[0]) or is_pattern_two(index, prev_row[0]):
            pattern_active = True
            pattern_content.append(prev_row)
        else:
            if pattern_active:
                pattern_content.append(prev_row)
                print("Pattern found", pattern_content)
                write_to_output(sheet_out, pattern_content)
                pattern_content = []
                pattern_active = False

        prev_row = (index, value)

    save_workbook(workbook_output)


# Does stuff
do_stuff()
