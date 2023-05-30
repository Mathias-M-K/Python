import openpyxl


def is_pattern_one(current_line, prev_line):
    return current_line == prev_line + 1


def is_pattern_two(current_line, prev_line):
    return current_line == prev_line - 1


def do_stuff():

    # opening workbook and getting the active sheet. I guess it's always number one? No idea why
    workbook = openpyxl.load_workbook("Book2.xlsx")
    sheet = workbook.active

    # Creating a list to store our pattern
    pattern_content = []

    # To recognize a pattern, we need to know the content of previous row and if we are already working a pattern
    prev_row = None
    pattern_active = False
    for i in range(1, sheet.max_row + 2):

        # if it's the first row, we set prev_row and skip
        if i == 1:
            prev_row = sheet.cell(row=i, column=1).value
            continue

        row = sheet.cell(row=i, column=1).value

        # If current and prev values match any of our patterns, we save the value in our list
        if is_pattern_one(row, prev_row) or is_pattern_two(row, prev_row):
            pattern_active = True
            pattern_content.append(prev_row)
        else:
            if pattern_active:
                pattern_content.append(prev_row)
                print("Pattern found", pattern_content)
                pattern_content = []
                pattern_active = False

        prev_row = row


# Does stuff
do_stuff()
