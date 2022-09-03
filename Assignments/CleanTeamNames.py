from ToolBox import ToolScript

import re
import openpyxl

# load excel with its path
wrkbk = openpyxl.load_workbook("../Resources/nba.xlsx")
sh = wrkbk.active


def clean_text(text):
    # Removing anything that isn't a normal letter or number
    # Maybe we could remove parenthesis and numbers here, but just in case a team name actually contains a number
    # we'll stick to only removing number that exist between parenthesis
    cleantext = re.sub("[^A-Za-z0-9().\s]+", '', text)

    # If there is a start parenthesis, we'll remove it and anything behind it
    if "(" in cleantext:
        cleantext = cleantext[0:cleantext.index("(")]

    return cleantext


team_column_pos = ToolScript.getTeamColumnPos(sh)

# iterate through excel and display data
for i in range(2, sh.max_row + 1):
    print("\n")
    print("Row ", i)

    # Printing original value
    cell_content = sh.cell(row=i, column=team_column_pos).value
    print("Original String -> " + cell_content)

    # Printing clean value
    cleanCellContent = clean_text(cell_content)
    print('Clean String -> ' + cleanCellContent)

    # Replaceing the old value with the corected one
    sh.cell(row=i, column=team_column_pos).value = cleanCellContent;

# Saving file
wrkbk.save(filename="../output/Assignment_1.xlsx")
