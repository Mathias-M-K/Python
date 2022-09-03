


def getTeamColumnPos(workbook):
    for y in range(1, workbook.max_column + 1):

        if workbook.cell(row=1, column=y).value == "team":
            return y