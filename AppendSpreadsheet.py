import csv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Side, Border
from datetime import datetime

with open("Results.tsv") as tsv:
    print('Both files found')

    wb = load_workbook("Callouts.xlsx")
    currentSheet = wb['Callouts 2018']

    # Find which row to start appending spreadsheet
    alarmColumn = currentSheet['A']

    print('Working out where to start appending spreadsheet')
    startAppending = False
    i = 2    # Start at 2 as first result will always be 'column header'
    while not startAppending:
        if alarmColumn[i].value is None:
            startAppending = True
            startingRowNumber = i + 1
        i += 1

    print('Will append spreadsheet from row #' + str(startingRowNumber))

    # Skips first row of TSV file due to column headers
    iterResults = iter(csv.reader(tsv, dialect="excel-tab"))
    next(iterResults)

    # Extracting data from TSV file
    i = 0
    nagiosFiller = "Nagios issued and cleared service alarm"
    for row in iterResults:
        ticketID = int(row[0])

        ticketCreated = datetime.strptime(row[15], '%Y-%m-%d %H:%M:%S')
        dateCreated = ticketCreated.strftime('%d/%m/%Y')
        timeCreated = ticketCreated.strftime('%H:%M:%S')

        alarm = row[2]
        if nagiosFiller in alarm:
            alarm = alarm[len(nagiosFiller) + 1:]    # Just leaves Nagios alarm

        # Putting data into spreadsheet
        currentRow = startingRowNumber + i
        currentSheet.cell(row=currentRow, column=1, value=alarm)          # Alarm name
        currentSheet.cell(row=currentRow, column=3, value=dateCreated)    # Date issued
        currentSheet.cell(row=currentRow, column=4, value=timeCreated)    # Time issued
        # RT query
        currentSheet.cell(row=currentRow, column=5, value=ticketID).alignment = Alignment(horizontal='center')
        currentSheet.cell(row=currentRow, column=5).hyperlink = 'https://helpdesk.gridpp.rl.ac.uk/Ticket/Display' \
                                                                '.html?id=' + str(ticketID)
        i += 1

    # Merge cells
    currentDate = datetime.now().strftime('%d-%b')
    currentSheet.cell(row=startingRowNumber, column=12, value=currentDate).alignment = Alignment(horizontal='center', vertical='center')
    currentSheet.merge_cells(start_row=startingRowNumber, start_column=12, end_row=currentRow, end_column=12)

    # Setting inner borders
    rows = currentSheet.iter_rows(min_row=startingRowNumber, min_col=1, max_row=currentRow, max_col=17)
    innerBorderStyle = Side(border_style='thin', color='FF000000')
    innerBorderFormat = Border(left=innerBorderStyle, right=innerBorderStyle, top=innerBorderStyle, bottom=innerBorderStyle)
    for row in rows:
        for cell in row:
            cell.border = innerBorderFormat

    # Setting outer border
    # Code found at: https://stackoverflow.com/questions/34520764/apply-border-to-range-of-cells-using-openpyxl
    # Written by Yaroslav Admin, edited by Adam Stewart
    outerRows = currentSheet.iter_rows(min_row=startingRowNumber, min_col=1, max_row=currentRow, max_col=17)
    outerBorderStyle = Side(border_style='medium', color='FF000000')
    outerRows = list(outerRows)
    max_y = len(outerRows) - 1
    for pos_y, cells in enumerate(outerRows):
        max_x = len(cells) - 1  # index of the last cell
        for pos_x, cell in enumerate(cells):
            border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom)

            # Checking if an edge cell
            if pos_x == 0:
                border.left = outerBorderStyle
            if pos_x == max_x:
                border.right = outerBorderStyle
            if pos_y == 0:
                border.top = outerBorderStyle
            if pos_y == max_y:
                border.bottom = outerBorderStyle

            # Set new border only if it's one of the edge cells
            if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                cell.border = border

    wb.save("Callouts.xlsx")
    print('Spreadsheet changes saved')