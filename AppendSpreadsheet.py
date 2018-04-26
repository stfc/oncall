import csv
from xlrd import open_workbook
from xlutils.copy import copy

with open("Results.tsv") as tsv, open("Callouts.xls"):
    print('Both files found')

    # Skips first row with column headers
    iterResults = iter(csv.reader(tsv, dialect="excel-tab"))
    next(iterResults)

    wb = open_workbook("Callouts.xls", formatting_info=True)
    # Make a writeable copy - not needed at this point
    calloutsBook = copy(wb)
    currentSheet = calloutsBook.get_sheet('Callouts 2018')

    # Find which row to start appending
    firstRow = currentSheet.row(0)
    readSheet = wb.sheet_by_name('Callouts 2018')
    alarmColumn = readSheet.col_values(0)

    print('Working out where to start appending spreadsheet')
    startAppending = False
    i = 1    # Start at 1 as first result will always be 'Alarm name'
    while not startAppending:
        if readSheet.col(0)[i].value == '':
            startAppending = True
            startingRowNumber = i + 1
        i = i + 1

    print('Will append spreadsheet from row #' + startingRowNumber - 1)

    for row in iterResults:
        # Getting data from the TSV file
        ticketID = row[0]
        ticketCreated = row[15]
        if row[20] == '':    # If no Nagios alarm, use ticket's subject - problem when ceph-mon1-5
            alarm = row[2]
        else:
            alarm = row[20]

    calloutsBook.save("Callouts.xls")
    print(currentSheet.name)





