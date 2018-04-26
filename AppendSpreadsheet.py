import csv
from xlrd import open_workbook
from xlutils.copy import copy

with open("Results.tsv") as tsv, open("Callouts.xls"):
    print('Both files found')

    wb = open_workbook("Callouts.xls", formatting_info=True)
    # Make a writeable copy
    calloutsBook = copy(wb)
    currentSheet = calloutsBook.get_sheet('Callouts 2018')

    # Find which row to start appending spreadsheet
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
        i += 1

    print('Will append spreadsheet from row #' + str(startingRowNumber - 1))

    # Skips first row of TSV file due to column headers
    iterResults = iter(csv.reader(tsv, dialect="excel-tab"))
    next(iterResults)

    # Extracting data from TSV file
    for row in iterResults:
        # Getting data from the TSV file
        ticketID = row[0]
        ticketCreated = row[15]
        if row[20] == '':    # If no Nagios alarm, use ticket's subject - problem when ceph-mon1-5 callout
            alarm = row[2]
        else:
            alarm = row[20]

    # Putting data into spreadsheet

    calloutsBook.save("Callouts.xls")
    print('Spreadsheet changes saved')





