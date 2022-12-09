import datetime
from openpyxl import load_workbook
from fpdf import FPDF

pdf = FPDF()
pdf.add_page()

# Lane Header #

# Lane Footer #

# Written Part of narrative #
#   -Bring in using a manually generated txt file with written narrative sections
#   -Check for existing txt file in current folder, if it doesnt exist create a new file template

# Lists showing started/completed activities #
#   -Generate from activity csv export

wb = load_workbook('schedule.xlsx')
activityWS = wb['TASK']


# Indexes [numbers(1-26), letters(A-Z)]
letterNumberEqList = []
i = 0
asciiCount = 64
while i < 26:
    asciiCount += 1
    tempList = [i, chr(asciiCount)]
    i += 1
    letterNumberEqList.append(tempList)


# Variables could be pulled in from TXT file to be dynamic
tableHeaderRow = 2
columnsToKeep = [1, 5, 8, 9]
rowMin = tableHeaderRow
rowMax = len(activityWS['A'])
colMax = max(columnsToKeep)
colMin = min(columnsToKeep)


activityData = []
activityDataHeader = []
for rowCount, row in enumerate(activityWS.iter_rows(rowMin, rowMax, colMin, colMax, True), rowMin):

    tempList = []
    for colCount, value in enumerate(row, colMin):
        if colCount in columnsToKeep:
            if rowCount == rowMin:
                activityDataHeader.append(value)
            else:
                tempList.append(value)
    activityData.append(tempList)
activityData.pop(0)


lastCutoff = "2022-10-16"


# Tables showing relationship and activity changes #
#   -Build a heading for all categories and create a null result
#   -Pull data from schedule comparison export csv

activityData = activityData[:1000]


# Iterate through columns and return optimal row widths based on string length
#   returns list of int
dateFormat = '%x'
tableRowWidths = []
test = activityData[1]
for column in test:
    tableRowWidths.append(1)

for row in activityData:
    for count, value in enumerate(row):
        x = value
        if type(x) == type(datetime.datetime.now()):
            x = value.strftime(dateFormat)

        if len(str(x)) > tableRowWidths[count]:
            tableRowWidths[count] = len(str(x))


tableRowHeight = 5
pdf.set_font('Arial')

# Temp fix for row width
tableRowWidths = [20,130,20,20]
pdf.set_xy(float(10),float(10))
# Print Head of Table
pdf.set_font('', 'B', 10)
for count, value in enumerate(activityDataHeader):
    x = value
    if count > 0:
        pdf.set_xy(float(sum(tableRowWidths[0:count])+10),float(10))

    if type(value) == type(datetime.datetime.now()):
        x = value.strftime(dateFormat)

    if (float(len(str(x)))*1.8) > tableRowWidths[count]:
         pdf.multi_cell(tableRowWidths[count],tableRowHeight,str(x),0,'C')
         continue

    pdf.multi_cell(tableRowWidths[count], tableRowHeight*3, str(x), 0,'C',)

# Print Body of Table
pdf.set_font('', '', 8)
for row in activityData:
    for count, value in enumerate(row):
        if str(value) == 'None':
            x = '-'
        else:
            x = value
        if type(value) == type(datetime.datetime.now()):
            x = value.strftime(dateFormat)
        pdf.cell(tableRowWidths[count], tableRowHeight, str(x), 1, 0)
    pdf.ln(tableRowHeight)
pdf.output('NarrativeTest.pdf', 'F')
