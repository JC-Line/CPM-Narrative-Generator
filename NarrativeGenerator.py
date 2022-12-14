import datetime
from openpyxl import load_workbook
from fpdf import FPDF


# Activity class definition
class Activity:
    def __init__(self,id,description):
        self.id = id
        self.description = description
        self.successors = set()

    # Find Successors
    # Pass worksheet converted into values
    def get_successors(self,successorWS):
        for row in successorWS:
            if row[0] != self.id:
                continue
            else:
                self.successors.add(row[1])





# Create a function that will return list of all values within an excel sheet
#   Variables: Worksheet Name
#   Table Header Row Number
#   List of column numbers to keep
def get_ws_vals(wb, ws_name, header_row, col_keep):
    ws = wb[ws_name]
    # Variables could be pulled in from TXT file to be dynamic
    columnsToKeep = col_keep
    rowMin = header_row
    rowMax = len(ws['A'])
    colMax = max(columnsToKeep)
    colMin = min(columnsToKeep)

    data = []
    dataHeader = []
    for rowCount, row in enumerate(ws.iter_rows(rowMin, rowMax, colMin, colMax, True), rowMin):

        tempList = []
        for colCount, value in enumerate(row, colMin):
            if colCount in columnsToKeep:
                if rowCount == rowMin:
                    dataHeader.append(value)
                else:
                    tempList.append(value)
        data.append(tempList)
    data.pop(0)
    return dataHeader, data

def print_table(colWidths,dataHeader,data):
    dateFormat = '%x'
    tableRowHeight = 5
    pdf.set_xy(float(10), float(10))
    # Print Head of Table
    pdf.set_font('', 'B', 10)
    for count, value in enumerate(dataHeader):
        x = value
        if count > 0:
            pdf.set_xy(float(sum(colWidths[0:count])+10), float(10))

        if type(value) == type(datetime.datetime.now()):
            x = value.strftime(dateFormat)

        if (float(len(str(x)))*1.8) > colWidths[count]:
            pdf.multi_cell(colWidths[count], tableRowHeight, str(x), 0, 'C')
            continue

        pdf.multi_cell(colWidths[count], tableRowHeight*3, str(x), 0, 'C',)

    # Print Body of Table
    pdf.set_font('', '', 8)
    for row in data:
        for count, value in enumerate(row):
            if str(value) == 'None':
                x = '-'
            else:
                x = value
            if type(value) == type(datetime.datetime.now()):
                x = value.strftime(dateFormat)
            pdf.cell(colWidths[count], tableRowHeight, str(x), 1, 0)
        pdf.ln(tableRowHeight)

















# Lane Header #

# Lane Footer #

# Written Part of narrative #
#   -Bring in using a manually generated txt file with written narrative sections
#   -Check for existing txt file in current folder, if it doesnt exist create a new file template


# Create
wb = load_workbook('3462-FDOT-Owner-11.xlsx')
oldWB = load_workbook('3462-FDOT-Owner-11.xlsx')
activityDataHeader, activityData = get_ws_vals(wb, "TASK", 2, [1, 8, 9, 5])
predDataHeader, predData = get_ws_vals(wb,"TASKPRED",2,[1,2,3,10,11])
oldActivityDataHeader, oldActivityData = get_ws_vals(oldWB, "TASK", 2, [1, 8, 9, 5])
oldPredDataHeader, oldPredData = get_ws_vals(oldWB,"TASKPRED",2,[1,2,3,10,11])

lastCutoff = "2022-10-16"


# Tables showing relationship and activity changes #
#   -Build a heading for all categories and create a null result
#   -Pull data from schedule comparison export csv

# Create lists showing new/deleted activities
activityDataIDS = {id[0] for id in activityData}
oldActivityDataIDS = {id[0] for id in oldActivityData}

##TEST
activityDataIDS.remove("BASE640")
activityDataIDS.remove("R610")

addedIDS = activityDataIDS - oldActivityDataIDS
deletedIDS = oldActivityDataIDS - activityDataIDS

activityDataAdded = [row for row in activityData if row[0] in addedIDS]
activityDataDeleted = [row for row in oldActivityData if row[0] in deletedIDS]
##

# Create lists of added and deleted relationships
predList = [[id] for id in activityDataIDS]

for row in predList:
    id = row[0]
    if len(row) == 1:
        row.append(set())
    for row1 in predData:
        if row1[0] != id:
            continue
        else:
            successor = row1[1]
            row[1].add(successor)

testActivity = Activity("BASE640","Testing Activity Description")
testActivity.get_successors(predData)

pdf = FPDF()
pdf.add_page()

activityData = activityData[:1000]
oldActivityData = oldActivityData[:1000]

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
cWidths = [20, 130, 20, 20]
print_table(cWidths,oldActivityDataHeader,oldActivityData)

pdf.output('NarrativeTest.pdf', 'F')
