import datetime
from openpyxl import load_workbook
from fpdf import FPDF


import Activity as Actfile
Activity = Actfile.Activity

def print_table(colWidths, dataHeader, data):
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


from Table import WorksheetTable
old_FileName = '3462-FDOT-Owner-11.xlsx'
new_FileName = '3462-FDOT-Owner-12.xlsx'
oldActivityData = WorksheetTable(old_FileName,"TASK")
oldPredData = WorksheetTable(old_FileName,"TASKPRED")
newActivityData = WorksheetTable(new_FileName,"TASK")
newPredData = WorksheetTable(new_FileName,"TASKPRED")


# Create activity dictionary {ID : Activity Obj}
newActivityDict = {}
for row in newActivityData.data:
    id = row[newActivityData.header['Activity ID']]
    newActivityDict.update({id : Activity(row,newActivityData.header)})

oldActivityDict = {}
for row in oldActivityData.data:
    id = row[oldActivityData.header['Activity ID']]
    oldActivityDict.update({id : Activity(row,oldActivityData.header)})

for row in oldPredData.data:
    key = row[oldPredData.header['Predecessor']]
    successor = row[oldPredData.header['Successor']]
    oldActivityDict[key].successors.add(successor)

for row in newPredData.data:
    key = row[newPredData.header['Predecessor']]
    successor = row[newPredData.header['Successor']]
    newActivityDict[key].successors.add(successor)





























pdf = FPDF()
pdf.add_page()

activityData = activityData[:1000]
oldActivityData = oldActivityData[:1000]

tableRowHeight = 5
pdf.set_font('Arial')

# Temp fix for row width
cWidths = [20, 130, 20, 20]
print_table(cWidths, oldActivityDataHeader, oldActivityData)

pdf.output('NarrativeTest.pdf', 'F')
