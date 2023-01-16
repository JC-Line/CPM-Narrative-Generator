import datetime
from openpyxl import load_workbook
import Activity as Actfile
from Table import WorksheetTable

old_FileName = '3462-FDOT-Owner-11.xlsx'
new_FileName = '3462-FDOT-Owner-12.xlsx'
oldActivityData = WorksheetTable(old_FileName,"TASK")
oldPredData = WorksheetTable(old_FileName,"TASKPRED")
newActivityData = WorksheetTable(new_FileName,"TASK")
newPredData = WorksheetTable(new_FileName,"TASKPRED")

comparison = Actfile.Comparison(oldActivityData,oldPredData,newActivityData,newPredData)



import wordDoc

WD = wordDoc.WordDoc
narrative = WD("NarrativeTemplate2.docx")

startedActivities = ('test','data','here')
completedActivities = ('test','data','here')
next30Crit = ('act1','act2','act3')

# NEED TO CREATE A FUNCTION TO GENERATE STARTED AND COMPLETED ACTIVITIES
narrative.actSC(startedActivities,completedActivities)
narrative.next30CP(next30Crit)
actName = comparison.changed_names
narrative.table_ActivityName(actName[0],actName[1])

narrative.saveFile("wordTest.docx")

