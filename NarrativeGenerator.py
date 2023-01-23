import wordDoc
import datetime
from openpyxl import load_workbook
import Activity as Actfile
from Table import WorksheetTable

old_FileName = '3462-FDOT-JAN.xlsx'
new_FileName = '3462-FDOT-Owner-DEC.xlsx'
oldActivityData = WorksheetTable(old_FileName, "TASK")
oldPredData = WorksheetTable(old_FileName, "TASKPRED")
newActivityData = WorksheetTable(new_FileName, "TASK")
newPredData = WorksheetTable(new_FileName, "TASKPRED")

c = Actfile.Comparison(
    oldActivityData, oldPredData, newActivityData, newPredData)





# Required Inputs
dataDate = datetime.datetime(2023, 1, 15)
previousDataDate = datetime.datetime(2022, 12, 18)
nextDataDate = dataDate + datetime.timedelta(days=30)


# Create Word Document
WD = wordDoc.WordDoc
narrative = WD("NarrativeTemplate2.docx")
startedActivities = c.actBetween(previousDataDate, dataDate, False, 3)
completedActivities = c.actBetween(
    previousDataDate, dataDate, False, 4)
startedActivities = c.concatActInfo(startedActivities, ["Activity Name"])
completedActivities = c.concatActInfo(completedActivities, ["Activity Name"])
next30CritData = c.actBetween(
    dataDate, nextDataDate, True)
next30CritData = c.concatActInfo(next30CritData, ["Activity Name"])

narrative.actSC(startedActivities, completedActivities)
narrative.next30CP(next30CritData)
narrative.activityMods(c.changed_names, c.added_activities, c.deleted_activities,
                       c.added_successors, c.deleted_successors, c.changed_durations)


narrative.saveFile("wordTest.docx")
