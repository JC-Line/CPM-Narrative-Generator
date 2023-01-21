import wordDoc
import datetime
from openpyxl import load_workbook
import Activity as Actfile
from Table import WorksheetTable

old_FileName = 'test_old_schedule.xlsx'
new_FileName = 'test_new_schedule.xlsx'
oldActivityData = WorksheetTable(old_FileName, "TASK")
oldPredData = WorksheetTable(old_FileName, "TASKPRED")
newActivityData = WorksheetTable(new_FileName, "TASK")
newPredData = WorksheetTable(new_FileName, "TASKPRED")

newPredData.data = newPredData.data[:900]

comparison = Actfile.Comparison(
    oldActivityData, oldPredData, newActivityData, newPredData)


WD = wordDoc.WordDoc
narrative = WD("NarrativeTemplate2.docx")

# Required Inputs
dataDate = datetime.datetime(2022, 12, 14)
previousDataDate = datetime.datetime(2022, 11, 14)
startedActivities = comparison.actBetween(previousDataDate, dataDate, False, 3)
completedActivities = comparison.actBetween(previousDataDate, dataDate, False, 4)

next30CritData = comparison.actBetween(
    dataDate, dataDate + datetime.timedelta(days=30), True)

narrative.actSC(startedActivities, completedActivities)
narrative.next30CP(next30CritData)
narrative.activityMods(comparison.changed_names, comparison.added_activities, comparison.deleted_activities,
                       comparison.added_successors, comparison.deleted_successors, comparison.changed_durations)


narrative.saveFile("wordTest.docx")
