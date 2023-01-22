import wordDoc
import datetime
from openpyxl import load_workbook
import Activity as Actfile
from Table import WorksheetTable

old_FileName = '3462-FDOT-Owner-DEC.xlsx'
new_FileName = '3462-FDOT-JAN.xlsx'
oldActivityData = WorksheetTable(old_FileName, "TASK")
oldPredData = WorksheetTable(old_FileName, "TASKPRED")
newActivityData = WorksheetTable(new_FileName, "TASK")
newPredData = WorksheetTable(new_FileName, "TASKPRED")

comparison = Actfile.Comparison(
    oldActivityData, oldPredData, newActivityData, newPredData)


WD = wordDoc.WordDoc
narrative = WD("NarrativeTemplate2.docx")

# Concatenate together the activity ID from list provided with any combination of activity.data 
def concatActInfo(ids:list,info:list, separator:str = " - "):
    concatList = []
    for actID in ids:
        concatIdInfo = actID
        for item in info:
            concatIdInfo += separator + comparison.allData.get(actID).get_data(item)
        concatList.append(concatIdInfo)
    return concatList

# Required Inputs
dataDate = datetime.datetime(2023, 1, 15)
previousDataDate = datetime.datetime(2022, 12, 18)
nextDataDate = dataDate + datetime.timedelta(days=30)


startedActivities = comparison.actBetween(previousDataDate, dataDate, False, 3)
completedActivities = comparison.actBetween(previousDataDate, dataDate, False, 4)
startedActivities = concatActInfo(startedActivities,["Activity Name"])
completedActivities = concatActInfo(completedActivities,["Activity Name"])
next30CritData = comparison.actBetween(
    dataDate, nextDataDate, True)
next30CritData = concatActInfo(next30CritData,["Activity Name"])

narrative.actSC(startedActivities, completedActivities)
narrative.next30CP(next30CritData)
narrative.activityMods(comparison.changed_names, comparison.added_activities, comparison.deleted_activities,
                       comparison.added_successors, comparison.deleted_successors, comparison.changed_durations)


narrative.saveFile("wordTest.docx")
