import datetime
# Activity class definition


class Activity:
    def __init__(self, worksheetTableRow, worksheetHeader):
        self.id = worksheetTableRow[worksheetHeader['Activity ID']]
        self.name = worksheetTableRow[worksheetHeader['Activity Name']]
        self.data = dict(zip(worksheetHeader, worksheetTableRow))
        self.successors = set()

    # Return a value from the data dictionary
    def get_data(self, dataKey):
        value = self.data[dataKey]
        return value


class Comparison:
    # Returns dictionary in form {Activity ID# : Activity Object} for all activities in passed schedule
    def TableToDict(self, activityData, predData):
        activityDict = {}
        for row in activityData.data:
            id = row[activityData.header['Activity ID']]
            activityDict.update({id: Activity(row, activityData.header)})

        for row in predData.data:
            key = row[predData.header['Predecessor']]
            successor = row[predData.header['Successor']]
            activityDict[key].successors.add(successor)
        return activityDict
    
    # Returns a comparison of the passed data title between new and old schedules in form [[Header] , [body]]
    def comparisonTable(self,dataTitle:str):
        # Build comparison header
        compTableHeader = ["Activity ID",
                            "Old " + dataTitle, "New " + dataTitle]

        # Build comparison body
        compTableBody = []
        for activityID in self.allIDS:
            oldActivityObj = self.oldData.get(activityID)
            newActivityObj = self.newData.get(activityID)
            if oldActivityObj == None or newActivityObj == None:
                continue
            oldDataValue = oldActivityObj.data.get(dataTitle)
            newDataValue = newActivityObj.data.get(dataTitle)
            if oldDataValue != newDataValue:
                compTableBody.append(
                    [activityID, oldDataValue, newDataValue])
        return [compTableHeader,compTableBody]
    
    def successorTables(self):
        # Build comparison header
        succTableHeader = ["Predecessor Activity ID","Predecessor Activity Name", "Successor Activity ID"]

        # Build comparison body
        addedSuccessorBody = []
        deletedSuccessorBody = []
        for activityID in self.allIDS:
            oldActivityObj = self.oldData.get(activityID)
            newActivityObj = self.newData.get(activityID)
            
            if oldActivityObj == None:
                oldDataValue = set()
            else:
                oldDataValue = oldActivityObj.successors
            
            if newActivityObj == None:
                newDataValue = set()
            else:
                newDataValue = newActivityObj.successors
            
            addedSuccessors = newDataValue - oldDataValue
            deletedSuccessors = oldDataValue - newDataValue


            if addedSuccessors != {}:
                for succID in addedSuccessors:
                    succName = self.newData.get(succID).name
                    addedSuccessorBody.append(    
                        [activityID, newActivityObj.name, succID, succName])
            if deletedSuccessors != {}:
                for succID in deletedSuccessors:
                    succName = self.oldData.get(succID).name
                    deletedSuccessorBody.append(
                        [activityID, oldActivityObj.name, succID,succName])
        return [succTableHeader,addedSuccessorBody] , [succTableHeader,deletedSuccessorBody]

    def __init__(self, oldActivityTable, oldPredTable, newActivityTable, newPredTable):
        self.oldData = self.TableToDict(oldActivityTable, oldPredTable)
        self.newData = self.TableToDict(newActivityTable, newPredTable)
        self.allIDS = set([key for key in self.oldData] +
                          [key for key in self.newData])
        self.changed_names = self.comparisonTable("Activity Name")
        self.changed_durations = self.comparisonTable("Original Duration")
        self.added_Successors, self.deleted_Successors = self.successorTables()



    # Required Comparisons
    # Changed Activity Names
    # Changed Duration
    # Added/Deleted Activities
    # Added/ Deleted Relationships
