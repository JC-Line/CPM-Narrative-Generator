import datetime
# Activity class definition


class Activity:
    def __init__(self, worksheetTableRow, worksheetHeader):
        self.id = worksheetTableRow[worksheetHeader['Activity ID']]
        self.name = worksheetTableRow[worksheetHeader['Activity Name']]
        self.data = dict(zip(worksheetHeader, worksheetTableRow))
        self.successors = set()
        self.critical = self.isCrit()
        self.start = self.get_data('Start')
        self.start_actual = self.get_data('Actual Start')
        self.finish = self.get_data('Finish')
        self.finish_actual = self.get_data('Actual Finish')

    # Return a value from the data dictionary
    def get_data(self, dataKey):
        value = self.data.get(dataKey,None)
        return value
    
    # Checks to see if activity is critical
    def isCrit(self):
        critStatus = 'Unknown'
        if self.get_data('Critical') == 'Y':
            critStatus = 'Critical'
        if self.get_data('Critical') == 'N':
            critStatus = 'Normal'
        return critStatus   




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
        compTableHeader = ["Activity ID","Activity Name",
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
                    [activityID,newActivityObj.name, oldDataValue, newDataValue])
        return [compTableHeader,compTableBody]
    
    def successorTables(self):
        # Build  header
        succTableHeader = ["Predecessor Activity ID","Predecessor Activity Name", "Successor Activity ID" , "Successor Activity Name"]

        # Create body
        addedSuccessorBody = []
        deletedSuccessorBody = []
        for activityID in self.allIDS:
            # Create two sets of activity ID's old schedule and new schedule
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
            
            # Find differences in old/new schedule successors
            addedSuccessors = newDataValue - oldDataValue
            deletedSuccessors = oldDataValue - newDataValue

            # Fetch table data for added/deleted successors
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
    
    # Returns added and deleted activity tables
    def activityTables(self):
        # Header
        actTableHeader = ['Activity ID', 'Activity Name']
        # Body
        oldIDS = set((key for key in self.oldData))
        newIDS = set((key for key in self.newData))

        addedIDS = self.allIDS - oldIDS
        deletedIDS = self.allIDS - newIDS 
        
        addedIdBody = self.idsToBody(addedIDS)
        deletedIdBody = self.idsToBody(deletedIDS)
        return [actTableHeader,addedIdBody] , [actTableHeader,deletedIdBody]
    
    # Concatenate together activity ID's with any combination of activity.data
    def concatActInfo(self,ids: list, info: list, separator: str = " - "):
        concatList = []
        for actID in ids:
            concatIdInfo = actID
            for item in info:
                concatIdInfo += separator + \
                    self.allData.get(actID).get_data(item)
            concatList.append(concatIdInfo)
        return concatList


    # Returns a table of activity data when given a set of activity ID's 
    # Starting with (Activity ID, Activity Name) + Data fields from passed list
    def idsToBody(self,ids,dataFields=()):
        body = []
        
        for id in ids:
            idName = self.allData.get(id).name    
            # Gather data values from dataFields if provided
            fieldData = ()
            if dataFields != ():
                
                for count,field in enumerate(dataFields):
                    currentFieldData = self.allData.get(id).data.get(field[count])
                    fieldData = fieldData + currentFieldData

            body.append(list((id,idName) + fieldData))
        return body
            

    # Finds activities in new schedule between two datetimes
    # Sort perameter values 
    # 1:activity.start
    # 2:activity.finish
    # 3:activity.start_actual
    # 4:activity.finish_actual
    def actBetween(self,startDate,endDate,criticalOnly = False,sortPerameter = 1):
        
        activities = []
        for activity in self.newData.values():
            perameters = {1:activity.start,2:activity.finish,3:activity.start_actual,4:activity.finish_actual}
            sortDate = perameters[sortPerameter]
            
            critStatus = False
            if activity.critical == "Critical":
                critStatus = True

            if sortDate == None:
                continue

            if sortDate <= startDate:
                continue
            
            if sortDate >= endDate:
                continue
                
            if criticalOnly == True:
                if critStatus == False:
                    continue
            activities.append(activity.id)
        return activities
     









    def __init__(self, oldActivityTable, oldPredTable, newActivityTable, newPredTable):
        self.oldData = self.TableToDict(oldActivityTable, oldPredTable)
        self.newData = self.TableToDict(newActivityTable, newPredTable)
        self.allData = self.oldData.copy()
        self.allData.update(self.newData)
        self.allIDS = set([key for key in self.allData])
        self.changed_names = self.comparisonTable("Activity Name")
        self.changed_durations = self.comparisonTable("Original Duration")
        self.added_successors, self.deleted_successors = self.successorTables()
        self.added_activities, self.deleted_activities = self.activityTables()
