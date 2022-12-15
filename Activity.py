import datetime
# Activity class definition


class Activity:
    def __init__(self, worksheetTableRow, worksheetHeader):
        self.id = worksheetTableRow[worksheetHeader['Activity ID']]
        self.name = worksheetTableRow[worksheetHeader['Activity Name']]
        self.data = dict(zip(worksheetHeader, worksheetTableRow))
        self.successors = set()


    # Find Successors
    # Pass worksheet converted into values

    def get_successors(self, successorWS):
        for row in successorWS:
            if row[0] != self.id:
                continue
            else:
                self.successors.add(row[1])
