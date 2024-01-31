from statement import Statement
from database import Database

chase = Statement('NovDec.pdf')

expList = chase.createExpenseList()

db = Database(expList)
db.createExpenseTable()
db.fillExpenseTable()

for expense in expList:
    print(expense)

"""TODO
added items to database. now extract items and perform math
possibly composite associate database class to statement class instead of main
"""