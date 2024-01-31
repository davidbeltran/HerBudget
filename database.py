import sqlite3

class Database:

    def __init__(self, statementList):
        self.statementList = statementList
        self.name = 'expenses.db'

    def createExpenseTable(self):
        try:
            conn = sqlite3.connect(self.name)
            c = conn.cursor()
            c.execute("""CREATE TABLE transactions (
                    Date text,
                    Details text,
                    Amount real
                    )
                    """)
            conn.commit()
            conn.close()
        except sqlite3.OperationalError:
            pass

    def fillExpenseTable(self):
        try:
            conn = sqlite3.connect(self.name)
            c = conn.cursor()
            rows = c.execute("""SELECT * FROM transactions""").fetchall()
            if not rows:
                c.executemany("INSERT INTO transactions VALUES (?, ?, ?)", self.statementList)
                conn.commit()
            else:
                pass
        except sqlite3.OperationalError:
            pass
