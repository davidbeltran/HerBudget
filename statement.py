from PyPDF2 import PdfReader
import regex as re

class Statement:
    def __init__(self, pdfDocument):
        self.pdfDocument = pdfDocument
        self.reDetail = ('(?:\n((?:0[1-9]|1[1,2])/(?:0[1-9]|[12][0-9]|3[01]))\s*(.+)'
            + ' ((?:-\d+\.\d{2})|(?:\d+\.\d{2})))')
        self.detailList, self.statementList = [], []

    def __preparePdf(self):
        reader = PdfReader(self.pdfDocument)
        pg = reader.pages[2]
        return pg.extract_text()
    
    def createExpenseList(self):
        pdfTxt = self.__preparePdf()
        return re.findall(self.reDetail, pdfTxt)
        """ i = 0
        for detail in self.detailList:
            detail = detail + (float(self.amountList[i]),)
            self.statementList.append(detail)
            i+=1 """

#print(len(reader.pages))
            
""" txtFile = open("NovDec.txt", "w")
txtFile.write(pdftxt)
txtFile.close()
 """