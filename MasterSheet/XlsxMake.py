from openpyxl import Workbook
from openpyxl import load_workbook
import re

class XlsxMake:
    def __init__(self):
        self.schools = {}
        self.getTemplate()
        self.openFile()
    #These are supporting functions for the constructor
    def getTemplate(self):
        res = input('Choose a template. Is this for pathways (P), clubs (C), or athletics (A)? ')
        res = res.lower()
        template = ''
        if(res == 'p'):
            template = 'pathways'
        elif(res == 'c'):
            template = 'clubs'
        else:
            template = 'athletics'
        self.template = load_workbook('templates/' + template + '.xlsx').active
        self.headers = load_workbook('headers/' + template + '.xlsx').active
    def openFile(self):
        self.fileName = input('Workbook to edit: ')
        try:
            self.wb = load_workbook('sheets/' + self.fileName + ".xlsx")
        except:
            print('Did not find workbook of that name. Creating new workbook.')
            self.wb = Workbook()
            self.wb.remove(self.wb['Sheet'])

    #Loads all user input into the schools dictionary
    def getUserInput(self):
        response = input("Sheet name: ")
        #An empty string response results in no save and no recursive request for another response
        if not response == '':
            self.schools[response] = (self.getUrl(), self.getProgramCount())
            self.getUserInput() #Request another response
    #Take user input for the number of programs at a school
    def getProgramCount(self):
        programCount = input("Maximum possible number of programs at this school: ")
        #Check to make sure the entered string can be easily evaluated as an integer
        for character in programCount:
            if not character.isdigit():
                print("Your response was not a number. ")
                return self.getProgramCount() #Request another response
        return int(programCount)
    #Get user input for URL
    def getUrl(self):
        url = input("URL to substitute in: ")
        #Check to make sure this is a valid URL
        if not url.startswith('https://'):
            print("Your response was not a valid URL. ")
            return self.getUrl() #Request another response because previous response was not acceptable
        return url
    #Replace all instances of the string "URL" with the desired URL
    def replaceUrls(self, ws, url):
        ret = Workbook()
        retws = ret.active #get default worksheet
        #iterate over all cells in the base worksheet, ws, and copy them with modification into the worksheet to be returned, retws
        for colIndex in range(ws.min_column, ws.max_column + 1):
            for rowIndex in range(ws.min_row, ws.max_row):
                if ws.cell(row = rowIndex, column = colIndex).value:
                    #copy with instances of "URL" replaced
                    try:
                        retws.cell(row = rowIndex, column = colIndex).value = ws.cell(row = rowIndex, column = colIndex).value.replace('URL', '"' + url + '"')
                    except:
                        pass
        return ret.active
    def addToNums(self, ws, numToAdd):
        ret = Workbook()
        retws = ret.active #get default worksheet
        #iterate over all cells in the base worksheet ws
        for colIndex in range(ws.min_column, ws.max_column + 1):
            for rowIndex in range(ws.min_row, ws.max_row):
                if ws.cell(row = rowIndex, column = colIndex).value:
                    #separate the value of the cell in ws at {{ or }}
                    parts = re.split(r'{{|}}', ws.cell(row = rowIndex, column = colIndex).value)
                    #iterate over the parts of the contents of the cell that need to be changed
                    for i in range(1, len(parts), 2):
                        parts[i] = str(int(parts[i]) + numToAdd)
                    retws.cell(row = rowIndex, column = colIndex).value = ''.join(parts)
        return ret.active
    def makeSheet(self):
        try:
            print(self.template.cell(row = 1, column = 2).value)
            templateNumberOfRows = int(self.template.cell(row = 1, column = 2).value)
        except:
            print("failed to get number of rows.")
            templateNumberOfRows = 20
        try:
            templateNumberOfCols = int(self.template.cell(row = 1, column = 4).value)
        except:
            templateNumberOfCols = 20
        try:
            headersNumberOfCols = int(self.headers.cell(row = 1, column = 4).value)
        except:
            headersNumberOfCols = 20
        for school in self.schools:
            if(school in self.wb.sheetnames): #it is necessary to replace a sheet that already exists
                if (len(self.wb.sheetnames) > 1):
                    self.wb.remove_sheet(self.wb[school])
                    self.wb.create_sheet(title = school)
                else: #it is not allowed to have only one sheet, so a temporary one must be created
                    self.wb.create_sheet(title = 'temporary sheet')
                    self.wb.remove(self.wb[school])
                    self.wb.create_sheet(title = school)
                    self.wb.remove(self.wb['temporary sheet'])
            else: self.wb.create_sheet(title = school)
            
            ws = self.wb[school]
            #Get a version of the template that is generally applicable for this school
            schoolTemplate = self.replaceUrls(self.template, self.schools[school][0])
            for programNumber in range(self.schools[school][1]):
                #Get a version of the template to be used for this specific program
                template = self.addToNums(schoolTemplate, programNumber * (templateNumberOfRows - self.template.min_row + 1))
                #Copy over the template
                for rowIndex in range(1, templateNumberOfRows + 1):
                    for colIndex in range(1, templateNumberOfCols + 1):
                        ws.cell(row = rowIndex + programNumber * (templateNumberOfRows - self.template.min_row + 1), column = colIndex).value = template.cell(row = rowIndex + 1, column = colIndex).value;
            ws.insert_rows(1)
            for column in range(1, headersNumberOfCols + 1):
                ws.cell(row = 1, column = column).value = self.headers.cell(row = 2, column = column).value
        return self.wb
    def save(self):
        try:
            self.makeSheet().save(filename = 'sheets/' + self.fileName + ".xlsx")
        except:
            print("Unable to save. Is the file open in another program?")
            self.save()