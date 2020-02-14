from openpyxl import Workbook
from openpyxl import load_workbook
import re

class XlsxMake:
    def __init__(self):
        self.schools = {}
        self.pathwaysTemplate = None
        self.clubsTemplate = None
        self.athleticsTemplate = None
        self.pathwaysHeaders = None
        self.clubsHeaders = None
        self.athleticsHeaders = None
        self.wb = None
        self.pathwaysTemplateType = ''
        self.getTemplate()
        self.openFile()
    #These are supporting functions for the constructor
    def getTemplate(self):
        self.pathwaysTemplate = load_workbook('templates/pathways.xlsx').active
        self.pathwaysHeaders = load_workbook('headers/pathways.xlsx').active
        self.clubsTemplate = load_workbook('templates/clubs.xlsx').active
        self.clubsHeaders = load_workbook('headers/clubs.xlsx').active
        self.athleticsTemplate = load_workbook('templates/athletics.xlsx').active
        self.athleticsHeaders = load_workbook('headers/athletics.xlsx').active
    def openFile(self):
        self.fileName = input('Workbook to edit: ')
        try:
            self.wb = load_workbook('sheets/' + self.fileName + ".xlsx")
        except FileNotFoundError:
            print('Did not find workbook of that name. Creating new workbook.')
            self.wb = Workbook()
            self.wb.remove(self.wb['Sheet'])
            print('new workbook created')

    #Loads all user input into the schools dictionary
    def getUserInput(self):
        response = input("Sheet name: ")
        #An empty string response results in no save and no recursive request for another response
        if response != '':
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
    def replaceUrls(self, ws, url, minRow, maxRow, minColumn, maxColumn):
        ret = Workbook()
        retws = ret.active #get default worksheet
        #iterate over all cells in the base worksheet, ws, and copy them with modification into the worksheet to be returned, retws
        for colIndex in range(minColumn, maxColumn + 2):
            for rowIndex in range(minRow, maxRow + 2):
                if ws.cell(row = rowIndex, column = colIndex).value:
                    #copy with instances of "URL" replaced
                    try:
                        retws.cell(row = rowIndex, column = colIndex).value = ws.cell(row = rowIndex, column = colIndex).value.replace('URL', '"' + url + '"')
                    except:
                        pass
        return ret.active
    def addToNums(self, ws, numToAdd, minRow, maxRow, minCol, maxCol):
        ret = Workbook()
        retws = ret.active #get default worksheet
        #iterate over all cells in the base worksheet ws
        for colIndex in range(minCol, maxCol + 2):
            for rowIndex in range(minRow, maxRow + 2):
                if ws.cell(row = rowIndex, column = colIndex).value:
                    #separate the value of the cell in ws at {{ or }}
                    parts = re.split(r'{{|}}', ws.cell(row = rowIndex, column = colIndex).value)
                    #iterate over the parts of the contents of the cell that need to be changed
                    for i in range(1, len(parts), 2):
                        parts[i] = str(int(parts[i]) + numToAdd)
                    retws.cell(row = rowIndex, column = colIndex).value = ''.join(parts)
        return ret.active
    def createSheet(self, sheetName):
        if(sheetName in self.wb.sheetnames): #it is necessary to replace a sheet that already exists
            if (len(self.wb.sheetnames) > 1):
                self.wb.remove_sheet(self.wb[sheetName])
                self.wb.create_sheet(title = sheetName)
            else: #it is not allowed to have only one sheet, so a temporary one must be created
                self.wb.create_sheet(title = 'temporary sheet')
                self.wb.remove(self.wb[sheetName])
                self.wb.create_sheet(title = sheetName)
                self.wb.remove(self.wb['temporary sheet'])
        else: self.wb.create_sheet(title = sheetName)
        return self.wb[sheetName]
    @staticmethod
    def getNumberOfCols(sheet):
        return int(sheet.cell(row = 1, column = 4).value)
    @staticmethod
    def getNumberOfRows(sheet):
        return int(sheet.cell(row = 1, column = 2).value)
    def makeSheetPathways(self):
        templateNumberOfRows = XlsxMake.getNumberOfRows(self.pathwaysTemplate)
        templateNumberOfCols = XlsxMake.getNumberOfCols(self.pathwaysTemplate)
        headersNumberOfCols = XlsxMake.getNumberOfCols(self.pathwaysHeaders)
        for school in self.schools:
            ws = self.createSheet(school)
            #Get a version of the template that is generally applicable for this school
            schoolTemplate = self.replaceUrls(self.pathwaysTemplate, self.schools[school][0], 1, templateNumberOfRows, 1, templateNumberOfCols)
            for programNumber in range(self.schools[school][1]):
                #Get a version of the template to be used for this specific program
                template = self.addToNums(schoolTemplate, programNumber * (templateNumberOfRows - self.pathwaysTemplate.min_row + 1), 1, templateNumberOfRows, 1, templateNumberOfCols)
                #Copy over the template
                for rowIndex in range(1, templateNumberOfRows + 1):
                    for colIndex in range(1, templateNumberOfCols + 1):
                        ws.cell(row = rowIndex + programNumber * (templateNumberOfRows), column = colIndex).value = template.cell(row = rowIndex + 1, column = colIndex).value;
            ws.insert_rows(1)
            for column in range(1, headersNumberOfCols + 1):
                ws.cell(row = 1, column = column).value = self.pathwaysHeaders.cell(row = 2, column = column).value
        return self.wb
    def makeSheetClubs(self):
        for school in self.schools:
            ws = self.createSheet(school + ' Clubs')
            templateNumberOfRows = XlsxMake.getNumberOfRows(self.clubsTemplate)
            templateNumberOfCols = XlsxMake.getNumberOfCols(self.clubsTemplate)
            headerNumberOfCols = XlsxMake.getNumberOfCols(self.clubsHeaders)
            clubsTemplate = self.replaceUrls(self.clubsTemplate, self.schools[school][0], 1, templateNumberOfRows, 1, templateNumberOfCols)
            for row in range(1, templateNumberOfRows + 2):
                for column in range(1, templateNumberOfCols + 2):
                    ws.cell(row=row, column=column).value = clubsTemplate.cell(row=row+1, column=column).value
            ws.insert_rows(1)
            for column in range(1, headerNumberOfCols + 1):
                ws.cell(row=1, column=column).value = self.clubsHeaders.cell(row=2, column=column).value
    def makeSheetAthletics(self):
        for school in self.schools:
            ws = self.createSheet(school + ' Athletics')
            templateNumberOfRows = XlsxMake.getNumberOfRows(self.athleticsTemplate)
            templateNumberOfCols = XlsxMake.getNumberOfCols(self.athleticsTemplate)
            headerNumberOfCols = XlsxMake.getNumberOfCols(self.athleticsHeaders)
            athleticsTemplate = self.replaceUrls(self.athleticsTemplate, self.schools[school][0], 1, templateNumberOfRows, 1, templateNumberOfCols)
            for row in range(1, templateNumberOfRows + 2):
                for column in range(1, templateNumberOfCols + 2):
                    ws.cell(row=row, column=column).value = athleticsTemplate.cell(row=row+1, column=column).value
            ws.insert_rows(1)
            for column in range(1, headerNumberOfCols + 1):
                ws.cell(row=1, column=column).value = self.athleticsHeaders.cell(row=2, column=column).value
    def makeSheet(self):
        self.makeSheetPathways()
        self.makeSheetClubs()
        self.makeSheetAthletics()
        return self.wb
    def save(self):
        try:
            self.makeSheet().save(filename = 'sheets/' + self.fileName + ".xlsx")
        except IOError:
            input("Unable to save. Is the file open in another program?")
            self.save()