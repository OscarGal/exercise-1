
import os
from win32com.client import Dispatch
from datetime import datetime
import shutil

class Excel():

    def __init__(self, baseFolder):
        self.excelApplication = Dispatch("Excel.Application")
        self.pathMasterPath = self.getMasterPath()
        self.baseFolder = baseFolder

    def getExtensionFile(self, file):
        if file.find(".") == -1:
            return ""
        return file.split(".")[-1]

    def getCurrentTime(self):
        now = datetime.now()
        return now.strftime("%y_%m%d_%H%M%S")

    def processedWorkbooksSheets(self, listExcelFiles):
        if len(listExcelFiles) == 0:
            return
        
        self.createMasterWorkbook()
        
        listProcessedFiles = []
        currentTime = self.getCurrentTime()
        masterExcelFile = self.excelApplication.Workbooks.Open(Filename=self.pathMasterPath)
        totalFileExcel = 1
        for file in listExcelFiles:
            if not os.path.exists(file):
                continue
            otherExcelFile = self.excelApplication.Workbooks.Open(Filename=file)
            try:
                countOtherExcelFile = otherExcelFile.Sheets.count
                for numberSheet in range(countOtherExcelFile):
                    numberSheet = numberSheet + 1
                    otherSheet = otherExcelFile.Worksheets(numberSheet)
                    otherSheet.Copy(Before=masterExcelFile.Worksheets(1))
                    nameNewSheet = "{}_{}_{}".format(currentTime, totalFileExcel, numberSheet)
                    nameNewFile = "{}_{}".format(currentTime, totalFileExcel)
                    masterExcelFile.Worksheets(1).Name = nameNewSheet
                listProcessedFiles.append({
                    'path': file,
                    'name': nameNewFile,
                })
            except Exception as error:
                print(error)
            finally:
                otherExcelFile.Close()

            totalFileExcel = totalFileExcel + 1

        masterExcelFile.Close(SaveChanges=True)
        self.excelApplication.Quit()
        self.moveFileProcessed(listProcessedFiles)
    
    def createMasterWorkbook(self):
        if not os.path.exists(self.pathMasterPath):
            fileExcel = self.excelApplication.Workbooks.Add()
            fileExcel.SaveAs(self.pathMasterPath)
            self.excelApplication.Quit()

    def getMasterPath(self):
        return os.path.join(os.getcwd(), 'master_workbook.xlsx')

    def moveFileProcessed(self, listFiles):
        pathProcessed = os.path.join(os.getcwd(), 'Processed')
        if not os.path.isdir(pathProcessed):
            os.mkdir(pathProcessed)

        for file in listFiles:
            pathOldFile = file['path']
            if not os.path.isfile(pathOldFile):
                continue
            fileExtension = self.getExtensionFile(pathOldFile)
            nameNewFile = file['name'] + "." + fileExtension
            pathNewFile = os.path.join(pathProcessed, nameNewFile)
            shutil.move(pathOldFile, pathNewFile)

    def moveFileNotApplicable(self, listFiles):
        pathNotApplicable = os.path.join(os.getcwd(), 'Not applicable')
        if not os.path.isdir(pathNotApplicable):
            os.mkdir(pathNotApplicable)

        for pathFile in listFiles:
            if not os.path.isfile(pathFile):
                continue
            nameFile = os.path.basename(pathFile)
            newPathFile = os.path.join(pathNotApplicable, nameFile)
            shutil.move(pathFile, newPathFile)

    def readFiles(self):
        if not os.path.isdir(self.baseFolder):
            print("Path not exists:", self.baseFolder)
            return
        print("Base path", self.baseFolder)
        listExcelType = []
        listOtherType = []
        listExtensionValid = ['xlsx', 'xls']
        for file in os.listdir(self.baseFolder):
            joinPath = os.path.join(self.baseFolder, file)
            fileExtension = self.getExtensionFile(joinPath)
            if fileExtension in listExtensionValid:
                listExcelType.append(joinPath)
            else:
                listOtherType.append(joinPath)

        self.processedWorkbooksSheets(listExcelType)
        self.moveFileNotApplicable(listOtherType)
    
