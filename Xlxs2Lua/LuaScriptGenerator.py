import os
import xlrd
import time

from Xlxs2Lua import GlobalConst
from Xlxs2Lua.EnumType import EnumConfigType

# gernate lua scripts for xls file
class LuaScriptGenerator():
    def __init__(self, file, out, talbeType):
        self.importPath = file
        self.exportPath = out
        self.status = False
        self.tableType = talbeType
        self.importFile = {}
        self.exportFile = {}
        self.titles = []
        self.types = []

    def __del__(self):
        del self.importFile
        del self.exportFile

    def CheckFile(self):
        self.status = True
        if not os.path.isfile(self.importPath):
            self.status = False
        if not os.access(self.exportPath, os.W_OK):
            self.status = False
        if not os.access(self.importPath, os.R_OK):
            self.status = False

    # generate lua scripts
    def GenerateScript(self):
        if not self.status:
            print("cant generate script by file {}".format(self.importPath))
            return

        self.importFile = xlrd.open_workbook(self.importPath)
        sheet = self.importFile.sheet_by_index(0)
        if os.path.isdir(self.exportPath):
            fileName, suffix = GetFileName(self.importPath)
            self.exportPath = "{}/{}.lua".format(self.exportPath, fileName)

        if sheet.nrows <= 3:
            print("This file's first sheet is empty which {}".format(self.importPath))
            return

        header = GlobalConst.GlobalCost.LUA_HEADER_CONST % ("%s%s" % (fileName, suffix), time.asctime(time.localtime()))

        self.exportFile = open(self.exportPath, "w", encoding="utf-8")
        self.exportFile.write(header)

        # get Title and type
        self.titles = sheet.row_values(1)
        self.types = sheet.row_values(2)

        for index in range(3, sheet.nrows-1):
            self.GenerateItems(sheet.row_values(index))

        self.exportFile.write(GlobalConst.GlobalCost.LUA_END_CONST)
        self.exportFile.close()

    def GenerateItems(self, row):
        #set space before the row
        self.exportFile.write("\t")

        # write the key of this row
        if self.tableType == EnumConfigType.Dictionary:
            key = "%s="
            if self.types[0].lower() == "number":
                key = "[%d]="

            key = key % (row[0])
            self.exportFile.write(key)

        # write the items of the row
        self.exportFile.write("{")
        for index in range(0, len(row)):
            self.WriteItemValue(row, index)

        self.exportFile.write("},\n")


    # wirte itme value
    def WriteItemValue(self, row, index):
        if row[index] == "":
            return

        key = "{}={},"
        if self.types[index].lower() == "string":
            key = "{}=\"{}\""

        key = key.format(self.titles[index], row[index])
        self.exportFile.write(key)

        #write separator between tow values
        if index < len(row)-1 and row[index+1] != "":
            self.exportFile.write(",")
#
def GetFileName(path):
    parent, file = os.path.split(path)
    shortName, suffix = os.path.splitext(file)
    return shortName, suffix
