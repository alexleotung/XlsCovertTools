import os
import xlrd
import time
import json

from Xlxs2Json import GlobalConst


# gernate lua scripts for xls file
class JsonScriptGenerator():
    def __init__(self, file, out, talbeType):
        self.importPath = file
        self.exportPath = out
        self.luaFile = "";
        self.status = False
        self.tableType = talbeType
        self.importFile = {}
        self.titles = []
        self.types = []

    def __del__(self):
        del self.importFile

    def CheckFile(self):
        self.status = True
        if not os.path.isfile(self.importPath):
            self.status = False
        if not os.access(self.exportPath, os.W_OK):
            self.status = False
        if not os.access(self.importPath, os.R_OK):
            self.status = False

    def GetScriptFileName(self):
        fileName, suffix = GetFileName(self.importPath)
        names = fileName.split("_")
        luafile = "%sLanguageConfig.ts" % (names[1].capitalize())
        parent = os.path.dirname(self.exportPath)
        parent =  os.path.dirname(parent)
        parent = "{}/Game_{}/Script/Const/".format(parent, names[1].capitalize())
        if not os.path.exists(parent):
            self.luaFile = ""
            return

        self.luaFile = FindFile(parent, luafile)

    # generate lua scripts
    def GenerateScript(self):
        if not self.status:
            print("cant generate script by file {}".format(self.importPath))
            return

        self.importFile = xlrd.open_workbook(self.importPath)
        sheet = self.importFile.sheet_by_index(0)
        fileName, suffix = GetFileName(self.importPath)
        names = fileName.split("_")
        if not os.path.exists( self.exportPath ):
            return

        self.exportPath = "{}/{}/lang/".format(self.exportPath, names[1])
        if not os.path.exists(self.exportPath):
            os.makedirs(self.exportPath)

        if sheet.nrows <= 3:
            print("This file's first sheet is empty which {}".format(self.importPath))
            return

        for index in range(1, sheet.ncols):
            self.GenerateJson(sheet.col_values(0), sheet.col_values(index))

    def GenerateJson(self, keys, column):
        # set space before the row
        if "" == column[3]:
            return

        fileName, suffix = GetFileName(self.importPath)
        names = fileName.split("_")
        tsContext = GlobalConst.GlobalCost.TS_LOCALIZATION_HEADER % (
            self.importPath, time.asctime(time.localtime()), names[1].capitalize(), names[1])
        dict = {}

        for index in range(3, column.__len__()):
            GeneratorItem(dict, column[2], keys[index], column[index], )
            tsContext = tsContext + GlobalConst.GlobalCost.TS_KEY_LINES_CONTEXT % (keys[index], keys[index])

        tsContext = tsContext + GlobalConst.GlobalCost.TS_END_CONST

        parent = "{}{}/".format(self.exportPath, column[1])
        if not os.path.exists(parent):
            os.makedirs(parent)

        path = "%slang.json" % (parent)
        f = open(path, "w", encoding="utf-8")
        json.dump(dict, f, ensure_ascii=False, sort_keys=True, indent=4)
        f.close()

        if self.luaFile == "":
            return

        f = open(self.luaFile, "w", encoding="utf-8")
        f.write(tsContext)
        f.close()


# wirte itme value
def GeneratorItem(doc, typeName, key, val, ):
    doc[key] = val

#
def FindFile(path, file):
    files = os.listdir(path)
    res = ""
    for fi in files:
        parent = "{}/{}".format(path, fi)
        if fi == file:
            return parent
        if os.path.isdir(parent):
            res = FindFile(parent, file)
            if res != "":
                break
    return res


#
def GetFileName(path):
    parent, file = os.path.split(path)
    shortName, suffix = os.path.splitext(file)
    return shortName, suffix
