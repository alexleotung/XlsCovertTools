import os
import xlrd
import time
from xml.dom import minidom

from Xlxs2Xml import GlobalConst

# gernate lua scripts for xls file
class LuaScriptGenerator():
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

    def GetLuaFileName(self):
        fileName, suffix = GetFileName(self.importPath)
        names = fileName.split("_")
        luafile = "const_localization_%s.lua" % (names[1])
        parent = "{}/{}/Lua/".format(self.exportPath, names[1])
        if  not os.path.exists(parent):
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
        if not os.path.exists("{}/{}/".format(self.exportPath, names[1])):
            return

        self.exportPath = "{}/{}/Localization/".format(self.exportPath, names[1])
        if not os.path.exists(self.exportPath):
            os.makedirs(self.exportPath)

        if sheet.nrows <= 3:
            print("This file's first sheet is empty which {}".format(self.importPath))
            return

        for index in range(1, sheet.ncols):
            self.GenerateXml(sheet.col_values(0), sheet.col_values(index))

    def GenerateXml(self, keys, column):
        # set space before the row
        if "" == column[3]:
            return


        fileName, suffix = GetFileName(self.importPath)
        names = fileName.split("_")
        luaContext = GlobalConst.GlobalCost.LUA_LOCALIZATION_HEADER % (self.importPath, time.asctime(time.localtime()), names[1], names[1]);
        header = GlobalConst.GlobalCost.XML_HEADER_CONST % (self.importPath, time.asctime(time.localtime()))
        ipml = minidom.getDOMImplementation()
        doc = ipml.createDocument(None, None, None)
        root = doc.createElement("Dictionaries")
        # comment = doc.createComment(header)
        # root.appendChild(comment)

        dict = doc.createElement("Dictionary")
        dict.setAttribute("Language", column[1])
        root.appendChild(dict)

        for index in range(3, column.__len__()):
            child = GeneratorItem(doc, column[2], keys[index], column[index], )
            luaContext += "\t\t%s=\"%s\",\n" % (keys[index], keys[index])
            dict.appendChild(child)

        luaContext += "}"
        parent = "{}{}/".format(self.exportPath, column[1])
        if not os.path.exists(parent):
            os.makedirs(parent)

        path = "%s%s.xml" % (parent, column[1])
        f = open(path, "w", encoding="utf-8")

        root.writexml(f, addindent="", newl="\n")
        f.close()

        if self.luaFile == "":
            return

        f = open(self.luaFile, "w", encoding="utf-8")
        f.write(luaContext)
        f.close()

# wirte itme value
def GeneratorItem(doc, typeName, key, val, ):
    child = doc.createElement("String")
    child.setAttribute("Key", key)
    child.setAttribute("Value", val)
    return child

#
def FindFile(path, file):
    files= os.listdir(path)
    res = ""
    for fi in files:
        parent = "{}/{}".format(path, fi)
        if fi == file:
            return parent
        if os.path.isdir(parent):
            res= FindFile(parent, file)
            if res != "":
                break
    return res
#
def GetFileName(path):
    parent, file = os.path.split(path)
    shortName, suffix = os.path.splitext(file)
    return shortName, suffix
