import os
import sys

from Xlxs2Xml.EnumType import EnumConfigType, EnumImportType
from Xlxs2Xml.LuaScriptGenerator import LuaScriptGenerator


class Xlsx2XmlUtils:

    def __init__(self):
        self.__tableType = EnumConfigType.Dictionary
        self.__importType = EnumImportType.Single
        self.__importPath = ""
        self.__outPath = os.getcwd()
        self.__files = []

    ## set the type for generate table type
    def SetGenerateTableType(self, configType):
        self.__tableType = configType

    ##set type of import files
    def SetImportType(self, importType):
        self.__importType = importType

    ##set the import path
    def SetImportPath(self, importPath):
        self.__importPath = importPath

    ## set the export path
    def SetOutPath(self, outPath):
        self.__outPath = outPath

    # check the parameters
    def CheckParameters(self):
        # check the import files
        if self.__importType == EnumImportType.Single:
            if not os.path.isfile(self.__importPath):
                print("import file is not exist :{}".format(self.__importPath))
                sys.exit()

        elif self.__importType == EnumImportType.Directory:
            if not os.path.isdir(self.__importPath):
                print("import directory is not exist :{}".format(self.__importPath))
                sys.exit()
        # check the permission as read for the import directory
        if not os.access(self.__importPath, os.R_OK):
            print("check the permission :{}".format(self.__importPath))
            sys.exit()
        # check the export file.
        if self.__importType == EnumImportType.Directory:
            if self.__outPath.endswith(".Lua") or os.path.isfile(self.__outPath):
                print("can't combine multi xls files to a lua script file")
                sys.exit()

            if not os.path.exists(self.__outPath):
                os.makedirs(self.__outPath, 666)

    # construct lua script generator
    def ConstructLuaScriptGenerators(self):
        self.__files.clear()
        if self.__importType == EnumImportType.Single:
            # generate script by single file
            file = LuaScriptGenerator(self.__importPath, self.__outPath, self.__tableType)
            self.__files.append(file)
        elif self.__importType == EnumImportType.Directory:
            # generate script by directory
            files = os.listdir(self.__importPath)
            for name in files:
                if not (name.lower().endswith("xls") or name.lower().endswith("xlsx")):
                    continue

                file = LuaScriptGenerator("{}/{}".format(self.__importPath, name), self.__outPath, self.__tableType)
                self.__files.append(file)

    # generat script by file
    def GenerateLuaScriptByGenerators(self):
        while len(self.__files) > 0:
            file = self.__files.pop()
            print("Converting file {} to {}".format(file.importPath, file.exportPath))
            # generate table by file
            file.CheckFile()
            file.GetLuaFileName()
            file.GenerateScript()
            del file
