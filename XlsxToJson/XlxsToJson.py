#!/usr/bin/python
# -*- coding:utf-8 -*-
import sys
from Xlxs2Json.Xlsx2JsonUtils import Xlsx2JsonUtils
from Xlxs2Json.EnumType import EnumConfigType, EnumImportType

help_string = """Generate lua script from the first sheet int the xls or xlsx files in the directory.
            sheet table format:           
            ------------------------------------------------------------------------------
            | description1 | description2 | description3 |  description4 |  description5 |
            ------------------------------------------------------------------------------
            |    id        |  parameter1  |  parameter2  |  parameter3   |  parameter4   |
            ------------------------------------------------------------------------------
            |   number     |    string    |    number    |    string     |     string    |          
            ------------------------------------------------------------------------------
            |     1        |    string1   |    100       |     string2   |     string3   |
            ------------------------------------------------------------------------------
        it has tow types of  parameters, witch are number and  string.             
            Parameters:
              -D: generate a dictionary,which keys are the first ro
              -L: generate a ArrayList by the files
              -f: the path of  config xls file
              -d: the directory of config xls files
              -o: the path of output script path 
              -h: help info
              --help: help info
            Example:
               XlsxToLua.py -f ./config.xls -o ./out.lua
               XlsxToLua.py -f ./configs/ -o ./out/"""

xlsx2JsonUtils = Xlsx2JsonUtils()

def show_help():
    print(help_string)


def split_parameters(arg_type, arg_str):
    if arg_type == "-D":
        xlsx2JsonUtils.SetGenerateTableType(EnumConfigType.Dictionary)
    elif arg_type == "-L":
        xlsx2JsonUtils.SetGenerateTableType(EnumConfigType.List)
    elif arg_type == "-d":
        xlsx2JsonUtils.SetImportType(EnumImportType.Directory)
        xlsx2JsonUtils.SetImportPath(arg_str)
    elif arg_type == "-f":
        xlsx2JsonUtils.SetImportType(EnumImportType.Single)
        xlsx2JsonUtils.SetImportPath(arg_str)
    elif arg_type == "-o":
        xlsx2JsonUtils.SetOutPath(arg_str)
    elif arg_type == "-h":
        show_help()
        sys.exit()
    elif arg_type == "--help":
        show_help()
        sys.exit()


def main(argus):
    if len(argus) == 1:
        show_help()
        return

    for index in range(1, len(argus)):
        split_parameters(argus[index - 1], argus[index])

    # split the last parameter.
    split_parameters(argus[len(argus) - 1], "")
    # check the parameters
    xlsx2JsonUtils.CheckParameters()
    # construct scripts  generator
    xlsx2JsonUtils.ConstructLuaScriptGenerators()
    # generate scripts
    xlsx2JsonUtils.GenerateLuaScriptByGenerators()


if __name__ == "__main__":
    main(sys.argv)
