# XlsToLua

        Generate lua script from the first sheet int the xls or xlsx files in the directory.
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
               XlsxToLua.py -f ./configs/ -o ./out/