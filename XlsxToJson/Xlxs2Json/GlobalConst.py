class GlobalCost:
    TS_LOCALIZATION_HEADER = """/**create ts by localization generator\n 
 *module:%s\n 
 *time:%s\n
 **/\n
 export default class %sLanguageConfig{\n
        public static GameName: string = '%s'\n;
 """

    TS_KEY_LINES_CONTEXT = """      public static %s: string = '%s';\n"""

    TS_END_CONST = """ };"""
