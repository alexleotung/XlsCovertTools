from enum import Enum

# generate lua table type
class EnumConfigType(Enum):
    Dictionary = 1
    List = 2

# generate file type
class EnumImportType(Enum):
    Single = 1
    Directory = 2

