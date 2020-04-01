# Включить поддержку Python и загрузить библиотеку DesignScript
# Подгрузка библиотек

import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager


# Разместите код под этой строкой

def U(elem):
	a = UnwrapElement(elem)
	return a
	
list = U(IN[0])
sought = IN[1]


for types in list:
	if types.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == sought:
		OUT = types
		
		