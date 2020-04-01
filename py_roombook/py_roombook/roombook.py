# Включить поддержку Python и загрузить библиотеку DesignScript
# Подгрузка библиотек

import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import*
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ObjectType , ISelectionFilter

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

import math

import sys
sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\Lib')
import collections
#import operator
from collections import defaultdict
#from operator import itemgetter, attrgetter
from System.Collections.Generic import*

# Получение текущего проекта
doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
uiapp=DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application

####################################  Выборка помещений  ################################################

class selectionfilter(ISelectionFilter):
	def __init__(self,categories):
		self.categories = categories
	def AllowElement(self,element):
		if element.Category.Name in [c.Name for c in self.categories] :
			return True
		else:
			return False
	def AllowReference(reference,point):
		return False

if isinstance(IN[0],list):
	categories = UnwrapElement(IN[0])
else:
	categories = [UnwrapElement(IN[0])]
	
selfilt = selectionfilter(categories)
catnames = ', '.join([c.Name for c in categories])

if True:
	sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element,selfilt,'Select %s' %(catnames))
	selelem = [doc.GetElement(s.ElementId) for s in sel]
	rooms = selelem
	
def room_sort_key(room):
    p1 = room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString()
    p2 = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
    return p1, p2

rooms = sorted(rooms, key = room_sort_key)


#lst = sorted(rooms, key=attrgetter(room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()))

###################################### Всякие шляпные функции ####################################################

def U(elem):
	a = UnwrapElement(elem)
	return a
	
calculator = SpatialElementGeometryCalculator(doc)
opt = SpatialElementBoundaryOptions()

def SpatialElement(room):
	faces = calculator.CalculateSpatialElementGeometry(room).GetGeometry().Faces
	spatial = list(calculator.CalculateSpatialElementGeometry(room).GetBoundaryFaceInfo(face) for face in faces)
	flattened_list = [y for x in spatial for y in x]
	return flattened_list

def get_type(elem):
	return doc.GetElement(elem.GetTypeId())

##############################  Ввод переменых  ########################################

type_elem = U(IN[1])
type_text = U(IN[2])
filt_mat = IN[3]#.split(", ")

sp_n = int(IN[4])
sp_pot = int(IN[5])
sp_plo_pot = int(IN[6])
sp_sten = int(IN[7])
sp_plo_sten = int(IN[8])
sp_kolon = int(IN[9])
sp_plo_kolon = int(IN[10])
sp_pol = int(IN[11])
sp_plo_pol = int(IN[12])
sp_plin = int(IN[13])
sp_plo_plin = int(IN[14])
sp_prim = int(IN[15])

##############################  Получение искомого вида  ########################################

Views=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()

if any("Форма 1 Ведомость отделки помещений" in x.ViewName for x in Views):
	view = filter(lambda x: "Форма 1 Ведомость отделки помещений" in x.ViewName, Views)[0]
else:
	t = Transaction(doc, 'Создание чертежного вида')
	t.Start()
	view = ViewDrafting.Create(doc,filter(lambda x: "Чертежный вид" in x.FamilyName, FilteredElementCollector(doc).OfClass(ViewFamilyType))[0].Id)
	view.LookupParameter('Имя вида').Set("Форма 1 Ведомость отделки помещений")
	view.LookupParameter('Масштаб вида').Set(1)
	
	t.Commit()

#######################  Списки для  подсчета отделки  ############################################

walls = []
floors = []
ceiling = []
columns = []
walls_area = []
floors_area = []
ceiling_area = []
columns_area = []

##################################  Заполнение словаря отделкой  #####################################

for ind, room in enumerate(rooms):
    room_dict_wall = defaultdict(float)
    room_dict_floor = defaultdict(float)
    room_dict_ceiling = defaultdict(float)
    room_dict_plinth = defaultdict(float)
    room_dict_columns = defaultdict(float)
    used = set()
    a = room.GetBoundarySegments(opt)
    flattened_list = [y for x in room.GetBoundarySegments(opt) for y in x]
    for i in flattened_list:
    	if hasattr(doc.GetElement(i.ElementId), 'Category'):
    		if doc.GetElement(i.ElementId).Category.Name is "Стены" or "Несущие колонны":
    			long = round(i.GetCurve().Length*0.3048, 2)
    			#if room.LookupParameter("APMV_Тип плинтуса").AsString() != None:
    				#room_dict_plinth[room.LookupParameter("APMV_Тип плинтуса").AsString()] += long
    #param_plinth = "\r\n".join(get_string_m(room_dict_plinth))

    for sub in SpatialElement(room):
    	elem_room = doc.GetElement(sub.SpatialBoundaryElement.HostElementId)
    	if hasattr(elem_room, 'Category'):
    		Cat_Elem =  elem_room.Category.Name
    		Elem_id = elem_room.Id
    		if Cat_Elem == "Стены" and Elem_id not in used and ("Отделка ЖБ" != elem_room.LookupParameter("Комментарии").AsString()):
    			Mats_id = elem_room.GetMaterialIds(False)
    			for id in Mats_id:
    				if filt_mat == doc.GetElement(id).LookupParameter("ADSK_Группирование").AsString():
    					mat_area = elem_room.GetMaterialArea(id, False)
    					mat_name = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString()
    					if elem_room.GetType() == FamilyInstance:
    						room_dict_wall[mat_name] += (round(mat_area*0.09290304/2, 2))
    					else:
    						room_dict_wall[mat_name] += (round(mat_area*0.09290304, 2))
    		elif Cat_Elem == "Перекрытия" and Elem_id not in used:
    			Mats_id = elem_room.GetMaterialIds(False)
    			for id in Mats_id:
    				if filt_mat == doc.GetElement(id).LookupParameter("ADSK_Группирование").AsString():
    					mat_area = elem_room.GetMaterialArea(id, False)
    					mat_name = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString()
    					if elem_room.GetType() == FamilyInstance:
    						room_dict_floor[mat_name] += (round(mat_area*0.09290304/2, 2))
    					else:
    						room_dict_floor[mat_name] += (round(mat_area*0.09290304, 2))
    		elif Cat_Elem == "Потолки" and Elem_id not in used:
    			Mats_id = elem_room.GetMaterialIds(False)
    			for id in Mats_id:
    				if filt_mat == doc.GetElement(id).LookupParameter("ADSK_Группирование").AsString():
    					mat_area = elem_room.GetMaterialArea(id, False)
    					mat_name = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString()
    					if elem_room.GetType() == FamilyInstance:
    						room_dict_ceiling[mat_name] += (round(mat_area*0.09290304/2, 2))
    					else:
    						room_dict_ceiling[mat_name] += (round(mat_area*0.09290304, 2))
    		elif ("Несущие колонны"  or  "Стены" in Cat_Elem) and (Elem_id not in used) and ("Отделка ЖБ" == elem_room.LookupParameter("Комментарии").AsString()):
    			Mats_id = elem_room.GetMaterialIds(False)
    			for id in Mats_id:
    				if filt_mat == doc.GetElement(id).LookupParameter("ADSK_Группирование").AsString():
    					mat_area = elem_room.GetMaterialArea(id, False)
    					mat_name = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString()
    					if elem_room.GetType() == FamilyInstance:
    						room_dict_columns[mat_name] += (round(mat_area*0.09290304/2, 2))
    					else:
    						room_dict_columns[mat_name] += (round(mat_area*0.09290304, 2))
    		used.add(Elem_id)
		
    lst_mat_name_wall = room_dict_wall.keys()
    lst_mat_name_floor = room_dict_floor.keys()
    lst_mat_name_ceileng = room_dict_ceiling.keys()
    lst_mat_name_columns = room_dict_columns.keys()

    lst_mat_area_wall = room_dict_wall.values()
    lst_mat_area_floor = room_dict_floor.values()
    lst_mat_area_ceileng = room_dict_ceiling.values()
    lst_mat_area_columns = room_dict_columns.values()

    walls.append(lst_mat_name_wall)
    walls_area.append(lst_mat_area_wall)
    floors.append(lst_mat_name_floor)
    columns.append(lst_mat_name_columns)

    floors_area.append(lst_mat_area_floor)
    ceiling.append(lst_mat_name_ceileng)
    ceiling_area.append(lst_mat_area_ceileng)
    columns_area.append(lst_mat_area_columns)

########################  Разработка и растановка значение отделки на виде ###################################
lpy_w = lpy_f = lpy_c = apy_w = apy_f = apy_c = lpy_cn = apy_cn = -0.003
point_Y = 0

#Создание транзакции
t = Transaction(doc, 'Разместить строки')

for number, room in enumerate(rooms):
	number_room = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
	name_room = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
	elem_id_for_group = []
	#room_for_group = []
	for ind, walls_room in enumerate(walls[number]):
		#Открытие транзакции
		t.Start("new")
		text_wall = TextNote.Create(doc, view.Id, XYZ((sp_n + sp_pot + sp_plo_pot + 2)/304.8,lpy_w,0), (sp_sten-3)/304.8, walls_room + ";", type_text.Id)
		area_wall = TextNote.Create(doc, view.Id, XYZ((sp_n + sp_pot + sp_plo_pot + sp_sten + 2)/304.8,apy_w,0), (sp_plo_sten-3)/304.8, str(walls_area[number][ind]) + ";", type_text.Id)		
		#Закрытие транзакции
		t.Commit()	

		elem_id_for_group.append(text_wall.Id)
		elem_id_for_group.append(area_wall.Id)
		lpy_w -= text_wall.Height + 0.008				
		apy_w -= text_wall.Height + 0.008
		

	for ind, floors_room in enumerate(floors[number]):
		#Открытие транзакции
		t.Start("new")
		text_floor = TextNote.Create(doc, view.Id, XYZ((sp_n  + sp_pot + sp_plo_pot + sp_sten + sp_plo_sten + sp_kolon + sp_plo_kolon + 2)/304.8,lpy_f,0), (sp_pol-3)/304.8, floors_room + ";", type_text.Id)
		area_floor = TextNote.Create(doc, view.Id, XYZ((sp_n  + sp_pot + sp_plo_pot + sp_sten + sp_plo_sten + sp_kolon + sp_plo_kolon + sp_pol +2)/304.8,apy_f,0), (sp_plo_pol-3)/304.8, str(floors_area[number][ind]) + ";", type_text.Id)
		#Закрытие транзакции
		t.Commit()

		elem_id_for_group.append(text_floor.Id)
		elem_id_for_group.append(area_floor.Id)
		lpy_f -= text_floor.Height + 0.008
		apy_f -= text_floor.Height + 0.008
		
	for ind, ceiling_room in enumerate(ceiling[number]):
		t.Start("new")
		text_ceiling = TextNote.Create(doc, view.Id, XYZ((sp_n + 2)/304.8, lpy_c, 0), (sp_pot- 3)/304.8, ceiling_room + ";", type_text.Id)
		area_ceiling = TextNote.Create(doc, view.Id, XYZ((sp_n  + sp_pot + 2)/304.8,apy_c,0), (sp_plo_pot - 3)/304.8, str(ceiling_area[number][ind]) + ";", type_text.Id)
		t.Commit()
		
		elem_id_for_group.append(text_ceiling.Id)
		elem_id_for_group.append(area_ceiling.Id)		
		lpy_c -= text_ceiling.Height + 0.008
		apy_c -= text_ceiling.Height + 0.008

	for ind, columns_room in enumerate(columns[number]):
		t.Start("new")
		text_columns = TextNote.Create(doc, view.Id, XYZ((sp_n  + sp_pot + sp_plo_pot + sp_sten + sp_plo_sten + 2)/304.8 ,lpy_cn,0), (sp_kolon -3)/304.8, columns_room + ";", type_text.Id)
		area_columns = TextNote.Create(doc, view.Id, XYZ((sp_n  + sp_pot + sp_plo_pot + sp_sten + sp_plo_sten + sp_kolon + 2)/304.8,apy_cn,0), (sp_plo_kolon - 3)/304.8, str(columns_area[number][ind]) + ";", type_text.Id)
		t.Commit()
		
		elem_id_for_group.append(text_columns.Id)
		elem_id_for_group.append(area_columns.Id)		
		lpy_cn -= text_columns.Height + 0.008
		apy_cn -= text_columns.Height + 0.008
		
	
	lpy_w = lpy_f = lpy_c = apy_w = apy_f = apy_c = apy_cn = lpy_cn = min(lpy_w, lpy_f, lpy_c,apy_f,apy_w,apy_c, lpy_cn, apy_cn)

	#Открытие транзакции
	t.Start()
	str_room = doc.Create.NewFamilyInstance(XYZ(0,point_Y,0),type_elem,view)
	str_room.LookupParameter('Высота строки').Set(-1*lpy_w + point_Y)	
	str_room.LookupParameter('Наименование').Set(number_room + ", " + name_room)
	str_room.LookupParameter("Нижнее подчеркивание").Set(1)
	elem_id_for_group.append(str_room.Id)
	str_room.LookupParameter('Имя').Set(sp_n/304.8)
	str_room.LookupParameter('Колонна').Set(sp_kolon/304.8)
	str_room.LookupParameter('ПКолонна').Set(sp_plo_kolon/304.8)
	str_room.LookupParameter('ППлинтус').Set(sp_plo_plin/304.8)
	str_room.LookupParameter('ППол').Set(sp_plo_pol/304.8)
	str_room.LookupParameter('ППотолок').Set(sp_plo_pot/304.8)
	str_room.LookupParameter('ПСтена').Set(sp_plo_sten/304.8)
	str_room.LookupParameter('Плинтус').Set(sp_plin/304.8)
	str_room.LookupParameter('Пол').Set(sp_pol/304.8)
	str_room.LookupParameter('Потолок').Set(sp_pot/304.8)
	str_room.LookupParameter('ПримечаниеС').Set(sp_prim/304.8)
	str_room.LookupParameter('Стена').Set(sp_sten/304.8)
	dotNETList = List[ElementId](elem_id_for_group)
	#group = doc.Create.NewGroup(dotNETList)
	#group.GroupType.Name = "ВОП " + number_room + ", " + name_room
	#room_for_group.append(group.Id)
	
	#dotNETList = List[ElementId](elem_id_for_group)
	#room_group 
	#Закрытие транзакции
	t.Commit()
	point_Y = lpy_w

	t.Start()
	group = doc.Create.NewGroup(dotNETList)
	group.GroupType.Name = "ВОП " + number_room# + ", " + name_room
	t.Commit()
# Закрытие транзакции

OUT = lst_mat_name_wall, lst_mat_name_floor, lst_mat_name_ceileng, lst_mat_name_columns, lst_mat_area_wall, lst_mat_area_columns
