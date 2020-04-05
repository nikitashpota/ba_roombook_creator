#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace ba_roombook
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            
            // Access current selection

            #region Выбор помещений
            Selection sel = uidoc.Selection;
            ISelectionFilter selFilter = new SelcetionCategorie();
            IList<Reference> rooms = uidoc.Selection.PickObjects(ObjectType.Element, selFilter, "Select rooms");
            #endregion


            // Поиск типа текста
            //ElementId defaultTextTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
            //TextNoteOptions opts = new TextNoteOptions(defaultTextTypeId);
            //opts.HorizontalAlignment = HorizontalTextAlignment.Left;
            FilteredElementCollector collectorTextType = new FilteredElementCollector(doc);
            collectorTextType.OfClass(typeof(TextElementType));
            ElementId searchTextType_Id = collectorTextType.Cast<TextElementType>().First(vft => vft.Name == "2").Id;

            // Поиск всех групп в проекте.
            FilteredElementCollector collectorGroup = new FilteredElementCollector(doc);
            var groupInDoc = collectorGroup.OfCategory(BuiltInCategory.OST_IOSDetailGroups).WhereElementIsNotElementType().ToElements(); ;//OfClass(typeof(Group));

            // Создание и поиск вида
            #region Поиск вида для ведомости отделки
            IList<Element> views = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements();
            FilteredElementCollector collectorViews = new FilteredElementCollector(doc);
            collectorViews.OfClass(typeof(ViewFamilyType));
            ViewFamilyType viewFamilyType = collectorViews.Cast<ViewFamilyType>().First(vft => vft.ViewFamily == ViewFamily.Drafting);
            
            List<Element> listViews = new List<Element>();
            listViews = views.ToList<Element>();
            bool boolView = listViews.Any(x => (x as View)?.Name == "Форма 1 Ведомость отделки помещений");
            #endregion
            #region Создание вида. Проверка на наличе вида.
            ElementId viewId = ElementId.InvalidElementId;

            if (boolView)
            {
                View view_rb = listViews.Where(x => (x as View)?.Name == "Форма 1 Ведомость отделки помещений").First() as View;
                viewId = view_rb.Id;
            }
            else
            {
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Create view");

                    ViewDrafting view_rb = ViewDrafting.Create(doc, viewFamilyType.Id);
                    view_rb.Name = "Форма 1 Ведомость отделки помещений";
                    viewId = view_rb.Id;
                    tx.Commit();
                }
            }
            #endregion

            // Work with geometry rooms
            #region Заведение опции ограничивающих поверхностей
            SpatialElementBoundaryOptions spatialElementBoundaryOptions = new SpatialElementBoundaryOptions();
            spatialElementBoundaryOptions.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
            SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(doc, spatialElementBoundaryOptions);
            #endregion
            double koef = 0.01;
            double yLocationWall, yLocationFloor, yLocationCeiling;
            yLocationWall = yLocationFloor = yLocationCeiling = -0.003; // Отступ от оси Y (ед. изм. футы)
            double widthWallName = 30 / 304.8; // Ширина столбца СТЕНЫ ИЛИ ПЕРЕГОРОДКИ
            double widthWallArea = 15 / 304.8; // Ширина столбца КОЛ-ВО для СТЕНЫ ИЛИ ПЕРЕГОРОДКИ
            double widthFloorName = 30 / 304.8; // Ширина столбца ПОЛЫ
            double widthFloorArea = 15 / 304.8; // Ширина столбца КОЛ-ВО для ПОЛЫ
            double widthCeilingName = 30 / 304.8; // Ширина столбца ПОТОЛКИ
            double widthCeilingArea = 15 / 304.8; // Ширина столбца КОЛ-ВО для ПОТОЛКИ

            double xLocationFloorName = 0;
            double xLocationFloorArea = 0;
            xLocationFloorArea += widthFloorName + koef;

            double xLocationCeilingName = 0;
            xLocationCeilingName += widthFloorArea + xLocationFloorArea + koef;

            double xLocationCeilingArea = 0;
            xLocationCeilingArea += xLocationCeilingName + widthCeilingName + koef;

            double xLocationWallName = 0;
            xLocationWallName += xLocationCeilingArea + widthCeilingArea + koef;

            double xLocationWallArea = 0;
            xLocationWallArea += xLocationWallName +  widthWallName + koef;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create text");
                foreach (Reference room_ref in rooms)
                {
                    // Создание ICollection для группы

                    List<ElementId> listGroup = new List<ElementId>();

                    //Создание словаря для ВОП (по категориям Формы 1)

                    RoomDict<string, double> dictRoomWall = new RoomDict<string, double>();
                    RoomDict<string, double> dictRoomFloor = new RoomDict<string, double>();
                    RoomDict<string, double> dictRoomCeiling = new RoomDict<string, double>();

                    //Создание множества для проверки уникальности элемента

                    HashSet<Element> set_element = new HashSet<Element>();
                    Room room = (doc.GetElement(room_ref.ElementId)) as Room;

                    // Получить имя и номер помещения

                    var numberRoom = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString();
                    var nameRoom = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                    // SpatialElementGeometryResults - Приведены результаты расчета геометрии пространственных элементов

                    SpatialElementGeometryResults room_results = calculator.CalculateSpatialElementGeometry(room);

                    // Получить геометрию помещения для разложения ее на фэйсы
                    Solid roomSolid = room_results.GetGeometry();

                    // Создание переменной для проверки и вывода в TaskShow
                    string msg = string.Empty;
                    foreach (Face roomSolidFace in roomSolid.Faces)
                    {
                        foreach (SpatialElementBoundarySubface subface in room_results.GetBoundaryFaceInfo(roomSolidFace))
                        {
                            Element element_in_room = doc.GetElement(subface.SpatialBoundaryElement.HostElementId);
                            int categorie_name = (doc.GetElement(element_in_room.Id)).Category.Id.IntegerValue;
                            set_element.Add(element_in_room);
                            // Заполнение словаря для стен
                            if (categorie_name.Equals((int)BuiltInCategory.OST_Walls))
                            {
                                ICollection<ElementId> material_id = element_in_room.GetMaterialIds(false);
                                foreach (ElementId id in material_id)
                                {
                                    var param = doc.GetElement(id).LookupParameter("ADSK_Группирование");
                                    //msg += $"\nParameter: {param}\n";
                                    if (param is null) continue;
                                    string string_group = param.AsString();
                                    //msg += $"Parameter: {string_group}\n";
                                    try
                                    {
                                        if ((string_group == "Отделка"))// & !(set_element.Contains(element_in_room)))
                                        {
                                            double material_area = element_in_room.GetMaterialArea(id, false);
                                            string material_name = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString();
                                            double area = UnitUtils.ConvertFromInternalUnits(material_area, DisplayUnitType.DUT_SQUARE_METERS);
                                            var type = element_in_room.GetType();
                                            if (type.Equals(typeof(FamilyInstance)))
                                            {
                                                dictRoomWall[material_name] += area / 2;
                                            }
                                            else
                                            {
                                                dictRoomWall[material_name] += area;
                                            }

                                        }
                                        //msg += $"\nmaterial_name: {room_dict_wall.keys.elementat(0)}\n";
                                        //msg += $"\nmaterial_area: {room_dict_wall.values.sum()}\n";
                                    }
                                    catch (Exception ex)
                                    {
                                        msg += $"\n\n\t" + ex.Message
                                                + ":\n\n\t" + ex.Source
                                                + ":\n\n\t" + ex.TargetSite
                                                + "\n\n\t" + ex.StackTrace;
                                    }
                                }
                                
                            }
                            // Заполнение словаря для полов
                            else if (categorie_name.Equals((int)BuiltInCategory.OST_Floors))
                            {
                                ICollection<ElementId> material_id = element_in_room.GetMaterialIds(false);
                                foreach (ElementId id in material_id)
                                {
                                    var param = doc.GetElement(id).LookupParameter("ADSK_Группирование");
                                    //msg += $"\nParameter: {param}\n";
                                    if (param is null) continue;
                                    string string_group = param.AsString();
                                    //msg += $"Parameter: {string_group}\n";
                                    try
                                    {
                                        if ((string_group == "Отделка"))// & !(set_element.Contains(element_in_room)))
                                        {
                                            double material_area = element_in_room.GetMaterialArea(id, false);
                                            string material_name = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString();
                                            double area = UnitUtils.ConvertFromInternalUnits(material_area, DisplayUnitType.DUT_SQUARE_METERS);
                                            var type = element_in_room.GetType();
                                            if (type.Equals(typeof(FamilyInstance)))
                                            {
                                                dictRoomFloor[material_name] += area / 2;
                                            }
                                            else
                                            {
                                                dictRoomFloor[material_name] += area;
                                            }

                                        }
                                        //msg += $"\nmaterial_name: {room_dict_wall.keys.elementat(0)}\n";
                                        //msg += $"\nmaterial_area: {room_dict_wall.values.sum()}\n";
                                    }
                                    catch (Exception ex)
                                    {
                                        msg += $"\n\n\t" + ex.Message
                                                + ":\n\n\t" + ex.Source
                                                + ":\n\n\t" + ex.TargetSite
                                                + "\n\n\t" + ex.StackTrace;
                                    }
                                }
                            }
                            // Заполнение словаря для потолков
                            else if (categorie_name.Equals((int)BuiltInCategory.OST_Ceilings))
                            {
                                ICollection<ElementId> material_id = element_in_room.GetMaterialIds(false);
                                foreach (ElementId id in material_id)
                                {
                                    var param = doc.GetElement(id).LookupParameter("ADSK_Группирование");
                                    //msg += $"\nParameter: {param}\n";
                                    if (param is null) continue;
                                    string string_group = param.AsString();
                                    //msg += $"Parameter: {string_group}\n";
                                    try
                                    {
                                        if ((string_group == "Отделка"))// & !(set_element.Contains(element_in_room)))
                                        {
                                            double material_area = element_in_room.GetMaterialArea(id, false);
                                            string material_name = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString();
                                            double area = UnitUtils.ConvertFromInternalUnits(material_area, DisplayUnitType.DUT_SQUARE_METERS);
                                            var type = element_in_room.GetType();
                                            if (type.Equals(typeof(FamilyInstance)))
                                            {
                                                dictRoomCeiling[material_name] += area / 2;
                                            }
                                            else
                                            {
                                                dictRoomCeiling[material_name] += area;
                                            }

                                        }
                                        //msg += $"\nmaterial_name: {room_dict_wall.keys.elementat(0)}\n";
                                        //msg += $"\nmaterial_area: {room_dict_wall.values.sum()}\n";
                                    }
                                    catch (Exception ex)
                                    {
                                        msg += $"\n\n\t" + ex.Message
                                                + ":\n\n\t" + ex.Source
                                                + ":\n\n\t" + ex.TargetSite
                                                + "\n\n\t" + ex.StackTrace;
                                    }
                                }
                            }

                        }
                    }

                    //tx.Start("Create text");
                    foreach (var wallsInRoom in dictRoomWall.Keys.Select((Value, Index) => new { Value, Index }))
                    {
                        string textBlock_1 = wallsInRoom.Value;
                        XYZ xyzWallName = new XYZ(xLocationWallName, yLocationWall, 0);
                        TextNote textWallName = TextNote.Create(doc, viewId, xyzWallName, widthWallName, textBlock_1, searchTextType_Id);
                        listGroup.Add(textWallName.Id);

                        string textBlock_2 = Math.Round(dictRoomWall.Values.ElementAt(wallsInRoom.Index), 2).ToString();
                        XYZ xyzWallArea = new XYZ(xLocationWallArea, yLocationWall, 0);
                        TextNote textWallArea = TextNote.Create(doc, viewId, xyzWallArea, widthWallArea, textBlock_2, searchTextType_Id);
                        listGroup.Add(textWallArea.Id);

                        doc.Regenerate();
                        textWallArea.HorizontalAlignment = HorizontalTextAlignment.Right;
                        yLocationWall -= (textWallName.Height + koef);
                    }

                    foreach (var floorInRoom in dictRoomFloor.Keys.Select((Value, Index) => new { Value, Index }))
                    {
                        string textBlock_1 = floorInRoom.Value;
                        XYZ xyzFloorName = new XYZ(xLocationFloorName, yLocationFloor, 0);
                        TextNote textFloorName = TextNote.Create(doc, viewId, xyzFloorName, widthFloorName, textBlock_1, searchTextType_Id);
                        listGroup.Add(textFloorName.Id);

                        string textBlock_2 = Math.Round(dictRoomFloor.Values.ElementAt(floorInRoom.Index), 2).ToString();
                        XYZ xyzFloorArea = new XYZ(xLocationFloorArea, yLocationFloor, 0);
                        TextNote textFloorArea = TextNote.Create(doc, viewId, xyzFloorArea, widthFloorArea, textBlock_2, searchTextType_Id);
                        listGroup.Add(textFloorArea.Id);
                        doc.Regenerate();
                        textFloorArea.HorizontalAlignment = HorizontalTextAlignment.Right;
                        yLocationFloor -= (textFloorName.Height + koef);
                    }

                    foreach (var ceilingInRoom in dictRoomCeiling.Keys.Select((Value, Index) => new { Value, Index }))
                    {
                        string textBlock_1 = ceilingInRoom.Value;
                        XYZ xyzCeilingName = new XYZ(xLocationCeilingName, yLocationCeiling, 0);
                        TextNote textCeilingName = TextNote.Create(doc, viewId, xyzCeilingName, widthCeilingName, textBlock_1, searchTextType_Id);
                        listGroup.Add(textCeilingName.Id);
                        string textBlock_2 = Math.Round(dictRoomCeiling.Values.ElementAt(ceilingInRoom.Index), 2).ToString();
                        XYZ xyzCeilingArea = new XYZ(xLocationCeilingArea, yLocationCeiling, 0);
                        TextNote textCeilingArea = TextNote.Create(doc, viewId, xyzCeilingArea, widthCeilingArea, textBlock_2, searchTextType_Id);
                        listGroup.Add(textCeilingArea.Id);
                        doc.Regenerate();
                        textCeilingArea.HorizontalAlignment = HorizontalTextAlignment.Right;
                        yLocationCeiling -= (textCeilingName.Height + koef);
                    }
                    // Определение минимального положения строки
                    List<double> minList = new List<double>() { yLocationCeiling, yLocationFloor, yLocationWall };
                    yLocationCeiling = minList.Min();
                    yLocationFloor  = minList.Min();
                    yLocationWall = minList.Min();

                    // Создание и поиск группы
                    bool boolGroup = groupInDoc.Any(x => (x)?.Name == "ВОП_" + numberRoom + "_" + nameRoom);
                    



                    if (boolGroup)
                    {
                        var idGroup = listGroup.Where(x => (doc.GetElement(x))?.Name == "ВОП_" + numberRoom + "_" + nameRoom).First() as ElementId;
                        doc.Delete(idGroup);
                    }
                    var group = doc.Create.NewGroup(listGroup);
                    group.GroupType.Name = "ВОП_" + numberRoom + "_" + nameRoom;

                }
                tx.Commit();
            }
            

            return Result.Succeeded;
        }
    }
}
