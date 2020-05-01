#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks; // Библиотека для работы параллельными задачами
using System.Collections.Concurrent; //Библиотека, содержащая потокобезопасные коллекции

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
    public class MainRoomBook : IExternalCommand
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

            // Сортировка списка помещений
            var listRoomsSort = rooms.ToList<Reference>().OrderBy(x => ((doc.GetElement(x.ElementId)) as Room).Number);
            #endregion

            // Поиск типа текста
            FilteredElementCollector collectorTextType = new FilteredElementCollector(doc);
            collectorTextType.OfClass(typeof(TextElementType));
            ElementId searchTextType_Id = collectorTextType.Cast<TextElementType>().First(vft => vft.Name == "ADSK_Основной текст_2.0").Id;

            // Поиск всех групп в проекте.
            FilteredElementCollector collectorGroup = new FilteredElementCollector(doc);
            var groupInDoc = collectorGroup.OfClass((typeof(GroupType)));//collectorGroup.OfCategory(BuiltInCategory.OST_IOSDetailGroups).WhereElementIsNotElementType().ToElements(); ;//OfClass(typeof(Group));

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

            //Поиск требуемой строки для отделка
            FilteredElementCollector collectorAnnot = new FilteredElementCollector(doc);
            collectorAnnot.OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsElementType().ToElements();
            var annotString = collectorAnnot.First(x => x.Name == "BA_ANNOT_String for RoomBook");//.Cast<AnnotationSymbolType>().First(vft => vft.FamilyName == "BA_ANNOT_String for RoomBook").Id;
            // Work with geometry rooms
            #region Заведение опции ограничивающих поверхностей
            SpatialElementBoundaryOptions spatialElementBoundaryOptions = new SpatialElementBoundaryOptions();
            spatialElementBoundaryOptions.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
            SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(doc, spatialElementBoundaryOptions);
            #endregion


            #region Ширина и кол-во столбцов

            double koef = 2 / 304.8;
            double yLocationString = 0;
            double yLocationWall, yLocationFloor, yLocationCeiling, yLocationPlinth;
            yLocationWall = yLocationFloor = yLocationCeiling =  yLocationPlinth = - 1 * koef; // Отступ от оси Y
            double widtRoomNumber = 15 / 304.8;
            double widthRoomName = 30 / 304.8;
            double widthWallName = 40 / 304.8; // Ширина столбца СТЕНЫ ИЛИ ПЕРЕГОРОДКИ
            double widthWallArea = 10 / 304.8; // Ширина столбца КОЛ-ВО для СТЕНЫ ИЛИ ПЕРЕГОРОДКИ
            double widthFloorName = 40 / 304.8; // Ширина столбца ПОЛЫ
            double widthFloorArea = 10 / 304.8; // Ширина столбца КОЛ-ВО для ПОЛЫ
            double widthCeilingName = 30 / 304.8 ; // Ширина столбца ПОТОЛКИ
            double widthCeilingArea = 10 / 304.8; // Ширина столбца КОЛ-ВО для ПОТОЛКИ
            double widthPlinthName = 20 / 304.8; //Ширина столбца плинус
            double widthPlinthLenght = 10 / 304.8; //Ширина столбца длина плинтуса

            double xLocationFloorName = 0;

            xLocationFloorName += widtRoomNumber + widthRoomName;
            double xLocationFloorArea = 0;
            xLocationFloorArea += widthFloorName + widtRoomNumber + widthRoomName;

            double xLocationCeilingName = 0;
            xLocationCeilingName += widthFloorArea + xLocationFloorArea;

            double xLocationCeilingArea = 0;
            xLocationCeilingArea += xLocationCeilingName + widthCeilingName;

            double xLocationWallName = 0;
            xLocationWallName += xLocationCeilingArea + widthCeilingArea;

            double xLocationWallArea = 0;
            xLocationWallArea += xLocationWallName + widthWallName;

            double xLocationPlinthName = 0;
            xLocationPlinthName += xLocationWallArea + widthWallArea;

            double xLocationPlinthLenght = 0;
            xLocationPlinthLenght += xLocationPlinthName + widthPlinthName;

            #endregion

            string longTime = string.Empty;
            using (Transaction tx = new Transaction(doc))
            {
                
                //Запуск счетчика
                var watch = System.Diagnostics.Stopwatch.StartNew();

                // Запуск транзакции (ТЕСТ)
                tx.Start("Create text");
                foreach (Reference room_ref in listRoomsSort)
                {
                    // Создание ICollection для группы

                    List<ElementId> listGroup = new List<ElementId>();

                    //Создание словаря для ВОП (по категориям Формы 1)

                    RoomDict<string, double> dictRoomWall = new RoomDict<string, double>();
                    RoomDict<string, double> dictRoomFloor = new RoomDict<string, double>();
                    RoomDict<string, double> dictRoomCeiling = new RoomDict<string, double>();

                    //Создание множества для проверки уникальности элемента

                    HashSet<ElementId> setElement = new HashSet<ElementId>();

                    Room room = (doc.GetElement(room_ref.ElementId)) as Room;

                    // Получить имя и номер помещения

                    var numberRoom = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString();
                    var nameRoom = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                    var perimeter = Math.Round(UnitUtils.ConvertFromInternalUnits(room.get_Parameter(BuiltInParameter.ROOM_PERIMETER).AsDouble(), DisplayUnitType.DUT_METERS), 2); // UnitUtils.ConvertFromInternalUnits(material_area, DisplayUnitType.DUT_SQUARE_METERS);

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
                                        if ((string_group == "Отделка") & !(setElement.Contains(element_in_room.Id))) 
                                        {
                                            double material_area = element_in_room.GetMaterialArea(id, false);
                                            string materialName = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString();
                                            double area = UnitUtils.ConvertFromInternalUnits(material_area, DisplayUnitType.DUT_SQUARE_METERS);
                                            var type = element_in_room.GetType();
                                            if (type.Equals(typeof(FamilyInstance)))
                                            {
                                                dictRoomWall[materialName] += area / 2;
                                            }
                                            else
                                            {
                                                dictRoomWall[materialName] += area;
                                            }
                                        }
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
                                        if ((string_group == "Отделка") & !(setElement.Contains(element_in_room.Id)))// & !(set_element.Contains(element_in_room)))
                                        {
                                            double material_area = element_in_room.GetMaterialArea(id, false);
                                            string materialName = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString();
                                            double area = UnitUtils.ConvertFromInternalUnits(material_area, DisplayUnitType.DUT_SQUARE_METERS);
                                            var type = element_in_room.GetType();
                                            if (type.Equals(typeof(FamilyInstance)))
                                            {
                                                dictRoomFloor[materialName] += area / 2;
                                            }
                                            else
                                            {
                                                dictRoomFloor[materialName] += area;
                                            }

                                        }
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
                                        if ((string_group == "Отделка") & !(setElement.Contains(element_in_room.Id)))// & !(set_element.Contains(element_in_room)))
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

                            setElement.Add(element_in_room.Id);
                            foreach (var elem in setElement)
                            {
                                msg += $"Set : {doc.GetElement(elem).Name}\n";
                            }

                        }
                    }

                    foreach (var wallsInRoom in dictRoomWall.Keys.Select((Value, Index) => new { Value, Index }))
                    {
                        string textBlock_1 = wallsInRoom.Value;
                        XYZ xyzWallName = new XYZ(xLocationWallName + koef, yLocationWall , 0);
                        TextNote textWallName = TextNote.Create(doc, viewId, xyzWallName, widthWallName - 2 * koef, textBlock_1, searchTextType_Id);
                        listGroup.Add(textWallName.Id);

                        string textBlock_2 = Math.Round(dictRoomWall.Values.ElementAt(wallsInRoom.Index), 2).ToString();
                        XYZ xyzWallArea = new XYZ(xLocationWallArea + koef, yLocationWall, 0);
                        TextNote textWallArea = TextNote.Create(doc, viewId, xyzWallArea, widthWallArea - 2 * koef, textBlock_2, searchTextType_Id);
                        listGroup.Add(textWallArea.Id);

                        doc.Regenerate();
                        textWallArea.HorizontalAlignment = HorizontalTextAlignment.Right;
                        yLocationWall -= (textWallName.Height);
                    }

                    foreach (var floorInRoom in dictRoomFloor.Keys.Select((Value, Index) => new { Value, Index }))
                    {
                        string textBlock_1 = floorInRoom.Value;
                        XYZ xyzFloorName = new XYZ(xLocationFloorName + koef, yLocationFloor, 0);
                        TextNote textFloorName = TextNote.Create(doc, viewId, xyzFloorName, widthFloorName - 2 * koef, textBlock_1, searchTextType_Id);
                        listGroup.Add(textFloorName.Id);

                        string textBlock_2 = Math.Round(dictRoomFloor.Values.ElementAt(floorInRoom.Index), 2).ToString();
                        XYZ xyzFloorArea = new XYZ(xLocationFloorArea + koef, yLocationFloor, 0);
                        TextNote textFloorArea = TextNote.Create(doc, viewId, xyzFloorArea, widthFloorArea - 2 * koef, textBlock_2, searchTextType_Id);
                        listGroup.Add(textFloorArea.Id);
                        doc.Regenerate();
                        textFloorArea.HorizontalAlignment = HorizontalTextAlignment.Right;
                        yLocationFloor -= (textFloorName.Height);
                    }

                    foreach (var ceilingInRoom in dictRoomCeiling.Keys.Select((Value, Index) => new { Value, Index }))
                    {
                        string textBlock_1 = ceilingInRoom.Value;
                        XYZ xyzCeilingName = new XYZ(xLocationCeilingName + koef, yLocationCeiling, 0);
                        TextNote textCeilingName = TextNote.Create(doc, viewId, xyzCeilingName, widthCeilingName - 2 * koef, textBlock_1, searchTextType_Id);
                        listGroup.Add(textCeilingName.Id);
                        string textBlock_2 = Math.Round(dictRoomCeiling.Values.ElementAt(ceilingInRoom.Index), 2).ToString();
                        XYZ xyzCeilingArea = new XYZ(xLocationCeilingArea + koef, yLocationCeiling, 0);
                        TextNote textCeilingArea = TextNote.Create(doc, viewId, xyzCeilingArea, widthCeilingArea - 2 * koef, textBlock_2, searchTextType_Id);
                        listGroup.Add(textCeilingArea.Id);
                        doc.Regenerate();
                        textCeilingArea.HorizontalAlignment = HorizontalTextAlignment.Right;
                        yLocationCeiling -= (textCeilingName.Height);
                    }

                    if (room.LookupParameter("ОП_Отделка плинтуса").AsString() != (null)) { 
                    XYZ xyzPlinthName = new XYZ(xLocationPlinthName + koef, yLocationPlinth, 0);
                    string textBlockPlinth = room.LookupParameter("ОП_Отделка плинтуса").AsString();
                    TextNote textPlinthName = TextNote.Create(doc, viewId, xyzPlinthName, widthPlinthName - 2 * koef, textBlockPlinth, searchTextType_Id);
                    listGroup.Add(textPlinthName.Id);
                    textPlinthName.HorizontalAlignment = HorizontalTextAlignment.Left;


                    XYZ xyzPlinthLength = new XYZ(xLocationPlinthLenght + koef, yLocationPlinth, 0);
                    string textBlockPerimeter = perimeter.ToString();
                    TextNote textPlinthLenght = TextNote.Create(doc, viewId, xyzPlinthLength, widthPlinthLenght - 2 * koef, textBlockPerimeter, searchTextType_Id);
                    listGroup.Add(textPlinthLenght.Id);
                        textPlinthLenght.HorizontalAlignment = HorizontalTextAlignment.Right;

                        yLocationCeiling -= (textPlinthName.Height);
                    }
                    

                    // Определение минимального положения строки
                    List<double> minList = new List<double>() { yLocationCeiling, yLocationFloor, yLocationWall, yLocationPlinth};

                    yLocationCeiling = minList.Min() - 1 * koef;
                    yLocationFloor  = minList.Min() - 1 * koef;
                    yLocationWall = minList.Min() - 1 * koef;
                    yLocationPlinth = minList.Min() - 1 * koef;

                    #region Манипуляция со строкой 

                    if (dictRoomWall.Values.Count > 0 |
                    dictRoomFloor.Values.Count > 0 |
                    dictRoomCeiling.Values.Count > 0)
                    {
                        var newString = doc.Create.NewFamilyInstance(new XYZ(0, yLocationString, 0), annotString as FamilySymbol, doc.GetElement(viewId) as View);//(new XYZ(0, yLocationString, 0), annotStringId, viewId);

                        newString.LookupParameter("H").Set(yLocationString + 1 * koef - (minList.Min()));
                        newString.LookupParameter("NUMBER").Set(widtRoomNumber);
                        newString.LookupParameter("NAME").Set(widthRoomName);
                        newString.LookupParameter("CEILING").Set(widthCeilingName);
                        newString.LookupParameter("CEILING_AREA").Set(widthCeilingArea);
                        newString.LookupParameter("FLOOR").Set(widthFloorName);
                        newString.LookupParameter("FLOOR_AREA").Set(widthFloorArea);
                        newString.LookupParameter("WALL").Set(widthWallName);
                        newString.LookupParameter("WALL_AREA").Set(widthWallArea);
                        newString.LookupParameter("PLINTH").Set(widthPlinthName);
                        newString.LookupParameter("PLINTH_LENGTH").Set(widthPlinthLenght);
                        newString.LookupParameter("NUMBER_ROOM").Set(numberRoom);
                        newString.LookupParameter("NAME_ROOM").Set(nameRoom);

                        listGroup.Add(newString.Id);

                        yLocationString = minList.Min() - 1 * koef;
                    }
                    #endregion

                    // Создание и поиск группы

                    bool boolGroup = groupInDoc.Any(x => (x)?.Name == "ВОП_" + numberRoom + "_" + nameRoom);

                    if (boolGroup)
                    {
                        var idGroup =from element in groupInDoc where element.Name ==  "ВОП_" + numberRoom + "_" + nameRoom select element;//.Where(x => (doc.GetElement(x))?.Name == "ВОП_" + numberRoom + "_" + nameRoom).First() as ElementId;
                        doc.Delete(idGroup.First().Id);
                    }

                    //msg += $"Result : {groupType.Name}\n";
                    //TaskDialog.Show("sd", msg);
                    try
                    {
                        var group = doc.Create.NewGroup(listGroup);
                        group.GroupType.Name = "ВОП_" + numberRoom + "_" + nameRoom;
                    }
                    catch { }

                }
                tx.Commit();

                // Конец счетчика
                watch.Stop();
                var elapsedMs = watch.Elapsed;//.ElapsedMilliseconds;
                longTime += $"\nTIME: {elapsedMs}\n";

            }
            


            TaskDialog.Show("Result", longTime);
            return Result.Succeeded;
        }


        // TEST 
        public class FailUI : IFailuresProcessor
        {
            public void Dismiss(Document document)
            {
                // This method is being called in case of exception or document destruction to 
                // dismiss any possible pending failure UI that may have left on the screen
            }

            public FailureProcessingResult ProcessFailures(FailuresAccessor failuresAccessor)
            {
                IList<FailureResolutionType> resolutionTypeList = new List<FailureResolutionType>();
                IList<FailureMessageAccessor> failList = new List<FailureMessageAccessor>();
                // Inside event handler, get all warnings
                failList = failuresAccessor.GetFailureMessages();
                string errorString = "";
                bool hasFailures = false;
                foreach (FailureMessageAccessor failure in failList)
                {
                    // check how many resolutions types were attempted to try to prevent
                    // entering infinite loop
                    resolutionTypeList = failuresAccessor.GetAttemptedResolutionTypes(failure);
                    if (resolutionTypeList.Count >= 3)
                    {
                        TaskDialog.Show("Error", "Cannot resolve failures - transaction will be rolled back.");
                        return FailureProcessingResult.ProceedWithRollBack;
                    }

                    errorString += "IDs ";
                    foreach (ElementId id in failure.GetFailingElementIds())
                    {
                        errorString += id.ToString() + ", ";
                        hasFailures = true;
                    }
                    errorString += "\nWill be deleted because: " + failure.GetDescriptionText() + "\n";
                    failuresAccessor.DeleteElements(failure.GetFailingElementIds() as IList<ElementId>);
                }
                if (hasFailures)
                {
                    TaskDialog.Show("Error", errorString);
                    return FailureProcessingResult.ProceedWithCommit;
                }

                return FailureProcessingResult.Continue;
            }
        }

    }
}
