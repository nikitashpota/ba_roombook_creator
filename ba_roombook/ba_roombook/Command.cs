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
            FilteredElementCollector collector_texttype = new FilteredElementCollector(doc);
            collector_texttype.OfClass(typeof(TextElementType));
            ElementId searchTextType_Id = collector_texttype.Cast<TextElementType>().First(vft => vft.Name == "2").Id;


            // Crete views

            #region Поиск вида для ведомости отделки
            IList<Element> views = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));
            ViewFamilyType viewFamilyType = collector.Cast<ViewFamilyType>().First(vft => vft.ViewFamily == ViewFamily.Drafting);
            
            List<Element> views_list = new List<Element>();
            views_list = views.ToList<Element>();
            bool result_view = views_list.Any(x => (x as View)?.Name == "Форма 1 Ведомость отделки помещений");
            #endregion
            #region Создание вида. Проверка на наличе вида.
            ElementId viewId = ElementId.InvalidElementId;

            if (result_view)
            {
                View view_rb = views_list.Where(x => (x as View)?.Name == "Форма 1 Ведомость отделки помещений").First() as View;
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
            double koef = 0.003;
            double yLocationWall, yLocationFloor, yLocationCeiling;
            yLocationWall = yLocationFloor = yLocationCeiling = -0.003;
            double widthWallName = 30 / 304.8; // Ширина столбца СТЕНЫ ИЛИ ПЕРЕГОРОДКИ
            double widthWallArea = 15 / 304.8; // Ширина столбца КОЛ-ВО для СТЕНЫ ИЛИ ПЕРЕГОРОДКИ
            double widthFloorName = 30 / 304.8; // Ширина столбца ПОЛЫ
            double widthFloorArea = 15 / 304.8; // Ширина столбца КОЛ-ВО для ПОЛЫ
            double xLocationWallName = 0;
            double xLocationWallArea = 0;
            xLocationWallArea += widthWallName+ koef;

            using (Transaction tx = new Transaction(doc))
            {
                foreach (Reference room_ref in rooms)
                {
                    //Создание словаря для ВОП (по категориям Формы 1)
                    RoomDict<string, double> dictRoomWall = new RoomDict<string, double>();
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

                    // Создание 
                    string msg = string.Empty;
                    foreach (Face roomSolidFace in roomSolid.Faces)
                    {
                        foreach (SpatialElementBoundarySubface subface in room_results.GetBoundaryFaceInfo(roomSolidFace))
                        {
                            Element element_in_room = doc.GetElement(subface.SpatialBoundaryElement.HostElementId);
                            int categorie_name = (doc.GetElement(element_in_room.Id)).Category.Id.IntegerValue;
                            set_element.Add(element_in_room);
                            // For walls
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
                        }
                    }

                    tx.Start("Create text");
                    foreach (var wallsInRoom in dictRoomWall.Keys.Select((Value, Index) => new { Value, Index }))//(var walls_in_room in room_dict_wall.Keys)
                    {
                        string textBlock_1 = wallsInRoom.Value;
                        XYZ xyzWallName = new XYZ(xLocationWallName, yLocationWall, 0);
                        TextNote textWallName = TextNote.Create(doc, viewId, xyzWallName, widthWallName, textBlock_1, searchTextType_Id);

                        string textBlock_2 = Math.Round(dictRoomWall.Values.ElementAt(wallsInRoom.Index), 2).ToString();
                        XYZ xyzWallArea = new XYZ(xLocationWallArea, yLocationWall, 0);
                        TextNote textWallArea = TextNote.Create(doc, viewId, xyzWallArea, widthWallArea, textBlock_2, searchTextType_Id);
                        doc.Regenerate();

                        yLocationWall -= (textWallName.Height + koef);
                    }

                    tx.Commit();

                }
            }

            #region Мусор
            /*
            #region Create double
            double lpy_w, lpy_f, lpy_c, apy_w,apy_f, apy_c, lpy_cn, apy_cn;
            y_location_wall_name = lpy_f = lpy_c = y_location_wall_area = apy_f = apy_c = lpy_cn = apy_cn = -0.003;
            double point_Y = 0;
            double sp_sten = 25/304.8;
            #endregion
            
            string prs = string.Empty;
            foreach (var room_ref in rooms.Select((Value, Index) => new { Value, Index }))
            {

                prs += $"Index: {room_ref.Value}\n";
                TaskDialog.Show("SD", prs + $"\nGood work");
                Room room = (doc.GetElement(room_ref.Value.ElementId)) as Room;
                var number_room = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString();
                var name_room = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                //Создание списка с ID для групппирования
                List<ElementId> element_id_for_group = new List<ElementId>();

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("test");
                    foreach (var walls_room in walls.Select((Value_wr, Index_wr) => new { Value_wr, Index_wr }))
                    {
                        XYZ text_lock = new XYZ(0, lpy_w, 0);
                        string text_1 = walls_room.Value_wr[0];
                        TextNote text_wall = TextNote.Create(doc, view_id, text_lock, sp_sten, text_1, opts);
                    }
                    tx.Commit();
                }

            }*/
            #endregion

            return Result.Succeeded;
        }
    }
}
