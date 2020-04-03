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
            if (result_view)
            {
                View view_rb = views_list.Where(x => (x as View)?.Name == "Форма 1 Ведомость отделки помещений").First() as View;
            }
            else
            {
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Create view");

                    ViewDrafting view_rb = ViewDrafting.Create(doc, viewFamilyType.Id);
                    view_rb.Name = "Форма 1 Ведомость отделки помещений";

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

            #region List for roombook
            List<List<string>> walls = new List<List<string>>();
            List<List<string>> floors = new List<List<string>>();
            List<List<string>> ceiling = new List<List<string>>();
            List<List<string>> columns = new List<List<string>>();
            List<List<double>> walls_area = new List<List<double>>();
            List<List<double>> floors_area = new List<List<double>>();
            List<List<double>> ceiling_area = new List<List<double>>();
            List<List<double>> columns_area = new List<List<double>>();
            #endregion

            foreach (Reference room_ref in rooms)
            {
                //Создание словаря для ВОП (по категориям Формы 1)
                RoomDict<string, double> room_dict_wall = new RoomDict<string, double>();
                //Создание множества для проверки уникальности элемента
                HashSet<Element> set_element = new HashSet<Element>();
                Room room = (doc.GetElement(room_ref.ElementId)) as Room;

                SpatialElementGeometryResults room_results = calculator.CalculateSpatialElementGeometry(room);
                Solid roomSolid = room_results.GetGeometry();

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
                                    if ((string_group == "Отделка") & !(set_element.Contains(element_in_room)))
                                    {
                                        double material_area = element_in_room.GetMaterialArea(id, false);
                                        string material_name = doc.GetElement(id).get_Parameter(BuiltInParameter.MATERIAL_NAME).AsString();
                                        double area = UnitUtils.ConvertFromInternalUnits(material_area, DisplayUnitType.DUT_SQUARE_METERS);
                                        var type = element_in_room.GetType();
                                        if (type.Equals(typeof(FamilyInstance)))
                                        {
                                            room_dict_wall[material_name] += area/2;
                                        }
                                        else
                                        {
                                            room_dict_wall[material_name] += area;
                                        }
                                        //msg += $"Material name: {room_dict_wall.Values.Sum()}\n";
                                    }

                                }
                                catch (Exception ex)
                                {
                                    msg += $"\n\n\t" + ex.Message                                            + ":\n\n\t" + ex.Source                                            + ":\n\n\t" + ex.TargetSite                                            + "\n\n\t" + ex.StackTrace;
                                }

                            }

                        }
                    }
                }
                //Добавить в список значение словарей
                walls.Add(room_dict_wall.Keys.ToList());
                walls_area.Add(room_dict_wall.Values.ToList());

                
            }

            #region Create double
            double lpy_w, lpy_f, lpy_c, apy_w,apy_f, apy_c, lpy_cn, apy_cn = -0.003;
            var point_Y = 0;
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

                foreach (var walls_room in walls.Select((Value_wr, Index_wr) => new { Value_wr, Index_wr }))
                {
                    var text_wall = TextNote.Create(doc, view_rb.Id);
                }



            }

            return Result.Succeeded;
        }
    }
}
