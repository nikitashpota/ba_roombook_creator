#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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

            Selection sel = uidoc.Selection;
            ISelectionFilter selFilter = new SelcetionCategorie();
            IList<Reference> rooms = uidoc.Selection.PickObjects(ObjectType.Element, selFilter, "Select rooms");

            // Crete views

            IList <Element> views = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));
            ViewFamilyType viewFamilyType = collector.Cast<ViewFamilyType>().First(vft => vft.ViewFamily == ViewFamily.Drafting);

            //List<Element> subProducts = Model.subproducts as List<Element>;
            List<Element> views_list = new List<Element>();
            //IList<Element> obj = `Your Data Will Be Here`;
            views_list = views.ToList<Element>();
            bool result_view = views_list.Any(x => (x as View)?.Name == "Форма 1 Ведомость отделки помещений");

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

            FilteredElementCollector col
              = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.INVALID)
                .OfClass(typeof(Wall));

            // Filtered element collector is iterable

            foreach (Element e in col)
            {
                Debug.Print(e.Name);
            }

            // Modify document within a transaction

            /*using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Transaction Name");
                tx.Commit();
            }*/

            return Result.Succeeded;
        }
    }
}
