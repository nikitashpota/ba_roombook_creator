#region Namespace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace ba_roombook
{
    public class SelcetionCategorie : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            if (element.Category.Id.IntegerValue == BuiltInCategory.OST_Rooms.GetHashCode())
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
}
