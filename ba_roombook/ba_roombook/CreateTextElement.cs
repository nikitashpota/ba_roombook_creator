using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ba_roombook
{
    public class СlassCreateTextElement
    {
        //public static double FuncCreateTextElement(
        //    RoomDict<string, double> myDict, 
        //    double xLocationElement, 
        //    double xLocationArea, 
        //    double yLocation,
        //    double widthElement,
        //    double widthArea,
        //    Document doc,
        //    ElementId viewId,
        //    ElementId searchTextType_Id,
        //    double koef)
        //{
        //    try
        //    {

        //        foreach (var elemInRoom in myDict.Keys.Select((Value, Index) => new { Value, Index }))
        //        {
        //            string textBlock_1 = elemInRoom.Value;
        //            XYZ xyzWallName = new XYZ(xLocationElement, yLocation, 0);
        //            TextNote textWallName = TextNote.Create(doc, viewId, xyzWallName, widthElement, textBlock_1, searchTextType_Id);

        //            string textBlock_2 = Math.Round(myDict.Values.ElementAt(elemInRoom.Index), 2).ToString();
        //            XYZ xyzWallArea = new XYZ(xLocationArea, yLocation, 0);
        //            TextNote textWallArea = TextNote.Create(doc, viewId, xyzWallArea, widthArea, textBlock_2, searchTextType_Id);
        //            doc.Regenerate();

        //            return yLocation -= (textWallName.Height + koef);
        //        }
        //    }
        //    catch (ArgumentException)
        //    {
        //        return 0;
        //    }
        //}

    }
}
