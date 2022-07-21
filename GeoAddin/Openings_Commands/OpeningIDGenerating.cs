using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Configuration;
using System.Reflection;
using System.Collections.Specialized;
using System.Linq;
using System.Xml.Linq;
using System.Windows;
using RevitApplication = Autodesk.Revit.ApplicationServices.Application;
using Revit.GeometryConversion;


namespace GeoAddin
{
    [Transaction(TransactionMode.Manual)]
    public class OpeningIDGenerating : IExternalCommand
    {
        static UIApplication uiapp;
        static UIDocument uidoc;
        static RevitApplication app;
        static Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;

            //Получение всех проверенных отверстий без ID
            List<Element> openingInstances = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_DataDevices)
                        .WhereElementIsNotElementType()
                        .Where(ele => ele.Name.Contains("Отверстие"))
                        .Where(ele => ele.LookupParameter("Проверено").AsValueString() == "Да")
                        .Where(ele => ele.LookupParameter("ID_Отверстия").AsString() == "")
                        .ToList();

            for (int i = 0; i < openingInstances.Count; i++)
            {
                using (Transaction t = new Transaction(doc, "Генерация ID"))
                {
                    t.Start();
                    openingInstances[i].LookupParameter("ID_Отверстия").Set(i.ToString() + Guid.NewGuid().ToString() + openingInstances[i].LookupParameter("ADSK_Отверстие_Функция").AsString().ToLower());
                    t.Commit();
                }
            }
            return Result.Succeeded;
        }
    }   
}
