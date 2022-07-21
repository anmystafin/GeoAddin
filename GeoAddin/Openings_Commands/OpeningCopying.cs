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
    public class OpeningCopying : IExternalCommand
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

            OpeningCopyWindow win = new OpeningCopyWindow(uiapp);
            win.ShowDialog();
            bool clickedon = win.clickedon;
            bool clickedoff = win.clickedoff;

            while (clickedon == false & clickedoff == false)
            {
                continue;
            }
            if (clickedon == true)
            {
                //Переменные исходных данных
                RevitLinkInstance link = null;
                Document linkDoc = null;

                List<RevitLinkInstance> linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
                foreach (RevitLinkInstance item in linkInstances)
                {
                    foreach (Document document in app.Documents)
                    {
                        if ((item.Name == win.LinkInstance.Text) && (item.Name.Contains(document.Title))) { link = item; linkDoc = document; }
                    }
                }
                //Получение всех проверенных отверстий с заполненным ID из линкованного файла
                List<Element> openingInstancesFromMEP = new FilteredElementCollector(linkDoc)
                            .OfCategory(BuiltInCategory.OST_DataDevices)
                            .WhereElementIsNotElementType()
                            .Where(ele => ele.Name.Contains("Отверстие"))
                            .Where(ele => ele.LookupParameter("Проверено").AsValueString() == "Да")
                            .Where(ele => ele.LookupParameter("ID_Отверстия").AsString() != "")
                            .ToList();
            }

         return Result.Succeeded;
        }
    }   
}
