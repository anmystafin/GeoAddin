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

namespace GeoAddin
{
    [Transaction(TransactionMode.Manual)]
    public class OpeningGenerating : IExternalCommand
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

            OpeningGenWindow win = new OpeningGenWindow(uiapp);
            win.ShowDialog();
            bool clickedon = win.clickedon;
            bool clickedoff = win.clickedoff;

            //Переменные исходных данных
            Level level;
            RevitLinkInstance link;
            double openingBoundOffset;
            string openingFunc ;

            
            while (clickedon == false & clickedoff == false)
            {
                continue;
            }
            if (clickedon == true)
            {
                //Взятие исходных данных из окна и запись в основные переменные
                List<Level> levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast<Level>().ToList();
                List<RevitLinkInstance> linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
                foreach (Level item in levels) {if (item.Name == win.LevelBox.Text) {level = item;}}
                foreach (RevitLinkInstance item in linkInstances){if (item.Name == win.LinkInstance.Text){link = item;}}
                openingBoundOffset = Convert.ToDouble(win.BoundOffset.Text);
                openingFunc = win.OpeningFunc.Text;
                //
                



            }


                return Result.Succeeded;
        }
    }
}
