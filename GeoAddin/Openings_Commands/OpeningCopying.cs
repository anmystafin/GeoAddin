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
                //Получение типоразмера отверстия
                List<FamilySymbol> wallOpeningFamilySymbol = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .WhereElementIsElementType()
                    .Where(el => el.Name == "Отверстие_вСтене")
                    .Cast<FamilySymbol>()
                    .ToList();
                List<FamilySymbol> floorOpeningFamilySymbol = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .WhereElementIsElementType()
                    .Where(el => el.Name == "Отверстие_вПолу")
                    .Cast<FamilySymbol>()
                    .ToList();
                //Получение всех проверенных отверстий с заполненным ID из линкованного файла, а также стен и перекрытий в текущем доке
                List<Element> openingInstancesFromMEP = new FilteredElementCollector(linkDoc)
                            .OfCategory(BuiltInCategory.OST_DataDevices)
                            .WhereElementIsNotElementType()
                            .Where(ele => ele.Name.Contains("Отверстие"))
                            .Where(ele => ele.LookupParameter("Проверено").AsValueString() == "Да")
                            .Where(ele => ele.LookupParameter("ID_Отверстия").AsString() != "")
                            .ToList();
                using (Transaction t = new Transaction(doc, "Копирование отверстий"))
                {
                    t.Start();
                    foreach (Element opening in openingInstancesFromMEP)
                    {

                        Outline outline = new Outline(opening.get_BoundingBox(null).Min, opening.get_BoundingBox(null).Max);
                        BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
                        List<Element> walls = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .WhereElementIsNotElementType()
                                .WherePasses(filter)
                                .ToList();
                        List<Element> floors = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Floors)
                                .WhereElementIsNotElementType()
                                .WherePasses(filter)
                                .ToList();
                        XYZ locationPoint = new XYZ(
                            (opening.Location as LocationPoint).Point.X,
                            (opening.Location as LocationPoint).Point.Y,
                            (opening.Location as LocationPoint).Point.Z);
                        if (walls.Count > 0 && floors.Count == 0)
                        {   
                            
                           FamilyInstance instance = doc.Create.NewFamilyInstance(locationPoint, wallOpeningFamilySymbol[0], walls[0], Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                           instance.LookupParameter("ADSK_Размер_Длина").Set(opening.LookupParameter("ADSK_Размер_Длина").AsDouble());
                           instance.LookupParameter("ADSK_Размер_Ширина").Set(opening.LookupParameter("ADSK_Размер_Ширина").AsDouble());
                           instance.LookupParameter("ID_Отверстия").Set(opening.LookupParameter("ID_Отверстия").AsString());
                           instance.LookupParameter("ADSK_Отверстие_Функция").Set(opening.LookupParameter("ADSK_Отверстие_Функция").AsString());
                            
                            
                        }
                        
                        else
                        {
                            FamilyInstance instance = doc.Create.NewFamilyInstance(locationPoint, floorOpeningFamilySymbol[0], floors[0], Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            instance.LookupParameter("ADSK_Размер_Длина").Set(opening.LookupParameter("ADSK_Размер_Длина").AsDouble());
                            instance.LookupParameter("ADSK_Размер_Ширина").Set(opening.LookupParameter("ADSK_Размер_Ширина").AsDouble());
                            instance.LookupParameter("ID_Отверстия").Set(opening.LookupParameter("ID_Отверстия").AsString());
                            instance.LookupParameter("ADSK_Отверстие_Функция").Set(opening.LookupParameter("ADSK_Отверстие_Функция").AsString());
                            List<Element> upFloors = new List<Element>();
                            List<Element> downFloors = new List<Element>();
                            List<Element> otherFloors = floors.GetRange(1, floors.Count-1);
                            double upcut = 0;
                            double downcut = 0;
                            if (otherFloors.Count > 0)
                            {


                                foreach (Element floor in otherFloors)
                                {
                                    if (floor.get_BoundingBox(null).Max.Z > floors[0].get_BoundingBox(null).Max.Z)
                                    {
                                        upFloors.Add(floor);
                                    }
                                    else { downFloors.Add(floor); }
                                }
                                
                                if (upFloors.Count > 0)
                                {
                                    upcut = upFloors.Select(el => el.LookupParameter("Толщина").AsDouble()).Sum();

                                }
                                if (downFloors.Count > 0)
                                {
                                    downcut = downFloors.Select(el => el.LookupParameter("Толщина").AsDouble()).Sum();
                                }
                            }
                            
                            instance.LookupParameter("Вырез_Вверх").Set(upcut);
                            instance.LookupParameter("Вырез_Вниз").Set(downcut);
                            foreach (Element floor in floors)
                            {
                                try
                                {
                                    InstanceVoidCutUtils.AddInstanceVoidCut(doc, floor, instance);
                                }
                                catch { }
                            }

                        }
                        foreach (Element fwall in walls)
                        {
                            foreach(Element swall in walls)
                            {
                                try
                                {
                                    JoinGeometryUtils.JoinGeometry(doc, fwall, swall);
                                }
                                catch { }
                                
                            }
                        }
                        
                    }
                    t.Commit();
                }
            }

         return Result.Succeeded;
        }
    }   
}
