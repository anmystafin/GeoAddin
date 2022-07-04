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

           
            
            while (clickedon == false & clickedoff == false)
            {
                continue;
            }
            if (clickedon == true)
            {
                //Переменные исходных данных
                Level level = null;
                RevitLinkInstance link = null;
                double openingBoundOffset;
                string openingFunc;
                Document linkDoc = null;

                //Взятие исходных данных из окна и запись в основные переменные
                List<Level> levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast<Level>().ToList();
                List<RevitLinkInstance> linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
                foreach (Level item in levels) {if (item.Name == win.LevelBox.Text) {level = item;}}
                foreach (RevitLinkInstance item in linkInstances){if (item.Name == win.LinkInstance.Text){link = item; linkDoc = link.Document; } }
                openingBoundOffset = Convert.ToDouble(win.BoundOffset.Text);
                openingFunc = win.OpeningFunc.Text;
                //
                //Получение типоразмера отверстия
                List<FamilySymbol> openingFamilySymbol = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DataDevices)
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .Where(el => el.FamilyName == "Отверстие_Универсальное")
                    .ToList();

                //Получение перекрытий и стен из связанного файла
                List<Element> linkWalls = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .Where(wall => wall.LookupParameter("Зависимость снизу").AsValueString() == level.Name)
                    .ToList();
                List<Element> linkFloors = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .Where(floor => floor.LookupParameter("Уровень").AsValueString() == level.Name)
                    .ToList();
                List<Element> linkElements = linkFloors.Concat(linkWalls).ToList();

                //Получение воздуховодов, труб, кабельных лотков и коробов с активного вижа
                List<Element> ductCurves = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_DuctCurves)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();
                List<Element> pipeCurves = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_PipeCurves)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();
                List<Element> cableTrays = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_CableTray)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();
                List<Element> conduits = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_Conduit)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();
                List<Element> mepElements = ductCurves.Concat(pipeCurves).Concat(cableTrays).Concat(conduits).ToList();

                //Проверка пересечения элементов и получение точек вставки семейств
                
                foreach (Element mepElement in mepElements)
                {
                    foreach (Element linkElement in linkElements)
                    {
                       Solid solid = GetSolidOfElement(linkElement);
                       Line line = GetCenterLineOfElement(mepElement);
                       SolidCurveIntersection curves =  solid.IntersectWithCurve(line, new SolidCurveIntersectionOptions());
                       using (Transaction t = new Transaction(doc, "Генерация отверстий"))
                        {
                            t.Start();
                            foreach (Line curve in curves)
                            {

                                doc.Create.NewFamilyInstance(curve.Origin, openingFamilySymbol[0], doc.ActiveView);


                            }
                            t.Commit();
                        }
                        
                    }
                }

            }


                return Result.Succeeded;
        }
        private static Solid GetSolidOfElement(Element element)
        {
            Solid elementSolid = null;
            GeometryElement geoElement = element.get_Geometry(new Options());
            foreach (GeometryObject geoObject in geoElement)
            {
                GeometryInstance instance = geoObject as GeometryInstance;
                if (instance != null)
                {
                    foreach (GeometryObject instObj in instance.GetInstanceGeometry())
                    {
                        elementSolid = instObj as Solid;
                    }
                }
            }

            return elementSolid;
        }
        private static Line GetCenterLineOfElement(Element element)
        {
            Line elementLine = null;
            GeometryElement geoElement = element.get_Geometry(new Options());
            foreach (GeometryObject geoObject in geoElement)
            {
                GeometryInstance instance = geoObject as GeometryInstance;
                if (instance != null)
                {
                    foreach (GeometryObject instObj in instance.GetInstanceGeometry())
                    {
                        int insId = instObj.GraphicsStyleId.IntegerValue;
                        if (doc.GetElement(new ElementId(insId)).Name == "Осевая линия" )
                        {
                            elementLine = instObj as Line;
                        }

                        
                    }
                }
            }

            return elementLine;
        }
    }
}
