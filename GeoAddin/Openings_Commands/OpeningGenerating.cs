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
                List<ElementId> createdOpeningInstances = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfCategory(BuiltInCategory.OST_DataDevices)
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(ele => ele.Name.Contains("Отверстие"))
                        .Where(ele => ele.LookupParameter("Проверено").AsValueString() == "Нет")
                        .Select(ele => ele.Id)
                        .ToList();
                
                using (Transaction t = new Transaction(doc, "Удаление существующих отверстий"))
                {
                    t.Start();
                    doc.Delete(createdOpeningInstances);
                    t.Commit();
                }
                
                List<Level> levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast<Level>().ToList();
                List<RevitLinkInstance> linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
                foreach (Level item in levels) { if (item.Name == win.LevelBox.Text) { level = item; } }
                foreach (RevitLinkInstance item in linkInstances)
                {
                    foreach (Document document in app.Documents)
                    {
                        if ((item.Name == win.LinkInstance.Text) && (item.Name.Contains(document.Title))) { link = item; linkDoc = document; }
                    }

                }
                string linkDocTitle = linkDoc.Title;
                openingBoundOffset = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(win.BoundOffset.Text), UnitTypeId.Millimeters);
                openingFunc = win.OpeningFunc.Text;
                //
                //Получение типоразмера отверстия
                List<FamilySymbol> wallOpeningFamilySymbol = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DataDevices)
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .Where(el => el.FamilyName == "Отверстие_вСтене")
                    .ToList();
                List<FamilySymbol> floorOpeningFamilySymbol = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DataDevices)
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .Where(el => el.FamilyName == "Отверстие_вПолу")
                    .ToList();

                //Получение перекрытий и стен из связанного файла
                List<Element> linkWalls = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .Where(wall => wall.LookupParameter("Зависимость снизу").AsValueString() == level.Name)
                    .ToList();
                List<Element> linkFloors = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Floor))
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

                //Проверка пересечения элементов и генерация семейств отверстий
                
                foreach (Element mepElement in mepElements)
                {
                   foreach (Element linkElement in linkElements)
                   {
                            Solid solid = GetSolidOfElement(linkElement);
                            Line line = GetCenterLineOfElement(mepElement);
                            
                            Solid lineSolid = GetSolidOfElement(mepElement);
                            SolidCurveIntersection intersectionCurves = solid.IntersectWithCurve(line, new SolidCurveIntersectionOptions());
                            BoundingBoxUV bb = lineSolid.Faces.get_Item(0).GetBoundingBox();
                            double maxU = bb.Max.U;
                            double maxV = bb.Max.V;
                            double minU = bb.Min.U;
                            double minV = bb.Min.V;
                            double length = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxV - minV + openingBoundOffset),UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                            double width = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxU - minU + openingBoundOffset), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                            FamilyInstance el = null;


                        if (intersectionCurves.Count() > 0 )

                            {
                                
                                        Line intersectionLine = intersectionCurves.First() as Line;
                                        double x = (intersectionLine.GetEndPoint(0).X + intersectionLine.GetEndPoint(1).X) / 2;
                                        double y = (intersectionLine.GetEndPoint(0).Y + intersectionLine.GetEndPoint(1).Y) / 2;
                                        double z = (intersectionLine.GetEndPoint(0).Z + intersectionLine.GetEndPoint(1).Z) / 2;
                                        XYZ locationPoint = new XYZ(x, y, z);
                                        XYZ outlineFirstPoint =  new XYZ(locationPoint.X - 0.0001 , locationPoint.Y - 0.0001, locationPoint.Z - 0.0001);
                                        XYZ outlineSecondPoint = new XYZ(locationPoint.X + 0.0001, locationPoint.Y + 0.0001, locationPoint.Z + 0.0001);

                                        Outline outline = new Outline(outlineFirstPoint, outlineSecondPoint);
                                        BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
                                        List<FamilyInstance> intersectedElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                                                                                        .OfCategory(BuiltInCategory.OST_DataDevices)
                                                                                        .WhereElementIsNotElementType()
                                                                                        .WherePasses(filter)
                                                                                        .Cast<FamilyInstance>()
                                                                                        .Where(ele => ele.Name.Contains("Отверстие"))
                                                                                        .ToList();

                            if (intersectedElements.Count == 0)
                            {
                                if (linkElement.Category.Name == "Стены")
                                {

                                    using (Transaction t = new Transaction(doc, "Генерация отверстия"))
                                    {
                                        t.Start();
                                        el = doc.Create.NewFamilyInstance(locationPoint, wallOpeningFamilySymbol[0], line.Direction, null, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                        el.LookupParameter("ADSK_Отверстие_Функция").Set(openingFunc);
                                        t.Commit();
                                    }

                                    Wall twall = linkElement as Wall;
                                    if (twall.Orientation.Y == 1 | twall.Orientation.Y == -1)
                                    {
                                        using (Transaction t = new Transaction(doc, "Образмеривание"))
                                        {
                                            t.Start();
                                            el.LookupParameter("ADSK_Размер_Длина").Set(width);
                                            el.LookupParameter("ADSK_Размер_Ширина").Set(length);
                                            el.LookupParameter("ADSK_Размер_Толщина").Set(intersectionLine.ApproximateLength);
                                            t.Commit();
                                        }

                                        CombineOpenings(el, wallOpeningFamilySymbol[0], line,  "first", openingFunc);

                                    }

                                    else
                                    {
                                        using (Transaction t = new Transaction(doc, "Образмеривание"))
                                        {
                                            t.Start();
                                            el.LookupParameter("ADSK_Размер_Длина").Set(length);
                                            el.LookupParameter("ADSK_Размер_Ширина").Set(width);
                                            el.LookupParameter("ADSK_Размер_Толщина").Set(intersectionLine.ApproximateLength);
                                            t.Commit();
                                        }

                                        CombineOpenings(el, wallOpeningFamilySymbol[0], line, "second", openingFunc);

                                    }
                                }

                                else
                                {
                                    using (Transaction t = new Transaction(doc, "Образмеривание"))
                                    {
                                        t.Start();
                                        el = doc.Create.NewFamilyInstance(locationPoint, floorOpeningFamilySymbol[0], line.Direction, null, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                        el.LookupParameter("ADSK_Отверстие_Функция").Set(openingFunc);
                                        el.LookupParameter("ADSK_Размер_Длина").Set(width);
                                        el.LookupParameter("ADSK_Размер_Ширина").Set(length);
                                        el.LookupParameter("ADSK_Размер_Толщина").Set(intersectionLine.ApproximateLength);
                                        t.Commit();
                                    }

                                    CombineOpenings(el, floorOpeningFamilySymbol[0], line,  null, openingFunc);

                                }
                            }
                                
                        }
                       
                   }
                    
                }
                
            }
            return Result.Succeeded;
        }
        private static Solid GetSolidOfElement(Element element)
        {
            Solid elementSolid = null;
            Options geomOptions = new Options();
            geomOptions.ComputeReferences = true;
            GeometryElement geoElement = element.get_Geometry(geomOptions);
            foreach (GeometryObject geoObject in geoElement)
            {
               {
                 elementSolid = geoObject as Solid;
               }
              
            }

            return elementSolid;
        }
        private static Line GetCenterLineOfElement(Element element)
        {
            Line elementLine = null;
            LocationCurve locationCurve =  element.Location as LocationCurve;
            if (locationCurve != null)
            {
                Curve curve = locationCurve.Curve;
                elementLine = curve as Line;
            }
           
            return elementLine;
        }
        private static void CombineOpenings(FamilyInstance CurrentOpening, FamilySymbol fS, Line line, string orientation, string openingFunc)
        {
            List<FamilyInstance> openingInstances = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfCategory(BuiltInCategory.OST_DataDevices)
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(ele => ele.Name.Contains("Отверстие"))
                        .ToList();
            if ((openingInstances.Count > 1) && (openingInstances != null))
            {
                Outline outline = new Outline(CurrentOpening.get_BoundingBox(null).Min, CurrentOpening.get_BoundingBox(null).Max);
                BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
                List<FamilyInstance> intersectedElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_DataDevices)
                    .WhereElementIsNotElementType()
                    .WherePasses(filter)
                    .Cast<FamilyInstance>()
                    .Where(ele => ele.Name.Contains("Отверстие"))
                    .ToList();
                if ((intersectedElements.Count > 1) && (intersectedElements != null))
                {
                    /*List<double> x = new List<double>();
                    List<double> y = new List<double>();
                    List<double> z = new List<double>();*/
                    double maxX = -100000;
                    double maxY = -100000;
                    double maxZ = -100000;
                    double minX = 100000;
                    double minY = 100000;
                    double minZ = 100000;
                    double length = 0;
                    double width = 0;
                    double thick = 0;

                    foreach (FamilyInstance familyInstance in intersectedElements)
                    {

                        if (familyInstance.get_BoundingBox(null).Max.X > maxX) { maxX = familyInstance.get_BoundingBox(null).Max.X; }
                        if (familyInstance.get_BoundingBox(null).Max.Y > maxY) { maxY = familyInstance.get_BoundingBox(null).Max.Y; }
                        if (familyInstance.get_BoundingBox(null).Max.Z > maxZ) { maxZ = familyInstance.get_BoundingBox(null).Max.Z; }
                        if (familyInstance.get_BoundingBox(null).Min.X < minX) { minX = familyInstance.get_BoundingBox(null).Min.X; }
                        if (familyInstance.get_BoundingBox(null).Min.Y < minY) { minY = familyInstance.get_BoundingBox(null).Min.Y; }
                        if (familyInstance.get_BoundingBox(null).Min.Z < minZ) { minZ = familyInstance.get_BoundingBox(null).Min.Z; }

                    }
                    if (orientation == "first")
                    {
                        
                        length = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxZ - minZ), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                        width = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxX - minX), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                        thick = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxY - minY), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                    }
                   
                    else if (orientation == "second")
                    {
                        length = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxZ - minZ), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                        width = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxY - minY), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                        thick = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxX - minX), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                    }
                    else
                    {
                        length = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxX - minX), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                        width = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxY - minY), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                        thick = UnitUtils.ConvertToInternalUnits(Math.Round(UnitUtils.ConvertFromInternalUnits((maxZ - minZ), UnitTypeId.Millimeters) / 5) * 5, UnitTypeId.Millimeters);
                    }
                    XYZ intersectedPoint = new XYZ((maxX+minX)/2, (maxY + minY) / 2, (maxZ + minZ) / 2);


                    using (Transaction t = new Transaction(doc, "Вставка объединенного отверстия"))
                    {
                        t.Start();
                        FamilyInstance newElement = doc.Create.NewFamilyInstance(intersectedPoint, fS, line.Direction, null, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        newElement.LookupParameter("ADSK_Размер_Длина").Set(length);
                        newElement.LookupParameter("ADSK_Размер_Ширина").Set(width);
                        newElement.LookupParameter("ADSK_Размер_Толщина").Set(thick);
                        newElement.LookupParameter("ADSK_Отверстие_Функция").Set(openingFunc);
                        foreach (FamilyInstance familyInstance in intersectedElements)
                        {
                            doc.Delete(familyInstance.Id);
                        }

                        t.Commit();
                    }
                }

            }
        }
    }
}
