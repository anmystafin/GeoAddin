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
                            double length = (maxV - minV) + openingBoundOffset;
                            double width = (maxU - minU) + openingBoundOffset;
                            FamilyInstance el = null;
                            if (intersectionCurves.Count() > 0)

                            {
                                using (Transaction t = new Transaction(doc, "Генерация отверстий"))
                                {
                                t.Start();

                                        Line intersectionLine = intersectionCurves.First() as Line;
                                        double x = (intersectionLine.GetEndPoint(0).X + intersectionLine.GetEndPoint(1).X) / 2;
                                        double y = (intersectionLine.GetEndPoint(0).Y + intersectionLine.GetEndPoint(1).Y) / 2;
                                        double z = (intersectionLine.GetEndPoint(0).Z + intersectionLine.GetEndPoint(1).Z) / 2;
                                        XYZ locationPoint = new XYZ(x, y, z);
                                        if (linkElement.Category.Name == "Стены")
                                        {


                                            el = doc.Create.NewFamilyInstance(locationPoint, wallOpeningFamilySymbol[0], line.Direction, null, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                            el.LookupParameter("ADSK_Отверстие_Функция").Set(openingFunc);
                                            Wall twall = linkElement as Wall;
                                            if (twall.Orientation.Y == 1 | twall.Orientation.Y == -1)
                                            {
                                                el.LookupParameter("ADSK_Размер_Длина").Set(width);
                                                el.LookupParameter("ADSK_Размер_Ширина").Set(length);
                                                el.LookupParameter("ADSK_Размер_Толщина").Set(intersectionLine.ApproximateLength);

                                            }

                                            else
                                            {

                                                el.LookupParameter("ADSK_Размер_Длина").Set(length);
                                                el.LookupParameter("ADSK_Размер_Ширина").Set(width);
                                                el.LookupParameter("ADSK_Размер_Толщина").Set(intersectionLine.ApproximateLength);

                                            }
                                        }

                                        else
                                        {
                                            el = doc.Create.NewFamilyInstance(locationPoint, floorOpeningFamilySymbol[0], line.Direction, null, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                            el.LookupParameter("ADSK_Отверстие_Функция").Set(openingFunc);
                                            el.LookupParameter("ADSK_Размер_Длина").Set(width);
                                            el.LookupParameter("ADSK_Размер_Ширина").Set(length);
                                            el.LookupParameter("ADSK_Размер_Толщина").Set(intersectionLine.ApproximateLength);
                                        }
                                t.Commit();
                                }
                                
                            }
                       
                   }
                    
              }
                //Поиск отверстий и объединение
                List<FamilyInstance> openingInstances = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfCategory(BuiltInCategory.OST_DataDevices)
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(ele => ele.Name.Contains("Отверстие"))
                        .ToList();

                foreach (FamilyInstance fi in openingInstances)
                {

                    BoundingBoxXYZ elBB = GetSolidOfElement(fi).GetBoundingBox();
                    XYZ elMaxPoint = elBB.Max;
                    XYZ elMinPoint = elBB.Min;
                    Outline outline = new Outline(elMinPoint, elMaxPoint);
                    BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
                    List<FamilyInstance> intersectingElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                                                    .OfCategory(BuiltInCategory.OST_DataDevices)
                                                    .WhereElementIsNotElementType()
                                                    .WherePasses(filter)
                                                    .Where(ele => ele.Name.Contains("Отверстие"))
                                                    .Cast<FamilyInstance>()
                                                    .ToList();
                    double maxPointU = -1000000;
                    double maxPointV = -1000000;
                    double minPointU = 1000000;
                    double minPointV = 1000000;
                    double maxX = -1000000;
                    double maxY = -1000000;
                    double maxZ = -1000000;
                    double minX = 1000000;
                    double minY = 1000000;
                    double minZ = 1000000;

                    foreach (FamilyInstance familyInstance in intersectingElements)
                    {
                        UV maxPoint = GetSolidOfElement(familyInstance).Faces.get_Item(0).GetBoundingBox().Max;
                        UV minPoint = GetSolidOfElement(familyInstance).Faces.get_Item(0).GetBoundingBox().Min;
                        if (maxPoint.U > maxPointU) { maxPointU = maxPoint.U; }
                        if (maxPoint.V > maxPointV) { maxPointV = maxPoint.V; }
                        if (minPoint.U < minPointU) { minPointU = minPoint.U; }
                        if (minPoint.V < minPointV) { minPointV = minPoint.V; }
                        BoundingBoxXYZ familyInstanceBB = GetSolidOfElement(familyInstance).GetBoundingBox();
                        if (familyInstanceBB.Max.X > maxX) { maxX = familyInstanceBB.Max.X; }
                        if (familyInstanceBB.Max.Y > maxY) { maxY = familyInstanceBB.Max.Y; }
                        if (familyInstanceBB.Max.Z > maxZ) { maxZ = familyInstanceBB.Max.Z; }
                        if (familyInstanceBB.Min.X < minX) { minX = familyInstanceBB.Min.X; }
                        if (familyInstanceBB.Min.Y < minY) { minY = familyInstanceBB.Min.Y; }
                        if (familyInstanceBB.Min.Z < minZ) { minZ = familyInstanceBB.Min.Z; }
                        using (Transaction t = new Transaction(doc, "Удаление старых отверстий"))
                        {
                            t.Start();
                            doc.Delete(familyInstance.Id);
                            t.Commit();
                        }
                        
                    }
                    double intersectedLength = maxPointV - minPointV;
                    double intersectedWidth = maxPointU - minPointU;
                    XYZ intersectedPoint = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);
                    using (Transaction t = new Transaction(doc, "Вставка объединенного отверстия"))
                    {
                        t.Start();
                        FamilyInstance intersectEl = doc.Create.NewFamilyInstance(intersectedPoint, wallOpeningFamilySymbol[0], intersectedPoint.Normalize(), null, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        t.Commit();
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
    }
}
