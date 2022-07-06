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
                foreach (Level item in levels) {if (item.Name == win.LevelBox.Text) {level = item;}}
                foreach (RevitLinkInstance item in linkInstances)
                {
                    foreach (Document document in app.Documents)
                    {
                        if ((item.Name == win.LinkInstance.Text) && (item.Name.Contains(document.Title))) { link = item; linkDoc = document; }
                    }
                     
                }
                string linkDocTitle = linkDoc.Title;
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

                //Проверка пересечения элементов и получение точек вставки семейств
                using (Transaction t = new Transaction(doc, "Генерация отверстий"))
                {
                    t.Start();
                    foreach (Element mepElement in mepElements)
                    {
                        foreach (Element linkElement in linkElements)
                        {
                            Solid solid = GetSolidOfElement(linkElement);
                            Line line = GetCenterLineOfElement(mepElement);
                            SolidCurveIntersection intersectionCurves = solid.IntersectWithCurve(line, new SolidCurveIntersectionOptions());
                            PlanarFace face = (PlanarFace)solid.Faces.get_Item(0);
                            XYZ normalFace = face.FaceNormal;
                            
                            if (intersectionCurves.Count() > 0)
                            {
                                /*var x =  (solid.ToProtoType().BoundingBox.Intersection(solidLine.ToProtoType().BoundingBox).MaxPoint.X + solid.ToProtoType().BoundingBox.Intersection(solidLine.ToProtoType().BoundingBox).MinPoint.X) / 2;
                                double y = (solid.ToProtoType().BoundingBox.Intersection(line.ToProtoType().BoundingBox).MaxPoint.Y + solid.ToProtoType().BoundingBox.Intersection(line.ToProtoType().BoundingBox).MinPoint.Y) / 2;
                                double z = (solid.ToProtoType().BoundingBox.Intersection(line.ToProtoType().BoundingBox).MaxPoint.Z + solid.ToProtoType().BoundingBox.Intersection(line.ToProtoType().BoundingBox).MinPoint.Z) / 2;*/
                                /*XYZ locationPoint = new XYZ(x, y, z);
                                XYZ directionPoint = solid.ToProtoType().Faces.First().SurfaceGeometry().TangentAtUParameter(solid.ToProtoType().Faces.First().SurfaceGeometry().UVParameterAtPoint(locationPoint.ToPoint()).U, solid.ToProtoType().Faces.First().SurfaceGeometry().UVParameterAtPoint(locationPoint.ToPoint()).V).ToXyz();
                                */
                                Line intersectionLine = intersectionCurves.First() as Line;
                                double x = (intersectionLine.GetEndPoint(0).X + intersectionLine.GetEndPoint(1).X) / 2;
                                double y = (intersectionLine.GetEndPoint(0).Y + intersectionLine.GetEndPoint(1).Y) / 2;
                                double z = (intersectionLine.GetEndPoint(0).Z + intersectionLine.GetEndPoint(1).Z) / 2;
                                XYZ locationPoint = new XYZ(x, y, z);
                                Transform transform = Transform.CreateRotationAtPoint(new XYZ(0, 0, 1), (Math.PI * 90) / 180, locationPoint);
                                XYZ transformedNormal = transform.OfVector(normalFace);
                                Line axis = Line.CreateUnbound(locationPoint, transformedNormal);

                                doc.Create.NewFamilyInstance(locationPoint, openingFamilySymbol[0], line.Direction, null, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                List<ElementId> openingIds = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DataDevices).WhereElementIsNotElementType().Select(el => el.Id).ToList();
                                ElementTransformUtils.RotateElements(doc, openingIds, axis, (Math.PI* 90)/180);


                            }
                        }
                    }
                    
                    t.Commit();


                }
                /*using (Transaction t = new Transaction(doc, "Образмеривание отверстий"))
                {
                    t.Start();
                    List<Element> openings = new FilteredElementCollector

                    t.Commit();
                }*/
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
