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
    public class WindowsSchema : IExternalCommand
    {
        //Основные переменные

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
            IList<View> views = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().Cast<View>().ToList();
            uiapp.ActiveUIDocument.ActiveView = GetWindowSchemaView(views);
            //Получение всех окон и видов в проекте, получение типоразмеров окон, используемых в проекте
            List<FamilyInstance> windowInstances = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
            List<string> windowTypesNames = new HashSet<string>(windowInstances.Select(el => el.Symbol.Name)).ToList();
            List<FamilySymbol> allWindowTypes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsElementType().Cast<FamilySymbol>().ToList();
            List<FamilySymbol> activeWindowTypes = new List<FamilySymbol>();
            foreach (FamilySymbol windowType in allWindowTypes)
            {
                if ((windowType.IsActive) && (windowType.FamilyName.Contains("Окно"))) { activeWindowTypes.Add(windowType); }

            }
            List<FamilySymbol> windowTypes = new List<FamilySymbol>();
            foreach (FamilySymbol familySymbol in activeWindowTypes)
            {
                foreach (string windowTypesName in windowTypesNames)
                {
                    if (familySymbol.Name == windowTypesName) { windowTypes.Add(familySymbol);}
                }
            } 

            //Получение легенду на виде и типизированный список его Id
            List<Element> legendComponent = new FilteredElementCollector(doc, GetWindowSchemaView(views).Id).OfCategory(BuiltInCategory.OST_LegendComponents).ToList();
            ICollection<ElementId> elementIds = legendComponent.Select(el => el.Id).ToList();

            double deltaX = 7;
            
            foreach (FamilySymbol familySymbol in activeWindowTypes)
            {

                using (Transaction t = new Transaction(doc, "Копирование компонентов легенды"))
                {
                    t.Start();
                    //Group group = doc.Create.NewGroup(uidoc.Selection.GetElementIds());
                    Group group = doc.Create.NewGroup(elementIds);
                    LocationPoint location = group.Location as LocationPoint;
                    XYZ windowLocationPoint = new XYZ(location.Point.X + deltaX, location.Point.Y, location.Point.Z);
                    XYZ markLocationPoint = new XYZ(location.Point.X + deltaX + 0.5, location.Point.Y/133.5, 0);
                    XYZ squareLocationPoint = new XYZ(location.Point.X + deltaX + 0.5, location.Point.Y / 135, 0);
                    Group newGroup = doc.Create.PlaceGroup(windowLocationPoint, group.GroupType);
                    group.UngroupMembers();
                    newGroup.UngroupMembers();
                    TextNote markTextNote = TextNote.Create(doc, GetWindowSchemaView(views).Id, markLocationPoint, familySymbol.LookupParameter("ADSK_Марка").AsString(), new TextNoteOptions(new ElementId(8047)));
                    string windowSquare = "S = " + ((UnitUtils.ConvertFromInternalUnits(familySymbol.LookupParameter("Примерная ширина").AsDouble(), UnitTypeId.Millimeters) * UnitUtils.ConvertFromInternalUnits(familySymbol.LookupParameter("Примерная высота").AsDouble(), UnitTypeId.Millimeters))/1000000).ToString() + " м2";    
                    TextNote squareTextNote = TextNote.Create(doc, GetWindowSchemaView(views).Id, squareLocationPoint, windowSquare, new TextNoteOptions(new ElementId(8047)));
                    t.Commit();

                }
                
                deltaX += 7;
            }
            using (Transaction t = new Transaction(doc, "Удаление лишней схемы"))
            {
                t.Start();
                doc.Delete(legendComponent[0].Id);
                t.Commit();
            }
           
            using (Transaction t = new Transaction(doc, "Сопоставление типов окон"))
            {
                t.Start();
                List<Element> newLegendComponents = new FilteredElementCollector(doc, GetWindowSchemaView(views).Id).OfCategory(BuiltInCategory.OST_LegendComponents).ToList();
                for (int i = 0; i < newLegendComponents.Count; i++)
                {
                    newLegendComponents[i].LookupParameter("Тип компонента").Set(activeWindowTypes[i].Id);
                }
                t.Commit();
            }
             
            return Result.Succeeded;
        }
        //Функция для нахождения легенды схем окон
        private View GetWindowSchemaView(IList<View> views)
        { 
            View view = null;
            foreach (View element in views)
            {
                if (element.Name == "Схемы окон") { view = element; }
            }
            return view;    
        }

    }
}
