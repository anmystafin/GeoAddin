#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using RevitApplication = Autodesk.Revit.ApplicationServices.Application;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Configuration;
using System.Reflection;
using System.Collections.Specialized;
using System.Linq;
using System.Xml.Linq;
using System.Windows;
#endregion

namespace GeoAddin
{
    [Transaction(TransactionMode.Manual)]
    public class ApartmentNumFilling: IExternalCommand
    {
        static UIApplication uiapp;
        static UIDocument uidoc;
        static RevitApplication app;
        static Document doc;
        public ElementId schemid = new ElementId(15969);
        public ElementId catid = new ElementId(BuiltInCategory.OST_Rooms);

        Phase phase;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;

            phase = doc.Phases.get_Item(doc.Phases.Size - 1);
            IList<FamilyInstance> entryDoors = new FilteredElementCollector(doc, doc.ActiveView.Id). // Находим входные двери квартиры
                OfCategory(BuiltInCategory.OST_Doors).
                OfClass(typeof(FamilyInstance)).
                Cast<FamilyInstance>().
                Where(door => door.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsString().Contains("Дверь.Квартирная")). // Здесь и применяется тяжелый фильтр для поиска входных дверей
                ToList();
            FilteredElementCollector allRooms = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();

            using (Transaction t = new Transaction(doc, "Нумерация комнат квартир")) // Основная транзакция
            {
                t.Start();
                foreach (FamilyInstance entryDoor in entryDoors) // Проходим по каждой входной двери
                {
                    List<Room> apartmnetRooms = GetApartmentRooms(entryDoor.get_FromRoom(phase), allRooms, null, entryDoor); // Эта функция отвечает за нахождение всех комнат в картире
                    Level lvl = doc.GetElement(entryDoor.LevelId) as Level; // Часть, просто отвечающая за взятие номера квартиры, у нас по форме L01_001 с указанием уровня и номера квартиры
                    string doorNumber = entryDoor.LookupParameter("ADSK_Номер квартиры").AsString();
                    string apartmentNumber = null;
                    if (!doorNumber.Contains("L") && !doorNumber.Contains("_"))
                    {
                        apartmentNumber = $"L{lvl.Name.Replace("Этаж ", "")}_{LeadingZeros(entryDoor.LookupParameter("ADSK_Номер квартиры").AsString())}";
                    }
                    else
                    {
                        apartmentNumber = doorNumber;
                    }
                    foreach (Room room in apartmnetRooms)
                    {
                        try
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(apartmentNumber);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Ошибка");
                        }
                    }
                    entryDoor.LookupParameter("ADSK_Номер квартиры").Set(apartmentNumber);
                }
                t.Commit();
                using (Transaction tx = new Transaction(doc))
                {
                    try
                    {
                        tx.Start("Transaction Name");
                        doc.ActiveView.SetColorFillSchemeId(catid, schemid);
                        tx.Commit();
                    }
                    catch { }
                    
                }
            }
            return Result.Succeeded;
        }
        /*
         * Эта чудо-функция находит все комнаты. Объясню как работает на следующей неделе, потому что текстом я не знаю как объяснить.
         */
        private List<Room> GetApartmentRooms(Room currentRoom, FilteredElementCollector allRooms, List<Room> apartmentRooms = null, FamilyInstance entryDoor = null)
        {
            if (apartmentRooms == null)
            {
                apartmentRooms = new List<Room>();
            }
            apartmentRooms.Add(currentRoom);
            List<int> roomsIds = GetIdsOfRoom(apartmentRooms);

            List<FamilyInstance> allDoorsOfRoom = new FilteredElementCollector(doc, doc.ActiveView.Id).
                OfCategory(BuiltInCategory.OST_Doors).
                OfClass(typeof(FamilyInstance)).
                Cast<FamilyInstance>().
                Where(door =>
                door.get_FromRoom(phase).Id.IntegerValue == currentRoom.Id.IntegerValue || door.get_ToRoom(phase).Id.IntegerValue == currentRoom.Id.IntegerValue).
                ToList();
            if (entryDoor != null)
            {
                allDoorsOfRoom = allDoorsOfRoom.Where(door => door.Id.IntegerValue != entryDoor.Id.IntegerValue).ToList();
            }

            if (allDoorsOfRoom.Count > 0)
            {
                foreach (FamilyInstance door in allDoorsOfRoom)
                {
                    if (door.get_FromRoom(phase).Id.IntegerValue != currentRoom.Id.IntegerValue)
                    {
                        if (!roomsIds.Contains(door.get_FromRoom(phase).Id.IntegerValue))
                        {
                            apartmentRooms = GetApartmentRooms(door.get_FromRoom(phase), allRooms, apartmentRooms, door);
                        }
                    }
                    else
                    {
                        if (!roomsIds.Contains(door.get_ToRoom(phase).Id.IntegerValue))
                        {
                            apartmentRooms = GetApartmentRooms(door.get_ToRoom(phase), allRooms, apartmentRooms, door);
                        }
                    }
                }
            }
            Solid currentRoomSolid = GetSolidOfRoom(currentRoom);
            List<Room> adjoiningRooms = allRooms.Cast<Room>().
                Where(room => SolidsAreToching(currentRoomSolid, GetSolidOfRoom(room)) && !roomsIds.Contains(room.Id.IntegerValue)).
                ToList();
            if (adjoiningRooms.Count > 0)
            {
                foreach (Room room in adjoiningRooms)
                {
                    apartmentRooms = GetApartmentRooms(room, allRooms, apartmentRooms);
                }
            }

            return apartmentRooms;
        }
        private static bool SolidsAreToching(Solid solid1, Solid solid2)
        {
            Solid interSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Intersect);
            Solid unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Union);

            double sumArea = Math.Round(Math.Abs(solid1.SurfaceArea + solid2.SurfaceArea), 5);
            double sumFaces = Math.Abs(solid1.Faces.Size + solid2.Faces.Size);
            double unionArea = Math.Round(Math.Abs(unionSolid.SurfaceArea), 5);
            double unionFaces = Math.Abs(unionSolid.Faces.Size);

            if (sumArea > unionArea && sumFaces > unionFaces && interSolid.Volume < 0.00001)
            {
                return true;
            }
            return false;
        }
        private static Solid GetSolidOfRoom(Room room)
        {
            SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(doc);
            SpatialElementGeometryResults results = calculator.CalculateSpatialElementGeometry(room);
            Solid roomSolid = results.GetGeometry();

            return roomSolid;
        }
        private static List<int> GetIdsOfRoom(List<Room> rooms)
        {
            List<int> ids = new List<int>();
            foreach (Room room in rooms)
            {
                ids.Add(room.Id.IntegerValue);
            }
            return ids;
        }
        private static string LeadingZeros(string str)
        {
            if (str.Length == 1)
            {
                return "00" + str;
            }
            if (str.Length == 2)
            {
                return "0" + str;
            }
            if (str.Length == 3)
            {
                return str;
            }
            return null;


        }



    }
}
