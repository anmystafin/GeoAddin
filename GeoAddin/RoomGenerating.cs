#region Namespaces
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

#endregion

namespace GeoAddin
{
    [Transaction(TransactionMode.Manual)]
    public class RoomGenerating : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            RoomGenWindow win = new RoomGenWindow();
            win.ShowDialog();
            bool clickedon = win.clickedon;
            bool clickedoff = win.clickedoff;
            while (clickedon == false & clickedoff == false)
            {
                continue;
            }
            if (clickedon == true)
            {
                double upoffset = win.upoffset;
                double botoffset = win.botoffset;
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                RevitApplication app = uiapp.Application;
                Document doc = uidoc.Document;

                IList<Element> roomsToRemove = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToList();
                if (roomsToRemove.Count > 0)
                {
                    using (Transaction t = new Transaction(doc, "Удаление существующих помещений"))
                    {
                        t.Start();
                        foreach (Room room in roomsToRemove)
                        {
                            doc.Delete(room.Id);
                        }
                        t.Commit();
                    }
                }

                List<Room> rooms = new List<Room>();
                List<Phase> phases = new List<Phase>();
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Генерация помещений");

                    double roomUp = upoffset;
                    double roomDown = botoffset;
                    foreach (Phase phase in doc.Phases)
                    {
                        phases.Add(phase);
                    }

                    Level level = doc.ActiveView.GenLevel;

                    foreach (ElementId roomId in doc.Create.NewRooms2(level))
                    {
                        rooms.Add(doc.GetElement(roomId) as Room);
                    }

                    t.Commit();

                    t.Start("Смена тега помещений");
                    FilteredElementCollector roomtags = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsNotElementType();
                    foreach (Room room in rooms)
                    {
                        foreach (RoomTag roomTag in roomtags)
                        {
                            if (room.Id.IntegerValue == roomTag.Room.Id.IntegerValue)
                            {
                                roomTag.ChangeTypeId(new ElementId(159738));
                            }
                        }
                        room.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(UnitUtils.ConvertToInternalUnits(roomUp, UnitTypeId.Millimeters));
                        room.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET).Set(UnitUtils.ConvertToInternalUnits(roomDown, UnitTypeId.Millimeters));
                    }
                    t.Commit();
                }

                FilteredElementCollector plumbingFixtures = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType();

                List<FamilyInstance> kitchenPlumbs = new List<FamilyInstance>();
                List<FamilyInstance> toiletBowls = new List<FamilyInstance>();
                List<FamilyInstance> sinks = new List<FamilyInstance>();
                List<FamilyInstance> washers = new List<FamilyInstance>();
                List<FamilyInstance> baths = new List<FamilyInstance>();

                foreach (FamilyInstance fixture in plumbingFixtures)
                {
                    string fixtureName = fixture.Symbol.Name;
                    if (fixtureName.Contains("Варочная_панель"))
                    {
                        kitchenPlumbs.Add(fixture);
                    }
                    else if (fixtureName.Contains("Унитаз"))
                    {
                        toiletBowls.Add(fixture);
                    }
                    else if (fixtureName.Contains("Умывальник"))
                    {
                        sinks.Add(fixture);
                    }
                    else if (fixtureName.Contains("Стиральная_машина"))
                    {
                        washers.Add(fixture);
                    }
                    else if (fixtureName.Contains("Ванная") || fixtureName.Contains("ДушевойПоддон"))
                    {
                        baths.Add(fixture);
                    }
                }

                FilteredElementCollector windows = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType();
                FilteredElementCollector doors = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType();


               

                using (Transaction t = new Transaction(doc, "Определение помещений"))
                {
                    t.Start();
                    foreach (Room room in rooms)
                    {
                        List<string> roomFixtures = new List<string>();
                        foreach (FamilyInstance fixture in kitchenPlumbs)
                        {
                            XYZ located = (fixture.Location as LocationPoint).Point;
                            if (room.IsPointInRoom(located))
                            {
                                roomFixtures.Add("Кухня");
                            }
                        }
                        foreach (FamilyInstance fixture in toiletBowls)
                        {
                            if (room.IsPointInRoom((fixture.Location as LocationPoint).Point))
                            {
                                roomFixtures.Add("Унитаз");
                            }
                        }
                        foreach (FamilyInstance fixture in sinks)
                        {
                            if (room.IsPointInRoom((fixture.Location as LocationPoint).Point))
                            {
                                roomFixtures.Add("Умывальник");
                            }
                        }
                        foreach (FamilyInstance fixture in washers)
                        {
                            if (room.IsPointInRoom((fixture.Location as LocationPoint).Point))
                            {
                                roomFixtures.Add("Стиральная машина");
                            }
                        }
                        foreach (FamilyInstance fixture in baths)
                        {
                            if (room.IsPointInRoom((fixture.Location as LocationPoint).Point))
                            {
                                roomFixtures.Add("Ванна");
                            }
                        }

                        if (roomFixtures.Contains("Унитаз") && roomFixtures.Contains("Ванна"))
                        {
                            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set("С.У.");
                        }
                        else if (!roomFixtures.Contains("Ванна") && roomFixtures.Contains("Умывальник") && roomFixtures.Contains("Унитаз"))
                        {
                            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set("Уборная");
                        }
                        else if (!roomFixtures.Contains("Унитаз") && roomFixtures.Contains("Ванна"))
                        {
                            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set("Ванная");
                        }
                        else if (!roomFixtures.Contains("Ванна") && !roomFixtures.Contains("Умывальник") && !roomFixtures.Contains("Унитаз") && roomFixtures.Contains("Стиральная машина"))
                        {
                            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set("Постирочная");
                        }
                        else if (roomFixtures.Contains("Кухня"))
                        {
                            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set("Кухня");
                        }

                        List<int> roomWindowsIds = new List<int>();
                        foreach (FamilyInstance window in windows)
                        {
                            Room windowToRoom = window.get_ToRoom(phases[phases.Count - 1]);
                            if (windowToRoom != null)
                            {
                                roomWindowsIds.Add(windowToRoom.Id.IntegerValue);
                            }
                        }
                        foreach (int elementId in roomWindowsIds.Distinct())
                        {
                            if (room.Id.IntegerValue == elementId && room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString().Contains("Помещение"))
                            {
                                room.get_Parameter(BuiltInParameter.ROOM_NAME).Set("Жилая комната");
                            }
                        }
                        if (room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() == "Помещение")
                        {
                            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set("Коридор");
                        }
                        
                    }
                    foreach (FamilyInstance door in doors)
                    {
                        Room doorFromRoom = door.get_FromRoom(phases[phases.Count - 1]);
                        if (doorFromRoom != null && door.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsString().Contains("Балконная"))
                        {
                            doorFromRoom.get_Parameter(BuiltInParameter.ROOM_NAME).Set("Лоджия");
                         
                        }
                    }
                    t.Commit();
                }
                
            }
            Debug.Print("Complited the task.");
            return Result.Succeeded;
        }
    }
}
