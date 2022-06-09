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
    public class Apartmentgraphy : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            RevitApplication app = uiapp.Application;
            Document doc = uidoc.Document;

            ApartmentgraphyWindow win = new ApartmentgraphyWindow();
            win.ShowDialog();
            bool clickedon = win.clickedon;
            bool clickedoff = win.clickedoff;
            while (clickedon == false & clickedoff == false)
            {
                continue;
            }
            if (clickedon == true)
            {
                FilteredElementCollector areas = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();
                Dictionary<string, List<Room>> apartments = new Dictionary<string, List<Room>>();


                int roundNum = win.roundNum;
                double loggieAreaCoef = win.loggiacoef;
                double balconyAreaCoef = win.balconycoef;

                foreach (Room room in areas)
                {
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString() != null)
                    {
                        if (!apartments.ContainsKey(room.LookupParameter("ADSK_Номер квартиры").AsString()))
                        {
                            List<Room> rooms = new List<Room>() { room };
                            apartments.Add(room.LookupParameter("ADSK_Номер квартиры").AsString(), rooms);
                        }
                        else
                        {
                            apartments[room.LookupParameter("ADSK_Номер квартиры").AsString()].Add(room);
                        }
                        
                    }
                }
                using (Transaction t = new Transaction(doc, "Квартирография"))
                {
                    t.Start();
                    foreach (string num in apartments.Keys)
                    {
                        int numberOfLivingRooms = 0;
                        double apartmentAreaLivingRooms = 0;            // ADSK_Площадь квартиры жилая
                        double apartmaneAreaWithoutSummerRooms = 0;     // ADSK_Площадь квартиры с кф
                        double apartmentAreaGeneral = 0;                // ADSK_Площадь квартиры общая
                        double apartmentAreaGeneralWithoutCoef = 0;     // TRGR_Площадь квартиры без кф
                        foreach (Room room in apartments[num])
                        {
                            double areaOfRoom = Math.Round(UnitUtils.ConvertFromInternalUnits(room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters), roundNum);
                            try
                            {
                                room.LookupParameter("Площадь помещения").Set(UnitUtils.ConvertToInternalUnits(Math.Round(areaOfRoom, roundNum), UnitTypeId.SquareMeters));
                            }
                            catch (Exception ex)
                            {
                                Debug.Print(ex.ToString());
                            }
                            try
                            {
                                double coefficent = 1.0;


                               
                                if (room.LookupParameter("Имя").AsString() != "Лоджия" && room.LookupParameter("Имя").AsString() != "Балкон")
                                    
                                {
                                    apartmaneAreaWithoutSummerRooms += areaOfRoom;
                                    apartmentAreaGeneral += areaOfRoom;
                                    room.LookupParameter("ADSK_Коэффициент площади").Set(coefficent);
                                }
                                if (room.LookupParameter("Имя").AsString() == "Лоджия" || room.LookupParameter("Имя").AsString() == "Балкон")
                                {
                                    if (room.LookupParameter("Имя").AsString() == "Лоджия")
                                    {
                                        room.LookupParameter("ADSK_Коэффициент площади").Set(loggieAreaCoef);
                                    }
                                    else if (room.LookupParameter("Имя").AsString() == "Балкон")
                                    {
                                        room.LookupParameter("ADSK_Коэффициент площади").Set(balconyAreaCoef);
                                    }
                                    apartmentAreaGeneral += Math.Round(areaOfRoom * room.LookupParameter("ADSK_Коэффициент площади").AsDouble(), roundNum);
                                }
                                if (room.LookupParameter("Имя").AsString() == "Жилая комната" || room.LookupParameter("Имя").AsString() == "Гостиная" || room.LookupParameter("Имя").AsString() == "Спальня")
                                {
                                    numberOfLivingRooms++;
                                    apartmentAreaLivingRooms += areaOfRoom;
                                }
                                room.LookupParameter("ADSK_Площадь с коэффициентом").Set(UnitUtils.ConvertToInternalUnits(Math.Round(areaOfRoom * room.LookupParameter("ADSK_Коэффициент площади").AsDouble(), roundNum), UnitTypeId.SquareMeters));
                                apartmentAreaGeneralWithoutCoef += areaOfRoom;
                            }
                            catch (Exception ex)
                            {
                                Debug.Print(ex.ToString());
                            }
                        }

                        apartmaneAreaWithoutSummerRooms = UnitUtils.ConvertToInternalUnits(Math.Round(apartmaneAreaWithoutSummerRooms, roundNum), UnitTypeId.SquareMeters);
                        apartmentAreaLivingRooms = UnitUtils.ConvertToInternalUnits(Math.Round(apartmentAreaLivingRooms, roundNum), UnitTypeId.SquareMeters);
                        apartmentAreaGeneral = UnitUtils.ConvertToInternalUnits(Math.Round(apartmentAreaGeneral, roundNum), UnitTypeId.SquareMeters);
                        apartmentAreaGeneralWithoutCoef = UnitUtils.ConvertToInternalUnits(Math.Round(apartmentAreaGeneralWithoutCoef, roundNum), UnitTypeId.SquareMeters);
                        foreach (Room room in apartments[num])
                        {
                            try
                            {
                                room.LookupParameter("ADSK_Количество комнат").Set(numberOfLivingRooms);
                                room.LookupParameter("ADSK_Площадь квартиры").Set(apartmaneAreaWithoutSummerRooms);
                                room.LookupParameter("ADSK_Площадь квартиры жилая").Set(apartmentAreaLivingRooms);
                                room.LookupParameter("ADSK_Площадь квартиры общая").Set(apartmentAreaGeneral);
                                room.LookupParameter("Площадь квартиры без кф").Set(apartmentAreaGeneralWithoutCoef);
                            }
                            catch (Exception ex)
                            {
                                Debug.Print(ex.ToString());
                            }
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
