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
                    if (room.LookupParameter("ADSK_����� ��������").AsString() != null)
                    {
                        if (!apartments.ContainsKey(room.LookupParameter("ADSK_����� ��������").AsString()))
                        {
                            List<Room> rooms = new List<Room>() { room };
                            apartments.Add(room.LookupParameter("ADSK_����� ��������").AsString(), rooms);
                        }
                        else
                        {
                            apartments[room.LookupParameter("ADSK_����� ��������").AsString()].Add(room);
                        }
                        
                    }
                }
                using (Transaction t = new Transaction(doc, "��������������"))
                {
                    t.Start();
                    foreach (string num in apartments.Keys)
                    {
                        int numberOfLivingRooms = 0;
                        double apartmentAreaLivingRooms = 0;            // ADSK_������� �������� �����
                        double apartmaneAreaWithoutSummerRooms = 0;     // ADSK_������� �������� � ��
                        double apartmentAreaGeneral = 0;                // ADSK_������� �������� �����
                        double apartmentAreaGeneralWithoutCoef = 0;     // TRGR_������� �������� ��� ��
                        foreach (Room room in apartments[num])
                        {
                            double areaOfRoom = Math.Round(UnitUtils.ConvertFromInternalUnits(room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters), roundNum);
                            try
                            {
                                room.LookupParameter("������� ���������").Set(UnitUtils.ConvertToInternalUnits(Math.Round(areaOfRoom, roundNum), UnitTypeId.SquareMeters));
                            }
                            catch (Exception ex)
                            {
                                Debug.Print(ex.ToString());
                            }
                            try
                            {
                                double coefficent = 1.0;


                               
                                if (room.LookupParameter("���").AsString() != "������" && room.LookupParameter("���").AsString() != "������")
                                    
                                {
                                    apartmaneAreaWithoutSummerRooms += areaOfRoom;
                                    apartmentAreaGeneral += areaOfRoom;
                                    room.LookupParameter("ADSK_����������� �������").Set(coefficent);
                                }
                                if (room.LookupParameter("���").AsString() == "������" || room.LookupParameter("���").AsString() == "������")
                                {
                                    if (room.LookupParameter("���").AsString() == "������")
                                    {
                                        room.LookupParameter("ADSK_����������� �������").Set(loggieAreaCoef);
                                    }
                                    else if (room.LookupParameter("���").AsString() == "������")
                                    {
                                        room.LookupParameter("ADSK_����������� �������").Set(balconyAreaCoef);
                                    }
                                    apartmentAreaGeneral += Math.Round(areaOfRoom * room.LookupParameter("ADSK_����������� �������").AsDouble(), roundNum);
                                }
                                if (room.LookupParameter("���").AsString() == "����� �������" || room.LookupParameter("���").AsString() == "��������" || room.LookupParameter("���").AsString() == "�������")
                                {
                                    numberOfLivingRooms++;
                                    apartmentAreaLivingRooms += areaOfRoom;
                                }
                                room.LookupParameter("ADSK_������� � �������������").Set(UnitUtils.ConvertToInternalUnits(Math.Round(areaOfRoom * room.LookupParameter("ADSK_����������� �������").AsDouble(), roundNum), UnitTypeId.SquareMeters));
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
                                room.LookupParameter("ADSK_���������� ������").Set(numberOfLivingRooms);
                                room.LookupParameter("ADSK_������� ��������").Set(apartmaneAreaWithoutSummerRooms);
                                room.LookupParameter("ADSK_������� �������� �����").Set(apartmentAreaLivingRooms);
                                room.LookupParameter("ADSK_������� �������� �����").Set(apartmentAreaGeneral);
                                room.LookupParameter("������� �������� ��� ��").Set(apartmentAreaGeneralWithoutCoef);
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
