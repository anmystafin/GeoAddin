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
    public class WindowsFilling : IExternalCommand
    {
        //Основные переменные
        static UIApplication uiapp;
        static UIDocument uidoc;
        static RevitApplication app;
        static Document doc;

        //Переменные для работы с параметрами
        string windowconstrType;
        string windowmaterial;
        string windowhigth;
        string windowlength;
        string windowwidth;
        string windowOpenType;
        string windowGost;
        string windowName;

        //Списки возможных параметров для генерации наименования
        List<string> constrTypes = new List<string>()
        {
            "О", "ОСш", "ОСвз", "ОСвп", "Б", "Бф", "ОБЛ", "ОБП", "ОБр", "ОБ"
        };
        List<string> materials = new List<string>()
        {
            "Д", "А", "П", "Ст", "Спл", "ДА", "ДАН", "ПА", "АД"
        };
        List<string> opentypes = new List<string>()
        {
            "ПР", "ОТ", "ПОТ", "ОТП", "ПВ", "СП", "ВП", "Рз", "П", "Ск", "Н", "К", "ГО"
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение всех окон в проекте и проверка параметров
            IList<Element> windowTypes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsElementType().ToList();
            foreach (Element window in windowTypes)
            {
                windowconstrType = window.LookupParameter("<Тип_Оконной_Конструкции>").AsString();
                windowmaterial = window.LookupParameter("<Материал_Профиля>").AsString();
                windowhigth = window.LookupParameter("<Высота>").AsDouble().ToString();
                windowlength = window.LookupParameter("<Длина>").AsDouble().ToString();
                windowwidth = window.LookupParameter("<Ширина>").AsDouble().ToString();
                windowOpenType = window.LookupParameter("<Тип_Открывания>").AsString();
                windowGost = window.LookupParameter("<ГОСТ>").AsString();

            }
            if (constrTypes.Contains(windowconstrType)) { } else { windowconstrType = "Неверно указан тип оконной конструкции"; }

            if (materials.Contains(windowmaterial)) { } else { windowmaterial = "Неверно указан тип материала"; }

            if (opentypes.Contains(windowOpenType)) { } else { windowOpenType = "Неверно указан тип открывания"; }

            if ((windowconstrType != "Неверно указан тип оконной конструкции") && (windowmaterial != "Неверно указан тип материала") && (windowOpenType != "Неверно указан тип открывания"))
            {
                windowName = windowconstrType + " " + windowmaterial + " " + windowhigth + "x" + windowlength + windowOpenType + "-" + windowGost;
            }
            else { windowName = "Данные указаны неверно"; }


        return Result.Succeeded;
        }
    }
}
