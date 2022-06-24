#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

#endregion

namespace GeoAddin
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            string tabName = "ООО Георекон";

            string archpanelName = "Архитектура";
            a.CreateRibbonTab(tabName);
            var archpanel = a.CreateRibbonPanel(tabName,archpanelName);

            string commonpanelName = "Общее";
            var commonpanel = a.CreateRibbonPanel(tabName, commonpanelName);

            //Создание кнопки генерациии помещений
            var ApartGenButton = new PushButtonData("Генерация квартир", "Генерация квартир", Assembly.GetExecutingAssembly().Location, "GeoAddin.RoomGenerating");
            var ApartGenPushBtn = archpanel.AddItem(ApartGenButton) as PushButton;
            Image RoomGenPic = Properties.Resources.RoomGenPic;
            ImageSource RoomGenPicSrc = Convert(RoomGenPic);
            ApartGenPushBtn.LargeImage = RoomGenPicSrc;
            ApartGenPushBtn.Image = RoomGenPicSrc;

            
            //Создание кнопки квартирографии
            var ApartmentgraphyButton = new PushButtonData("Квартирография", "Квартирография", Assembly.GetExecutingAssembly().Location, "GeoAddin.Apartmentgraphy");
            var ApartmentgraphyPushBtn = archpanel.AddItem(ApartmentgraphyButton) as PushButton;
            Image ApartmentgraphyButtonPic = Properties.Resources.ApartmentgraphyPic;
            ImageSource ApartmentgraphyButtonPicSrc = Convert(ApartmentgraphyButtonPic);
            ApartmentgraphyPushBtn.LargeImage = ApartmentgraphyButtonPicSrc;
            ApartmentgraphyPushBtn.Image = ApartmentgraphyButtonPicSrc;

            //Создание кнопки заполнения окон
            var WindowsFillingButton = new PushButtonData("Заполнение окон", "Заполнение окон", Assembly.GetExecutingAssembly().Location, "GeoAddin.WindowsFilling");
            var WindowsFillingPushBtn = archpanel.AddItem(WindowsFillingButton) as PushButton;
            Image WindowsFillingButtonPic = Properties.Resources.WindowFilling;
            ImageSource WindowsFillingButtonPicSrc = Convert(WindowsFillingButtonPic);
            WindowsFillingPushBtn.LargeImage = WindowsFillingButtonPicSrc;
            WindowsFillingPushBtn.Image = WindowsFillingButtonPicSrc;

            //Создание кнопки создания оконных схем
            var WindowsSchemaButton = new PushButtonData("Схема окон", "Схема окон", Assembly.GetExecutingAssembly().Location, "GeoAddin.WindowsSchema");
            var WindowsSchemaPushBtn = archpanel.AddItem(WindowsSchemaButton) as PushButton;
            Image WindowsSchemaButtonPic = Properties.Resources.WindowSchema;
            ImageSource WindowsSchemaButtonPicSrc = Convert(WindowsSchemaButtonPic);
            WindowsSchemaPushBtn.LargeImage = WindowsSchemaButtonPicSrc;
            WindowsSchemaPushBtn.Image = WindowsSchemaButtonPicSrc;

            //Создание кнопки отсоединения файла
            var DetachFileButton = new PushButtonData("Отсоединение файла", "Отсоединение файла", Assembly.GetExecutingAssembly().Location, "GeoAddin.DetachFile");
            var DetachFilePushBtn = commonpanel.AddItem(DetachFileButton) as PushButton;
            Image DetachFileButtonPic = Properties.Resources.DetachFilePic;
            ImageSource DetachFileButtonPicSrc = Convert(DetachFileButtonPic);
            DetachFilePushBtn.LargeImage = DetachFileButtonPicSrc;
            DetachFilePushBtn.Image = DetachFileButtonPicSrc;

            return Result.Succeeded;
        }
        //Метод для конвертации картинки
        public BitmapImage Convert (Image img)
        {
            using (var memory = new MemoryStream())
            {
                img.Save(memory, ImageFormat.Png);
                memory.Position = 0;

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;

            }
        }
        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
