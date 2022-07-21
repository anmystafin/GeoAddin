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

            string openingpanelName = "Отверстия";
            var openingpanel = a.CreateRibbonPanel(tabName, openingpanelName);

            //Создание кнопки генерациии помещений
            var ApartGenButton = new PushButtonData("Генерация\nквартир", "Генерация\nквартир", Assembly.GetExecutingAssembly().Location, "GeoAddin.RoomGenerating");
            var ApartGenPushBtn = archpanel.AddItem(ApartGenButton) as PushButton;
            Image RoomGenPic = Properties.Resources.RoomGenPic;
            ApartGenPushBtn.LargeImage = Convert(RoomGenPic, new Size(32, 32));
            ApartGenPushBtn.Image = Convert(RoomGenPic, new Size(16, 16));


            //Создание кнопки квартирографии
            var ApartmentgraphyButton = new PushButtonData("Квартирография", "Квартирография", Assembly.GetExecutingAssembly().Location, "GeoAddin.Apartmentgraphy");
            var ApartmentgraphyPushBtn = archpanel.AddItem(ApartmentgraphyButton) as PushButton;
            Image ApartmentgraphyButtonPic = Properties.Resources.ApartmentgraphyPic;
            ApartmentgraphyPushBtn.LargeImage = Convert(ApartmentgraphyButtonPic, new Size(32, 32));
            ApartmentgraphyPushBtn.Image = Convert(ApartmentgraphyButtonPic, new Size(16, 16));

            //Создание кнопки заполнения окон
            var WindowsFillingButton = new PushButtonData("Заполнение\nокон", "Заполнение\nокон", Assembly.GetExecutingAssembly().Location, "GeoAddin.WindowsFilling");
            var WindowsFillingPushBtn = archpanel.AddItem(WindowsFillingButton) as PushButton;
            Image WindowsFillingButtonPic = Properties.Resources.WindowFilling;
            WindowsFillingPushBtn.LargeImage = Convert(WindowsFillingButtonPic, new Size(32, 32));
            WindowsFillingPushBtn.Image = Convert(WindowsFillingButtonPic, new Size(16, 16));

            //Создание кнопки создания оконных схем
            var WindowsSchemaButton = new PushButtonData("Схема\nокон", "Схема\nокон", Assembly.GetExecutingAssembly().Location, "GeoAddin.WindowsSchema");
            var WindowsSchemaPushBtn = archpanel.AddItem(WindowsSchemaButton) as PushButton;
            Image WindowsSchemaButtonPic = Properties.Resources.WindowSchema;
            WindowsSchemaPushBtn.LargeImage = Convert(WindowsSchemaButtonPic, new Size(32, 32));
            WindowsSchemaPushBtn.Image = Convert(WindowsSchemaButtonPic, new Size(16, 16));

            //Создание кнопки отсоединения файла
            var DetachFileButton = new PushButtonData("Отсоединение\nфайла", "Отсоединение\nфайла", Assembly.GetExecutingAssembly().Location, "GeoAddin.DetachFile");
            var DetachFilePushBtn = commonpanel.AddItem(DetachFileButton) as PushButton;
            Image DetachFileButtonPic = Properties.Resources.DetachFilePic;
            DetachFilePushBtn.LargeImage =  Convert(DetachFileButtonPic, new Size(32, 32)) ;
            DetachFilePushBtn.Image = Convert(DetachFileButtonPic, new Size(16, 16));

            //Создание кнопки генерации отверстий в МЕР
            var OpeningGeneratingButton = new PushButtonData("Генерация\nотверстий", "Генерация\nотверстий", Assembly.GetExecutingAssembly().Location, "GeoAddin.OpeningGenerating");
            var OpeningGeneratingPushBtn = openingpanel.AddItem(OpeningGeneratingButton) as PushButton;
            Image OpeningGeneratingButtonPic = Properties.Resources.CM_OpeningMEP;
            OpeningGeneratingPushBtn.LargeImage = Convert(OpeningGeneratingButtonPic, new Size(32, 32));
            OpeningGeneratingPushBtn.Image = Convert(OpeningGeneratingButtonPic, new Size(16, 16));

            //Создание кнопки генерации ID отверстий в МЕР
            var OpeningIDGeneratingButton = new PushButtonData("Генерация ID", "Генерация ID", Assembly.GetExecutingAssembly().Location, "GeoAddin.OpeningIDGenerating");
            var OpeningIDGeneratingPushBtn = openingpanel.AddItem(OpeningIDGeneratingButton) as PushButton;
            Image OpeningIDGeneratingButtonPic = Properties.Resources.OpeningPic;
            OpeningIDGeneratingPushBtn.LargeImage = Convert(OpeningIDGeneratingButtonPic, new Size(32, 32));
            OpeningIDGeneratingPushBtn.Image = Convert(OpeningIDGeneratingButtonPic, new Size(16, 16));

            //Создание кнопки копирования отверстий в АР/КР
            var OpeningCopyingButton = new PushButtonData("Копирование\nотверстий", "Копирование\nотверстий", Assembly.GetExecutingAssembly().Location, "GeoAddin.OpeningCopying");
            var OpeningCopyingPushBtn = openingpanel.AddItem(OpeningCopyingButton) as PushButton;
            Image OpeningCopyingButtonPic = Properties.Resources.OpeningPic;
            OpeningCopyingPushBtn.LargeImage = Convert(OpeningCopyingButtonPic, new Size(32, 32));
            OpeningCopyingPushBtn.Image = Convert(OpeningCopyingButtonPic, new Size(16, 16));

            return Result.Succeeded;
        }
        //Метод для конвертации картинки
        public BitmapImage Convert (Image img, Size size)
        {   
            img = (Image)(new Bitmap(img, size));
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
