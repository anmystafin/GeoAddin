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
            var panel = a.CreateRibbonPanel(tabName,archpanelName);

            //Создание кнопки генерациии помещений
            var ApartGenButton = new PushButtonData("Генерация помещений", "Генерация помещений", Assembly.GetExecutingAssembly().Location, "GeoAddin.RoomGenerating");
            var ApartGenPushBtn = panel.AddItem(ApartGenButton) as PushButton;
            Image RoomGenPic = Properties.Resources.RoomGenPic;
            ImageSource RoomGenPicSrc = Convert(RoomGenPic);
            ApartGenPushBtn.LargeImage = RoomGenPicSrc;
            ApartGenPushBtn.Image = RoomGenPicSrc;

            //Создание кнопки нумерации квартир
            var ApartNumButton = new PushButtonData("Нумерация квартир", "Нумерация квартир", Assembly.GetExecutingAssembly().Location, "GeoAddin.ApartmentNumFilling");
            var ApartNumPushBtn = panel.AddItem(ApartNumButton) as PushButton;
            Image ApartNumPic = Properties.Resources.ApartmentNum;
            ImageSource ApartNumPicSrc = Convert(ApartNumPic);
            ApartNumPushBtn.LargeImage = ApartNumPicSrc;
            ApartNumPushBtn.Image = ApartNumPicSrc;

            //Создание кнопки квартирографии
            var ApartmentgraphyButton = new PushButtonData("Квартирография", "Квартирография", Assembly.GetExecutingAssembly().Location, "GeoAddin.Apartmentgraphy");
            var ApartmentgraphyPushBtn = panel.AddItem(ApartmentgraphyButton) as PushButton;
            Image ApartmentgraphyButtonPic = Properties.Resources.ApartmentgraphyPic;
            ImageSource ApartmentgraphyButtonPicSrc = Convert(ApartmentgraphyButtonPic);
            ApartmentgraphyPushBtn.LargeImage = ApartmentgraphyButtonPicSrc;
            ApartmentgraphyPushBtn.Image = ApartmentgraphyButtonPicSrc;

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
