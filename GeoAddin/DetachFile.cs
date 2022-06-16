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
    public class DetachFile: IExternalCommand
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
            DetachFileWindow win = new DetachFileWindow();
            win.ShowDialog();
            bool clickedon = win.clickedon;
            bool clickedoff = win.clickedoff;
            while (clickedon == false & clickedoff == false)
            {
                continue;
            }
            if (clickedon == true)
            {
                //Синхронизация файла
                TransactWithCentralOptions toptions = new TransactWithCentralOptions();
                SynchronizeWithCentralOptions soptions = new SynchronizeWithCentralOptions();
                soptions.SaveLocalBefore = true;
                soptions.SaveLocalAfter = true;
                soptions.Comment = win.comment;
                doc.SynchronizeWithCentral(toptions, soptions);

                //Сохранение файла
                string activedoccentralpath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                string activedocpath = doc.PathName;
                string sharedpath = activedoccentralpath.Replace("Project", "Shared");
                string sharedoctitlepath = sharedpath.Replace(".rvt", "_Отсоединено.rvt");
                string detachdocpath = sharedoctitlepath;
                SaveAsOptions saveoptions = new SaveAsOptions();
                WorksharingSaveAsOptions wsoptions = new WorksharingSaveAsOptions();
                wsoptions.SaveAsCentral = true;
                saveoptions.SetWorksharingOptions(wsoptions);
                saveoptions.OverwriteExistingFile = true;
                doc.SaveAs(detachdocpath, saveoptions);
                Document detachdoc = uidoc.Document;
                

                //Открытие локального файла и закрытие сохраненной отсоединенной копии
                OpenOptions openoptions = new OpenOptions();
                WorksetConfiguration wsconfig = new WorksetConfiguration();
                openoptions.SetOpenWorksetsConfiguration(wsconfig);
                FilePath filePath = new FilePath(activedocpath);
                uiapp.OpenAndActivateDocument(filePath, openoptions, false);

                detachdoc.Close();





            }
            return Result.Succeeded;
        }
        
    }
}
