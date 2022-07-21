using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using RevitServices.Persistence;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using RevitApplication = Autodesk.Revit.ApplicationServices.Application;

namespace GeoAddin
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class OpeningCopyWindow : Window
    {
        
        
        public bool clickedon;
        public bool clickedoff;

        
        private UIDocument uidoc;
        
        private Document doc;

        public OpeningCopyWindow(UIApplication uiapp)
        {
            
            InitializeComponent();
            
            uidoc = uiapp.ActiveUIDocument;
            
            doc = uidoc.Document;
            clickedon = false;
            clickedoff = false;
            
            List<string> linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().Select(el => el.Name).ToList();
            
            LinkInstance.ItemsSource = linkInstances;
        }

        private void But_Click_On(object sender, RoutedEventArgs e)
        {
            clickedon = true;
            Close();
        }

        private void But_Click_Off(object sender, RoutedEventArgs e)
        {
            clickedoff = true;
            Close();
        }
    }
}
