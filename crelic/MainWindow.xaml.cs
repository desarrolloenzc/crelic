using System;
using System.IO;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using System.Globalization;

namespace crelic
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
        
    public partial class MainWindow : System.Windows.Window
    {       
        public static Util utilidad = new Util();
        public static List<Upazilas> upazilas = new List<Upazilas>();
        public static List<Hazard> hazards = new List<Hazard>();
        public static List<List<UpazilaHazard>> padre = new List<List<UpazilaHazard>>();

        public static Frame frame;

        /* de proyectos */
        public static string carpeta = string.Empty;
        public static string pathConfiguracion = System.IO.Directory.GetCurrentDirectory() + "\\Administration\\upazila.xlsx";
        public static string pathBbdd = System.IO.Directory.GetCurrentDirectory();
        
        public static string strlog = string.Empty;
        public static NumberFormatInfo provider = new NumberFormatInfo();
        //public static NumberFormatInfo provider = new CultureInfo("en-US", false).NumberFormat;

        public static System.Data.DataSet dataSet; //configuracion
        public static System.Data.DataSet dataConfiguraSet; //Configuracion
        public static System.Data.DataSet dataOrigenSet; //sap
        public static System.Data.DataSet dataEliminacionesSet; //bbdd
        public static System.Data.DataSet dataAjustesSet; //ajustes
        public static System.Data.DataSet dataModeloSet; //modelo

        Page1 page1 = new Page1();
        

        Color color1 = Color.FromRgb(217, 91, 38);
        Color color2 = Color.FromRgb(128, 128, 128);
        Color color3 = Color.FromRgb(255, 165, 90);

        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Content = page1;
            provider.NumberDecimalSeparator = ",";
            provider.NumberGroupSeparator = ".";
            frame = MainFrame;
            MainWindow.utilidad.KillExcel();
            ComprobarAdministracion();

            DirectoryInfo parentDirectory = Directory.GetParent(pathBbdd);
            pathBbdd = parentDirectory.FullName + "\\Database\\database.xlsx";
        }

        private void QuitBtn_Click(object sender, RoutedEventArgs e)
        {
         
            Environment.Exit(1);
        }

        private void configuracion_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Content = page1;
        }

        private void reportes_Click(object sender, RoutedEventArgs e)
        {
        }

        private void ComprobarAdministracion()
        {
           
        }
    }
}
