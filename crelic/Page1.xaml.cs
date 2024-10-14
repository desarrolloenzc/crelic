using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Linq;

namespace crelic
{
    /// <summary>
    /// Interaction logic for Page1.xaml
    /// </summary>
    public partial class Page1 : System.Windows.Controls.Page
    {       

        public Page1()
        {
            InitializeComponent();
            //lectura de upazilas
            MainWindow.utilidad.LecturaUpazila();
            //lectura de hazards
            MainWindow.utilidad.LecturaHazard();
            //crear estructura
            MainWindow.utilidad.CrearEstructura();

        }

        private void Fuente_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                MainWindow.carpeta = dialog.FileName;
            }

            labeluno.Text = "Working...";           
            MainWindow.utilidad.ForceUIToUpdate();

            MainWindow.utilidad.LlenaData();

            List<UpazilaHazard> apuntador;
            bool existe = false;
            Excel exe = new Excel(MainWindow.pathBbdd, "Model", ref existe);

            /*for (int i = 0; i < MainWindow.padre.Count; i++)
            {
                string upazila = MainWindow.padre[i][0].upazila;
                exe.ws = exe.wb.Worksheets[upazila];
                Range range = exe.ws.Range["B10:G20"];
                range.ClearContents();
            }*/

            for (int i = 0; i < MainWindow.padre.Count; i++)
            {
                string upazila = MainWindow.padre[i][0].upazila;
                if (upazila == "")
                    break;
                try
                {
                    exe.ws = exe.wb.Worksheets[upazila];
                }
                catch 
                {                   
                    exe.ws = exe.wb.Worksheets.Add();                
                    exe.ws.Name = upazila;
                    exe.ws = exe.wb.Sheets["Model"];
                    Range sourceRange = exe.ws.Range["A1:G20"]; // Adjust the range as needed
                    exe.ws = exe.wb.Sheets[upazila];
                    Range targetRange = exe.ws.Range["A1:G20"]; // Adjust the starting cell as needed

                    // Copy the data from the source range
                    sourceRange.Copy();

                    // Paste the data into the target range
                    targetRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
                    targetRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
                    for (int m = 0; m < MainWindow.upazilas.Count; m++)
                        if (MainWindow.upazilas[m].upazila == upazila)
                        {
                            exe.ws.Cells[3, 1] = "Division: " + MainWindow.upazilas[m].division;
                            exe.ws.Cells[4, 1] = "District: " + MainWindow.upazilas[m].distrito;
                            exe.ws.Cells[5, 1] = "Upazila: " + MainWindow.upazilas[m].upazila;
                            break;
                        }
                }
                apuntador = MainWindow.padre[i];
                for (int p = 0; p < MainWindow.hazards.Count; p++)
                {
                    if (apuntador[p].leido)
                        MainWindow.utilidad.LlenaIndicadores(ref exe, MainWindow.hazards[p].hazard, apuntador, p);
                }
            }

            if (existe)
                exe.wb.Save();
            exe.wb.Close();

            string[] files = Directory.GetFiles(MainWindow.carpeta);
            DirectoryInfo parentDirectory = Directory.GetParent(MainWindow.carpeta);
            foreach (string file in files)
            {
                try
                {
                    string[] extension = file.Split('.');
                    if (extension[1] == "xlsx")
                    {
                        string archivo = Path.GetFileName(file);
                        string destinationFile = Path.Combine(parentDirectory.FullName + "\\Backup\\", archivo);
                        if (File.Exists(destinationFile))
                        {
                            // Delete the destination file if it exists
                            File.Delete(destinationFile);
                        }
                        File.Move(file, destinationFile);
                    }

                }
                catch (Exception ex)
                {
                    // Handle exceptions if any
                    Console.WriteLine("Error reading file: " + ex.Message);
                }

            }

            labeluno.Text = "Ready";
            MainWindow.utilidad.ForceUIToUpdate();

        }

        private void Fuente_F01_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void Otro_Click(object sender, RoutedEventArgs e)
        {

        }
        private void Ajustes_Click(object sender, RoutedEventArgs e)
        {
        }

        private void Tansforma_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                MainWindow.carpeta = dialog.FileName;
            }

            labeluno.Text = "Working...";
            MainWindow.utilidad.ForceUIToUpdate();

            MainWindow.utilidad.TransformaArchivos();

            labeluno.Text = "Ready";
            MainWindow.utilidad.ForceUIToUpdate();

        }

        private void Order_Click(object sender, RoutedEventArgs e)
        {
            String archivo = "";
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                archivo = dialog.FileName + "\\database.xlsx";
            }

            // Open the workbook
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(archivo);
            
            try
            {                
                // Get the sheets and their names
                Sheets sheets = workbook.Sheets;
                string[] sheetNames = new string[sheets.Count];
                for (int i = 0; i < sheets.Count; i++)
                {
                    sheetNames[i] = sheets[i + 1].Name;
                }

                // Sort the sheet names alphabetically
                Array.Sort(sheetNames, StringComparer.OrdinalIgnoreCase);

                // Rearrange the sheets according to the sorted names
                for (int i = 0; i < sheetNames.Length; i++)
                {
                    Worksheet sheet = (Worksheet)sheets[sheetNames[i]];
                    sheet.Move(After: sheets[sheets.Count]);
                }

                // Save the workbook
                workbook.Save();
                Console.WriteLine("Sheets sorted alphabetically and workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Close the workbook and quit the Excel application
                if (workbook != null)
                {
                    workbook.Close();
                }
                    
            }
        }
    }


}
