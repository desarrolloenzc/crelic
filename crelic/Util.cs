using System;
using System.Collections.Generic;
using System.Windows;
using System.IO;

using System.Diagnostics;

using System.Windows.Threading;
using ExcelDataReader;
using System.Net.Mail;
using System.Configuration;
using System.Threading;
using System.Data.SqlTypes;

namespace crelic
{
    public class Util
    {

        public void KillExcel()
        {
            try
            {
                Process[] process = Process.GetProcessesByName("excel");
                foreach (Process excel in process)
                {
                    excel.Kill();
                }
                process = null;
            }
            catch (Exception theException)
            {
                String errorMessage = string.Empty;
                MessageBox.Show(errorMessage, theException.Message);
                MainWindow.strlog = MainWindow.strlog + "\n" + theException.Message;
            }
        }

        public void ForceUIToUpdate()
        {
            DispatcherFrame frame = new DispatcherFrame();

            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Render, new DispatcherOperationCallback(delegate (object parameter)
            {
                frame.Continue = false;
                return null;
            }), null);

            Dispatcher.PushFrame(frame);
        }

        public FileStream ReallyRead(ref System.Data.DataSet dataSet, string filePath)
        {
            FileStream stream = null;

            try
            {
                using (stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader;
                    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
                    var conf = new ExcelDataSetConfiguration
                    {
                        UseColumnDataType = false,
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = false
                        }
                    };
                    dataSet = reader.AsDataSet(conf);
                }
            }
            catch (IOException)
            {
                MessageBox.Show(filePath + " Archivo abierto o no encontrado");
                MainWindow.strlog = MainWindow.strlog + "\n" + "ReallyRead " + filePath + " Archivo abierto o no encontrado";
                return null;
            }
            return stream;
        }

        public void Convert(string xlsFilePath,string xlsxFilePath)
        {

            // Create an instance of Excel Application
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Make Excel visible (optional)
            excelApp.Visible = false;

            // Open the XLS file
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(xlsFilePath);

            // Save as XLSX file format
            workbook.SaveAs(xlsxFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);

            // Close the workbook and Excel application
            workbook.Close();
            excelApp.Quit();

           File.Delete(xlsFilePath);
        }


    public void WriteLog(string strLog)
        {
            StreamWriter log;
            FileStream fileStream = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;
            DateTime current = DateTime.Now;
            strLog = "INICIO " + current.ToString() + "\n" + strLog + "\n";

            string logFilePath = MainWindow.pathConfiguracion;
            int pos = logFilePath.IndexOf(".xlsm");
            if (pos == -1)
                return;
            logFilePath = logFilePath.Substring(0, pos);
            logFilePath = logFilePath + " Log-" + System.DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
            logFileInfo = new FileInfo(logFilePath);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists) logDirInfo.Create();
            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);
            }
            log = new StreamWriter(fileStream);
            log.WriteLine(strLog);
            log.Close();
        }

        public bool envioEmail()
        {
            //smtp.office365.com
            //puerto: 587
            string destinatario = string.Empty;
            try
            {
                SmtpClient smtp = new SmtpClient();
                //smtp.Port = 587;
                smtp.Port = int.Parse(ConfigurationManager.AppSettings["smtport"].ToString());
                //smtp.Host = "smtp.office365.com";
                smtp.Host = ConfigurationManager.AppSettings["smthost"].ToString();
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                //smtp.Credentials = new System.Net.NetworkCredential("mopie@estrella.com.do", "bwjqpnnzjtmtzwdm");
                smtp.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["smtcreduser"].ToString(),
                    ConfigurationManager.AppSettings["smtcredpass"].ToString());

                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n Error enviando correo a destinatario: " + destinatario + " !");
                return false;
            }
        }

        public bool LecturaUpazila()
        {
            FileStream stream = MainWindow.utilidad.ReallyRead(ref MainWindow.dataConfiguraSet, MainWindow.pathConfiguracion);
            int fila = -1;
            foreach (System.Data.DataRow drCurrent1 in MainWindow.dataConfiguraSet.Tables["Upazila"].Rows)
            {
                fila++;
                Upazilas up = new Upazilas(drCurrent1[0].ToString(), drCurrent1[1].ToString(), drCurrent1[2].ToString());
                MainWindow.upazilas.Add(up);
            }

            return true;
        }

        public bool LecturaHazard()
        {
            int fila = -1;
            foreach (System.Data.DataRow drCurrent1 in MainWindow.dataConfiguraSet.Tables["Hazard"].Rows)
            {
                fila++;
                if (fila == 0)
                    continue;
                Hazard ha = new Hazard(drCurrent1[0].ToString());
                ha.veryhigh = float.Parse(drCurrent1[1].ToString());
                ha.high = float.Parse(drCurrent1[2].ToString());
                ha.medium = float.Parse(drCurrent1[3].ToString());
                ha.low = float.Parse(drCurrent1[4].ToString());
                ha.verylow = float.Parse(drCurrent1[5].ToString());
                MainWindow.hazards.Add(ha);
            }

            return true;
        }

        public bool CrearEstructura()
        {
            for (int i = 0; i < MainWindow.upazilas.Count; i++)
            {
                List<UpazilaHazard> upas = new List<UpazilaHazard>();                
                for (int j = 0; j < MainWindow.hazards.Count; j++)
                {
                    UpazilaHazard upita = new UpazilaHazard(MainWindow.upazilas[i].upazila, MainWindow.hazards[j].hazard, "",0);
                    upas.Add(upita);
                }
                MainWindow.padre.Add(upas);
            }
            return true;
        }

        public void TransformaArchivos()
        {
            if (Directory.Exists(MainWindow.carpeta))
            {
                // Get all files in the directory
                string[] files = Directory.GetFiles(MainWindow.carpeta);

                // Iterate through each file
                foreach (string file in files)
                {
                    try
                    {
                        string[] extension = file.Split('.');
                        if (extension[1] == "xls")
                        {
                            string destino = extension[0] + ".xlsx";
                            MainWindow.utilidad.Convert(file, destino);
                            Thread.Sleep(2000);
                        }

                    }
                    catch (Exception ex)
                    {
                        // Handle exceptions if any
                        Console.WriteLine("Error reading file: " + ex.Message);
                    }

                }

                files = Directory.GetFiles(MainWindow.carpeta);
                DirectoryInfo parentDirectory = Directory.GetParent(MainWindow.carpeta);
                string currentDirectory = Directory.GetCurrentDirectory();
                foreach (string file in files)
                {
                    try
                    {
                        string[] extension = file.Split('.');
                        if (extension[1] == "xlsx" && file.Contains("-graph"))
                        {
                            string archivo = Path.GetFileName(file);
                            //string destinationFile = Path.Combine(parentDirectory.FullName+"\\Origin\\", archivo);
                            string destinationFile = Path.Combine(currentDirectory + "\\Origin\\", archivo);
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

            }

        }

        public void LlenaData()
        {
            if (Directory.Exists(MainWindow.carpeta))
            {
                // Get all files in the directory
                string[] files = Directory.GetFiles(MainWindow.carpeta);

                // Iterate through each file
                foreach (string file in files)
                {
                    try
                    {
                        string[] extension = file.Split('.');
                        // Read the content of the file
                        FileStream stream = MainWindow.utilidad.ReallyRead(ref MainWindow.dataOrigenSet, file);

                        int fila = -1;
                        string upazila = "";
                        string hazard = "";
                        foreach (System.Data.DataRow drCurrent1 in MainWindow.dataOrigenSet.Tables[0].Rows)
                        {
                            fila++;
                            if (fila == 0)
                            {
                                string cadena = drCurrent1[0].ToString();
                                int i = cadena.IndexOf("Graph");
                                hazard = cadena.Substring(0, i - 1);
                                int longi = cadena.Length;
                                upazila = cadena.Substring(i + 7, longi - 8-i -8);
                            }
                            if (fila < 2)
                                continue;
                            int level = int.Parse(drCurrent1[1].ToString());
                            bool continuar = true;

                            List<UpazilaHazard> apuntador = new List<UpazilaHazard>();
                            for (int i = 0; i < MainWindow.padre.Count; i++)
                            {
                                apuntador = MainWindow.padre[i];
                                if (apuntador[0].upazila == upazila)
                                {
                                    for (int j = 0; j < apuntador.Count; j++)
                                    {
                                        if (apuntador[j].hazard.ToLower() == hazard.ToLower())
                                        {
                                            apuntador[j].leido = true;
                                            apuntador[j].level = level;
                                            continuar = false;
                                            break;
                                        }
                                    }
                                }
                                if (!continuar)
                                    break;
                            }
                            if (!continuar)
                                break;
                        }
                        stream.Close();
                    }
                    catch (Exception ex)
                    {
                        // Handle exceptions if any
                        Console.WriteLine("Error reading file: " + ex.Message);
                    }
                }
            }
            else
            {
                Console.WriteLine("Directory does not exist.");
            }

        }

        public void LlenaIndicadores(ref Excel exe, string hazard, List<UpazilaHazard> apuntador,int indice)
        {
            try
            {
                int j = indice;
                if (apuntador[j].hazard.Trim() == hazard)
                {
                    if (apuntador[j].level == 5)
                    {
                        float valor = this.retornaNumeroInd(hazard, apuntador[j].level);
                        exe.ws.Cells[10 + j, 2] = valor;
                    }
                    if (apuntador[j].level == 4)
                    {
                        float valor = this.retornaNumeroInd(hazard, apuntador[j].level);
                        exe.ws.Cells[10 + j, 3] = valor;
                    }
                    if (apuntador[j].level == 3)
                    {
                        float valor = this.retornaNumeroInd(hazard, apuntador[j].level);
                        exe.ws.Cells[10 + j, 4] = valor;
                    }
                    if (apuntador[j].level == 2)
                    {
                        float valor = this.retornaNumeroInd(hazard, apuntador[j].level);
                        exe.ws.Cells[10 + j, 5] = valor;
                    }
                    if (apuntador[j].level == 1)
                    {
                        float valor = this.retornaNumeroInd(hazard, apuntador[j].level);
                        exe.ws.Cells[10 +j, 6] = valor;
                    }
                    if (apuntador[j].level == 0)
                    {
                        exe.ws.Cells[10 + j, 7] = "X";
                    }
                }
            }
            catch { return; }
                            
        }

        public float retornaNumeroInd(string hazard,int level)
        {
            float indi = -1;
            for (int i = 0; i < MainWindow.hazards.Count; i++)
            {
                if (MainWindow.hazards[i].hazard == hazard)
                    switch (level) {
                        case 5:
                            indi = MainWindow.hazards[i].veryhigh;
                            break;
                        case 4:
                            indi = MainWindow.hazards[i].high;
                            break;
                        case 3:
                            indi = MainWindow.hazards[i].medium;
                            break;
                        case 2:
                            indi = MainWindow.hazards[i].low;
                            break;
                        case 1:
                            indi = MainWindow.hazards[i].verylow;
                            break;
                    }
            }
            return (float)Math.Round(indi, 2);
        }
    }
}


