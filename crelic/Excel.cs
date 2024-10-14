using Microsoft.Office.Interop.Excel;
using crelic;
using System;
using System.IO;
using System.Windows;

public class Excel
{
    public Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
    public Workbook wb;
    public Worksheet ws;
    public Range xlRange;
   
    public Excel(string path, string Sheet, ref bool existe)
    {
        if (!File.Exists(path))
        {
            MessageBox.Show("Archivo " + path + " no existe");
            existe = false;
            return;
        }
        try
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
            excel.DisplayAlerts = false;            

            xlRange = ws.UsedRange;
            //wb.Close();
            existe = true;
        }
        catch
        {
            existe = false;
            MessageBox.Show("Error accediendo a " + path);
            return;
        }
    }
    public Excel(string path, string v, int Sheet, ref bool existe)
    {
        if (!File.Exists(path))
        {
            MessageBox.Show("Archivo " + path + " no existe");
            existe = false;
            return;
        }
        try
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
            excel.DisplayAlerts = false;

            xlRange = ws.UsedRange;
            //wb.Close();
            existe = true;
        }
        catch
        {
            existe = false;
            MessageBox.Show("Error accediendo a " + path);
            return;
        }
    }

    public Excel(string path, ref bool existe)
    {
        if (!File.Exists(path))
        {
            MessageBox.Show("Archivo " + path + "no existe");
            existe = false;
            return;
        }
        existe = true;
    }

    public bool ReadWriteCell(Worksheet ws, string value, int fila, int columna)
    {
        Range xlRange = ws.UsedRange;
        ws.Cells[fila, columna].Value = value;
        return true;
    }
}

