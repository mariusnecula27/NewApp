using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace NewApp
{
    class ExcelFile
    {
        public Excel.Application xlApp;
        public Excel.Workbook xlWorkBook;
        public Excel.Worksheet xlWorkSheetSignal;
        public Excel.Worksheet xlWorkSheetEcuInstance;
        public object misValue;

        public ExcelFile()
        {
            
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
                return;
            }

            misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheetSignal = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheetSignal.Name = "Signals";

            xlWorkSheetSignal.Columns[1].ColumnWidth = 30;
            xlWorkSheetSignal.Columns[2].ColumnWidth = 25;
            xlWorkSheetSignal.Columns[3].ColumnWidth = 10;
            xlWorkSheetSignal.Columns[4].ColumnWidth = 10;
            xlWorkSheetSignal.Columns[5].ColumnWidth = 60;

            xlWorkSheetSignal.Cells[1, 1] = "Short Name";
            xlWorkSheetSignal.Cells[1, 2] = "Data Type Policy";
            xlWorkSheetSignal.Cells[1, 3] = "Length";
            xlWorkSheetSignal.Cells[1, 4] = "IDT";
            xlWorkSheetSignal.Cells[1, 5] = "Sistem Signal";

            xlWorkSheetEcuInstance = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            xlWorkSheetEcuInstance.Name = "ECU Instances";

            xlWorkSheetEcuInstance.Columns[1].ColumnWidth = 15;
            xlWorkSheetEcuInstance.Columns[2].ColumnWidth = 15;
            xlWorkSheetEcuInstance.Columns[3].ColumnWidth = 15;
            xlWorkSheetEcuInstance.Columns[4].ColumnWidth = 15;
            xlWorkSheetEcuInstance.Columns[5].ColumnWidth = 15;
            xlWorkSheetEcuInstance.Columns[6].ColumnWidth = 15;
            xlWorkSheetEcuInstance.Columns[7].ColumnWidth = 15;

            xlWorkSheetEcuInstance.Cells[1, 1] = "Name";
            xlWorkSheetEcuInstance.Cells[1, 2] = "GW TimeBase";
            xlWorkSheetEcuInstance.Cells[1, 3] = "TX TimeBase";
            xlWorkSheetEcuInstance.Cells[1, 4] = "RX TimeBase";
            xlWorkSheetEcuInstance.Cells[1, 5] = "Cyclic Transmission";
            xlWorkSheetEcuInstance.Cells[1, 6] = "Sleep Mode";
            xlWorkSheetEcuInstance.Cells[1, 7] = "Supported Wake-Up";

            for (int i = xlApp.ActiveWorkbook.Worksheets.Count; i > 0; i--)
            {
                Worksheet wkSheet = (Worksheet)xlApp.ActiveWorkbook.Worksheets[i];
                if (wkSheet.Name == "Sheet3")
                {
                    wkSheet.Delete();
                }
            }
        }

        public void writeSignalSheet(int index, string infos, Worksheet xlWorkSheetSignal)
        {
            string[] datas = infos.Split(null);
            xlWorkSheetSignal.Cells[index + 2, 1] = datas[0];
            xlWorkSheetSignal.Cells[index + 2, 2] = datas[1];
            xlWorkSheetSignal.Cells[index + 2, 3] = datas[2];
            xlWorkSheetSignal.Cells[index + 2, 4] = datas[3].Split('/').Last();
            xlWorkSheetSignal.Cells[index + 2, 5] = datas[4];
        }

        public void writeEcuInstanceSheet(int index, string infos, Worksheet xlWorkSheetEcuInstance)
        {
            string[] datas = infos.Split(null);
            xlWorkSheetEcuInstance.Cells[index + 2, 1] = datas[0];
            xlWorkSheetEcuInstance.Cells[index + 2, 2] = datas[1];
            xlWorkSheetEcuInstance.Cells[index + 2, 3] = datas[2];
            xlWorkSheetEcuInstance.Cells[index + 2, 4] = datas[3];
            xlWorkSheetEcuInstance.Cells[index + 2, 5] = datas[4];
            xlWorkSheetEcuInstance.Cells[index + 2, 6] = datas[5];
            xlWorkSheetEcuInstance.Cells[index + 2, 7] = datas[6];
        }

        public void saveExcelFile(string pathXls, string pathXlsx, Workbook xlWorkBookk, System.Windows.Forms.CheckBox checkBox1, System.Windows.Forms.CheckBox checkBox2, object misValue, System.Windows.Forms.TextBox textbox)
        {
            if(checkBox1.Checked is true && checkBox2.Checked is true)
            {
                xlWorkBookk.SaveAs(pathXlsx, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBookk.SaveAs(pathXls, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            else if(checkBox1.Checked is true && checkBox2.Checked is false)
            {
                xlWorkBookk.SaveAs(pathXls, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            else if(checkBox1.Checked is false && checkBox2.Checked is true)
            {
                xlWorkBookk.SaveAs(pathXlsx, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            else if(checkBox1.Checked is false && checkBox2.Checked is false)
            {
                textbox.Text = "Nu s-a ales nicio optiune pentru crearea fisierului de tip excel!";
                return;
            }

            xlWorkBookk.Close(true, misValue, misValue);
            xlApp.Quit();
            MessageBox.Show("Excel file/files created , you can find the it at the selected path!");

        }
    }
}
