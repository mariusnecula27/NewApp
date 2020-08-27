using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop;
using System.Data.OleDb;
using DataTable = System.Data.DataTable;
using System.Xml.Serialization;
using System.Xml.Schema;
using System.Diagnostics;
using System.Windows.Forms.VisualStyles;

namespace NewApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if(openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string strfilename = openFileDialog.FileName;

                textBox1.Text = strfilename ;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(radioButton1.Checked )
            {
                radioButton1_ExcelToXML();
            }
            if (radioButton2.Checked)
            {
                radioButton2_XMLToExcel();
            }
        
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void radioButton1_ExcelToXML()
        {
            string pathFile = textBox1.Text;
            var xlApp = new Excel.Application();
            //xlApp.Visible = true;


            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(pathFile);
            Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorkSheet.UsedRange;


            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Worksheet sheetOne = xlWorkBook.Sheets["Signals"];
            Excel.Worksheet sheetTwo = xlWorkBook.Sheets["ECU Instances"];

            string a11 = sheetOne.Cells[1, 1].Value.ToString();
            Console.WriteLine(a11);

            Excel.Range last = sheetOne.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheetOne.get_Range("A1", last);

            int lastUsedRow = last.Row;
            int lastUsedColumn = last.Column;



            Excel.Range lastsheetTwo = sheetTwo.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range rangeSheetTwo = sheetTwo.get_Range("A1", lastsheetTwo);


            xmlFile newXmlFile = new xmlFile();


            XmlElement shortName = newXmlFile.doc.CreateElement(string.Empty, "SHORT-NAME", string.Empty);
            XmlText text1 = newXmlFile.doc.CreateTextNode("ISIGNALS");
            shortName.AppendChild(text1);
            newXmlFile.nodePrincipal.AppendChild(shortName);

            XmlElement elements = newXmlFile.doc.CreateElement(string.Empty, "ELEMENTS", string.Empty);
            newXmlFile.nodePrincipal.AppendChild(elements);


            for (int i = 2; i <= lastUsedRow; i++)
            {
                List<string> lineElement = new List<string>();

                for (int j = 1; j <= lastUsedColumn; j++)
                {
                    lineElement.Add(sheetOne.Cells[i, j].Value.ToString());

                }
                newXmlFile.WriteSignalNode(lineElement, elements, newXmlFile.doc);
            }



            int lastUsedRowSheetTwo = lastsheetTwo.Row;
            int lastUsedColumnSheetTwo = lastsheetTwo.Column;



            XmlElement shortNameECU = newXmlFile.doc.CreateElement(string.Empty, "SHORT-NAME", string.Empty);
            XmlText text2 = newXmlFile.doc.CreateTextNode("ECUINSTANCES");
            shortNameECU.AppendChild(text2);
            newXmlFile.nodePrincipal2.AppendChild(shortNameECU);

            XmlElement elementsEcu = newXmlFile.doc.CreateElement(string.Empty, "ELEMENTS", string.Empty);
            newXmlFile.nodePrincipal2.AppendChild(elementsEcu);

            for (int i = 2; i <= lastUsedRowSheetTwo; i++)
            {
                List<string> lineElementECU = new List<string>();

                for (int j = 1; j <= lastUsedColumnSheetTwo; j++)
                {
                    lineElementECU.Add(sheetTwo.Cells[i, j].Value.ToString());

                }
                newXmlFile.WriteECU(lineElementECU, elementsEcu, newXmlFile.doc);
            }


            newXmlFile.doc.Save("D:\\newFile.xml");
            textBox2.Text = "Fisierul XML a fost creat cu succes!";
        }

        private void radioButton2_XMLToExcel()
        {
            XmlDataDocument xmldoc = new XmlDataDocument();
            XmlNodeList xmlnodeSignal;
            XmlNodeList xmlnodeEcuInstance;
        
            string str = null;
            FileStream fs = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read);
            xmldoc.Load(fs);

            DisplayInfo newDisplayInfo = new DisplayInfo();

            newDisplayInfo.comboBoxAllWriter("Fisierul a fost deschis cu succes!", textBox2, comboBox1, true);

            xmlnodeSignal = xmldoc.GetElementsByTagName("I-SIGNAL");
            xmlnodeEcuInstance = xmldoc.GetElementsByTagName("ECU-INSTANCE");

            ExcelFile newFile = new ExcelFile();

            for (int i = 0; i < xmlnodeSignal.Count; i++)
                {
                str = xmlnodeSignal[i].ChildNodes.Item(0).InnerText.Trim() + " " + xmlnodeSignal[i].ChildNodes.Item(1).InnerText.Trim() + " " + xmlnodeSignal[i].ChildNodes.Item(2).InnerText.Trim() + " " + xmlnodeSignal[i].ChildNodes.Item(3).InnerText.Trim() + " " + xmlnodeSignal[i].ChildNodes.Item(4).InnerText.Trim();
                newFile.writeSignalSheet(i, str, newFile.xlWorkSheetSignal);
                }

            for (int i = 0; i < xmlnodeEcuInstance.Count; i++)
            {
                str = xmlnodeEcuInstance[i].ChildNodes.Item(0).InnerText.Trim() + " " + xmlnodeEcuInstance[i].ChildNodes.Item(2).InnerText.Trim() + " " + xmlnodeEcuInstance[i].ChildNodes.Item(3).InnerText.Trim() + " " + xmlnodeEcuInstance[i].ChildNodes.Item(4).InnerText.Trim() + " " + xmlnodeEcuInstance[i].ChildNodes.Item(5).InnerText.Trim() + " " + xmlnodeEcuInstance[i].ChildNodes.Item(8).InnerText.Trim() + " " + xmlnodeEcuInstance[i].ChildNodes.Item(9).InnerText.Trim();
                newFile.writeEcuInstanceSheet(i, str, newFile.xlWorkSheetEcuInstance);
            }
       
            string pathXlsx = "d:\\csharpppp-Excel.xlsx";
            string pathXls = "d:\\csharpppp-Excel.xls";
            newFile.saveExcelFile(pathXls, pathXlsx, newFile.xlWorkBook, checkBox1, checkBox2, newFile.misValue, textBox2);


            newDisplayInfo.comboBoxAllWriter("Fisierul a fost creat cu succes!", textBox2, comboBox1, false);
            newDisplayInfo.comboBoxAllWriter("Fisierul a fost creasdaasdasdasdasdasdasdasdat cu succes!", textBox2, comboBox1, true);

            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(radioButton1.Checked is true)
            {
                string fineName = "D:\\newFile.xml";
                System.Diagnostics.Process.Start("notepad.exe", fineName);

            }
            else if(radioButton2.Checked is true)
            {
                string pathXlsx = "d:\\csharppp-Excel.xlsx";
                string pathXls = "d:\\csharppp-Excel.xls";

                var xlApp = new Excel.Application();
                xlApp.Visible = true;
                var xlApp1 = new Excel.Application();
                xlApp1.Visible = true;

                if (checkBox1.Checked is true && checkBox2.Checked is true)
                {
                    Excel.Workbook xlWorkBook  = xlApp.Workbooks.Open(pathXlsx);
                    Excel.Workbook xlWorkBook1 = xlApp1.Workbooks.Open(pathXls);
                }
                else if(checkBox1.Checked is true && checkBox2.Checked is false)
                {
                    xlApp.Visible = false;
                    Excel.Workbook xlWorkBook1 = xlApp1.Workbooks.Open(pathXls);
                }
                else if (checkBox1.Checked is false && checkBox2.Checked is true)
                {
                    xlApp1.Visible = false;
                    Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(pathXlsx);
                }
                else if (checkBox1.Checked is false && checkBox2.Checked is false)
                {
                    xlApp.Visible = false;
                    xlApp1.Visible = false;
                    DisplayInfo newDisplayInfo = new DisplayInfo();
                    newDisplayInfo.comboBoxAllWriter("Nu exista fisiere de deschis!", textBox2, comboBox1, true);
                }

            }
        }
    }
}
