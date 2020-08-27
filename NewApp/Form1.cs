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
            MessageBox.Show("Print 1 Excel to XML");
        }
        private void radioButton2_XMLToExcel()
        {
            XmlDataDocument xmldoc = new XmlDataDocument();
            XmlNodeList xmlnodeSignal;
            XmlNodeList xmlnodeEcuInstance;
        
            string str = null;
            FileStream fs = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read);
            xmldoc.Load(fs);
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
       
            string pathXlsx = "d:\\csharppp-Excel.xlsx";
            string pathXls = "d:\\csharppp-Excel.xls";
            newFile.saveExcelFile(pathXls, pathXlsx, newFile.xlWorkBook, checkBox1, checkBox2, newFile.misValue, textBox2);

            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);

        }
    }
}
