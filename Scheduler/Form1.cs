using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.ExcelApi.Tools.Utils;

namespace Scheduler
{
    public partial class Form1 : Form
    {
        public string filename = Environment.CurrentDirectory + "\\export.xls";
        public string techfile = Environment.CurrentDirectory + "\\techs.dat";

        public Form1()
        {
            InitializeComponent();
            SetupDataGrid();
            FillDataGrid();
        }

        public void SetupDataGrid()
        {
            // Create an unbound DataGridView by declaring a column count.
            dataGridView1.ColumnCount = 7;
            dataGridView1.ColumnHeadersVisible = true;

            // Set the column header style.
            var columnHeaderStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.Beige;
            columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            // Set the column header names.
            dataGridView1.Columns[0].Name = " ";
            dataGridView1.Columns[1].Name = "Mon";
            dataGridView1.Columns[2].Name = "Tue";
            dataGridView1.Columns[3].Name = "Wed";
            dataGridView1.Columns[4].Name = "Thu";
            dataGridView1.Columns[5].Name = "Fri";
            dataGridView1.Columns[6].Name = "Sat";
        }

        private void FillDataGrid()
        {
            var sr = new StreamReader(Environment.CurrentDirectory + "\\techs.dat");
            string[] values;
            var hours = "7:00 - 5:00";
            var blank = "   ";

            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                values = line.Split('|');

                if (values[1] == "1") values[1] = "Off";
                else if (values[1] == "2") values[1] = blank;
                else values[1] = hours;

                if (values[2] == "1") values[2] = "Off";
                else if (values[2] == "2") values[2] = blank;
                else values[2] = hours;

                if (values[3] == "1") values[3] = "Off";
                else if (values[3] == "2") values[3] = blank;
                else values[3] = hours;

                if (values[4] == "1") values[4] = "Off";
                else if (values[4] == "2") values[4] = blank;
                else values[4] = hours;

                if (values[5] == "1") values[5] = "Off";
                else if (values[5] == "2") values[5] = blank;
                else values[5] = hours;

                //if (values[6] == "1") { values[6] = "Off"; } else if (values[6] == "2") { values[6] = blank; } else { values[6] = hours; }

                if (values[0] == "none") dataGridView1.Rows.Add("  ");
                else dataGridView1.Rows.Add(values[0], values[1], values[2], values[3], values[4], values[5]);
            }
            sr.Close();
        } 

        private void StepSchedule(object sender, EventArgs e)
        {
            List<string> oldTechData = File.ReadAllLines(techfile).ToList();
            List<string> newTechData = new List<string>();

            foreach (string line in oldTechData)
            {
                string[] parsedData = line.Split('|');

                if (parsedData[1] == "2") { newTechData.Add(parsedData[0] + "|" + "2" + "|" + "2" + "|" + "2" + "|" + "2" + "|" + "2" + "|"); }
                if (parsedData[1] == "1") { newTechData.Add(parsedData[0] + "|" + "0" + "|" + "1" + "|" + "0" + "|" + "0" + "|" + "0" + "|"); }
                if (parsedData[2] == "1") { newTechData.Add(parsedData[0] + "|" + "0" + "|" + "0" + "|" + "1" + "|" + "0" + "|" + "0" + "|"); }
                if (parsedData[3] == "1") { newTechData.Add(parsedData[0] + "|" + "0" + "|" + "0" + "|" + "0" + "|" + "1" + "|" + "0" + "|"); }
                if (parsedData[4] == "1") { newTechData.Add(parsedData[0] + "|" + "0" + "|" + "0" + "|" + "0" + "|" + "0" + "|" + "1" + "|"); }
                if (parsedData[5] == "1") { newTechData.Add(parsedData[0] + "|" + "1" + "|" + "0" + "|" + "0" + "|" + "0" + "|" + "0" + "|"); }
            }

            if (File.Exists(techfile))
            {
                File.Delete(techfile);
            }

            File.WriteAllLines(techfile, newTechData);

            Application.Restart();
        }

        private void ExportToExcel(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            int i = 0;
            int j = 0;

            for (i = 0; i <= dataGridView1.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGridView1[j, i];
                    Excel.Worksheet.xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }

            xlWorkBook.SaveAs("csharp.net-informations.xls", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file c:\\csharp.net-informations.xls");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void ManageTechs(object sender, EventArgs e)
        {

        }
    }
}