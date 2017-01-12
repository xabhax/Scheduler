using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Scheduler
{
    public partial class Form1 : Form
    {
        private readonly string _techfile = Environment.CurrentDirectory + "\\techs.dat";
        private readonly string _satfile = Environment.CurrentDirectory + "\\saturday.dat";

        [DllImport("gdi32.dll", ExactSpelling = true, CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool BitBlt(IntPtr pHdc, int iX, int iY, int iWidth, int iHeight, IntPtr pHdcSource, int iXSource, int iYSource, System.Int32 dw);
        private const int SRC = 0xCC0020;

        public Form1()
        {
            InitializeComponent();
            SetupDataGrid();
            FillDataGrid();
        }

        private void SetupDataGrid()
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
            string[] values;
            var hours = "7:00 - 5:00";
            var sathours = "7:00 - 5:00";
            var blank = "   ";

            var sr = new StreamReader(Environment.CurrentDirectory + "\\techs.dat");
            var sf = new StreamReader(Environment.CurrentDirectory + "\\saturday.dat");

            var saturdayteam = sf.ReadLine();

            sf.Close();

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

                if (values[0] == "none") dataGridView1.Rows.Add("  ");
                else dataGridView1.Rows.Add(values[0], values[1], values[2], values[3], values[4], values[5]);
            }
            sr.Close();

            dataGridView1.CurrentCell = null;

            if (saturdayteam == "Team 1")
            {
                dataGridView1[6, 1].Value = sathours;
                dataGridView1[6, 2].Value = sathours;
                dataGridView1[6, 3].Value = sathours;
                dataGridView1[6, 4].Value = sathours;
                dataGridView1[6, 5].Value = sathours;
            }

            if (saturdayteam == "Team 2")
            {
                dataGridView1[6, 7].Value = sathours;
                dataGridView1[6, 8].Value = sathours;
                dataGridView1[6, 9].Value = sathours;
                dataGridView1[6, 10].Value = sathours;
                dataGridView1[6, 11].Value = sathours;
            }

            if (saturdayteam == "Team 3")
            {
                dataGridView1[6, 13].Value = sathours;
                dataGridView1[6, 14].Value = sathours;
                dataGridView1[6, 15].Value = sathours;
                dataGridView1[6, 16].Value = sathours;
                dataGridView1[6, 17].Value = sathours;
            }
        }

        private void StepSchedule(object sender, EventArgs e)
        {
            List<string> oldTechData = File.ReadAllLines(_techfile).ToList();
            List<string> newTechData = new List<string>();

            foreach (string line in oldTechData)
            {
                string[] parsedData = line.Split('|');

                if (parsedData[1] == "2")
                {
                    newTechData.Add(parsedData[0] + "|" + "2" + "|" + "2" + "|" + "2" + "|" + "2" + "|" + "2" + "|");
                }
                if (parsedData[1] == "1")
                {
                    newTechData.Add(parsedData[0] + "|" + "0" + "|" + "1" + "|" + "0" + "|" + "0" + "|" + "0" + "|");
                }
                if (parsedData[2] == "1")
                {
                    newTechData.Add(parsedData[0] + "|" + "0" + "|" + "0" + "|" + "1" + "|" + "0" + "|" + "0" + "|");
                }
                if (parsedData[3] == "1")
                {
                    newTechData.Add(parsedData[0] + "|" + "0" + "|" + "0" + "|" + "0" + "|" + "1" + "|" + "0" + "|");
                }
                if (parsedData[4] == "1")
                {
                    newTechData.Add(parsedData[0] + "|" + "0" + "|" + "0" + "|" + "0" + "|" + "0" + "|" + "1" + "|");
                }
                if (parsedData[5] == "1")
                {
                    newTechData.Add(parsedData[0] + "|" + "1" + "|" + "0" + "|" + "0" + "|" + "0" + "|" + "0" + "|");
                }
            }

            if (File.Exists(_techfile))
            {
                File.Delete(_techfile);
            }

            File.WriteAllLines(_techfile, newTechData);

            var sf = new StreamReader(Environment.CurrentDirectory + "\\saturday.dat");

            var saturdayteam = sf.ReadLine();
            sf.Close();

            if (saturdayteam == "Team 1") { File.WriteAllText(_satfile, "Team 2"); }
            if (saturdayteam == "Team 2") { File.WriteAllText(_satfile, "Team 3"); }
            if (saturdayteam == "Team 3") { File.WriteAllText(_satfile, "Team 1"); }

            Application.Restart();
        }

        private void ExportToExcel(object sender, EventArgs e)
        {
            //Answer to a StackOverflow Question
            DataGridView dg = dataGridView1;

            dg.Refresh();
            dg.Select();

            Graphics g = dg.CreateGraphics();
            Bitmap ibitMap = new Bitmap(dg.ClientSize.Width, dg.ClientSize.Height, g);
            Graphics iBitMap_gr = Graphics.FromImage(ibitMap);
            IntPtr iBitMap_hdc = iBitMap_gr.GetHdc();
            IntPtr me_hdc = g.GetHdc();

            BitBlt(iBitMap_hdc, 0, 0, dg.ClientSize.Width, dg.ClientSize.Height, me_hdc, 0, 0, SRC);
            g.ReleaseHdc(me_hdc);
            iBitMap_gr.ReleaseHdc(iBitMap_hdc);

            ibitMap.Save(Environment.CurrentDirectory + "\\schedule.bmp", ImageFormat.Bmp);
        }

        private void ManageTechs(object sender, EventArgs e)
        {
            var ManageTechsForm = new ManageTechs();
            ManageTechsForm.ShowDialog();
        }
    }
}