using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class MyForm : Form
    {
        private ExcelHandler handler;
        public MyForm()
        {
            InitializeComponent();
            this.Text = "Application";
        }
        private void OpenFileButtonClick(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.Title = "Select a Excel File";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePath.Text = openFileDialog1.FileName;
                handler = new ExcelHandler(filePath.Text, 1);
                MessageBox.Show($"Handler: {handler?.ToString()}\nApplication: {handler.GetApplication.Name.ToString()}\nWorkbook: {handler?.GetWorkbook?.ToString()}");
                //catalog = new BindingList<Dictionary<String, String>.ValueCollection>();
                //reportList = new List<Report>();
                //for(int i = 3; i < 10; i++)
                //{
                //    System.Array values = (System.Array)handler.GetWorksheet.get_Range("A" + i.ToString(), "AJ" + i.ToString()).Cells.Value;
                //    Report report = new Report(handler.GetCell(i, 1)?.ToString(),
                //        handler.GetCell(i, 2)?.ToString(),
                //        handler.GetCell(i, 3)?.ToString());
                //    //Report report = new Report();
                //    //for(int j = 1; j < handler.GetWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column; j++)
                //    //{
                //    //    report.Add(handler.GetCell(1, j).ToString(), handler.GetCell(i, j)?.ToString());
                //    //}
                //    //catalog.Add(report.Values);
                //    reportList.Add(report);
                //}
                //var catalogDataSource = new BindingSource(catalog, null);
                //dataGridView1.DataSource = catalogDataSource;
            }
        }
        private void openFileDialog1_FileOk(object sender, EventArgs e)
        {

        }

        private void RunClick(object sender, EventArgs e)
        {
            
            this.Text = "Running...";
            try
            {
                for(int i = 3; i < handler.GetRange.Rows.Count; i++)
                {
                    //handler.RemoveDuplicateInstructions(i, (int) ExcelHandler.COLUMN.WORK_INSTRUCTIONS, (int)ExcelHandler.COLUMN.NOTES);
                    //handler.RemoveDuplicateInstructions(i, (int) ExcelHandler.COLUMN.FOLDER_LOCATION, (int)ExcelHandler.COLUMN.ERS_LOCATION);
                    handler.UpdateSourceDepartment(i);
                    handler.UpdateOtherDepartment(i);
                    
                    //handler.UpdateReportType(i);
                    handler.UpdateNotes(i);
                    //handler.UpdateReportName(i);
                    //updateDepartment(i);
                }
                Save();
            } catch (Exception exception)
            {
                handler.Log(exception.ToString());
                MessageBox.Show($"{exception.GetType().ToString()} {exception.ToString()}");
            }
            this.Text = "Application";
            handler.Log("Program has finished.");
            handler.Close();
        }
        private void Save()
        {
            saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Title = "Save File As";
            saveFileDialog1.DefaultExt = "xlsx";
            //saveFileDialog1.CheckFileExists = true;
            //saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.InitialDirectory = @"C:\Users\gthao\Documents\";
            saveFileDialog1.RestoreDirectory = true;
            //saveFileDialog1.ShowDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath.Text = saveFileDialog1.FileName;
                handler.SaveExcel(filePath.Text);
            }

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
