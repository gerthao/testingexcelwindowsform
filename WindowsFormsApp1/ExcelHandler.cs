using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    class ExcelHandler
    {
        private Excel.Application application;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        private Excel.Range range;
        private string filePath;


        private static int entryCountUpdated = 0;
        private static int duplicateEntriesRemoved = 0;
        private static int notesUpdated = 0;
        private const string PARAMETERS = "PARAMETERS =>";
        private const string ERS = "ERS =>";
        private const string REPORT = "REPORT =>";
        private const string QUEST = "QUEST =>";
        private const string LINK = "LINK =>";
        private LogFile log;
        private String[] departments =
        {
            "BI Reporting", "C&G", "Fraud", "Quality", "Credentialing", "Customer Service", "AE"
        };
        public enum COLUMN
        {
            REPORT_NAME = 1,
            BUSINESS_CONTACT = 2,
            BUSINESS_OWNER = 3,
            DUE_DATE_1 = 4,
            DUE_DATE_2 = 5,
            DUE_DATE_3 = 6,
            DUE_DATE_4 = 7,
            DAY_DUE = 8,
            FREQUENCY = 9,
            DELIVERY_FUNCTION = 10,
            WORK_INSTRUCTIONS = 11,
            NOTES = 12,
            DAYS_AFTER_QUARTER = 13,
            FOLDER_LOCATION = 14,
            REPORT_TYPE = 15,
            RUN_WITH = 16,
            DELIVERY_METHOD = 17,
            DELIVERY_TO = 18,
            EFFECTIVE_DATE = 19,
            TERMINATION_DATE = 20,
            GROUP_NAME = 21,
            WW_GROUP_ID = 22,
            STATE = 23,
            REPORT_PATH = 24,
            OTHER_DEPARTMENT = 25,
            SOURCE_DEPARTMENT = 26,
            QUALITY_INDICATOR = 27,
            ERS_LOCATION = 28,
            ERR_STATUS = 29,
            DATE_ADDED = 30,
            SYSTEM_REFERESH_DATE = 31,
            LEGACY_REPORT_ID = 32,
            LEGACY_REPORT_ID_R2 = 33,
            ERS_REPORT_NAME = 34,
            OTHER_REPORT_LOCATION = 35,
            OTHER_REPORT_NAME = 36
        };

        public ExcelHandler(String pathName, int sheetNumber)
        {
            try
            {
                log = new LogFile();
                application = new Excel.Application();
                filePath = pathName;
                workbook = application.Workbooks.Open(filePath);
                worksheet = workbook.Sheets[sheetNumber];
                range = worksheet.UsedRange;
                
            } catch (Exception e)
            {
                MessageBox.Show("Exception: " + e.Message);
            }
            
        }
        public string GetCell(int row, int col) => ((range.Cells[row, col] as Excel.Range).Value2)?.ToString();
        public void SetCell(int row, int col, String value) => worksheet.Cells[row, col] = value;
        public string FilePath
        {
            get { return filePath; }
            set { filePath = value; }
        }
        public Excel.Range GetRange => range;
        public Excel.Worksheet SetCurrentWorksheet(int index) => worksheet = workbook.Sheets[index];
        public Excel.Worksheet GetWorksheet => worksheet;
        public Excel.Workbook GetWorkbook => workbook;
        public Excel.Application GetApplication => application;
        public void Close()
        {
            workbook?.Close(true, null, null);
            application?.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(application);
            worksheet = null;
            workbook = null;
            application = null;
            range = null;
            log?.Close();
            log = null;
        }
        public void Move(int row, int col, string val)
        {
            
        }
        public void UpdateReportType(int row)
        {
            string workInstructions = GetCell(row, (int)COLUMN.WORK_INSTRUCTIONS);
            string reportType = null, runWith = null;
            if (workInstructions == null) return;
            if (workInstructions.Contains(QUEST))
            {
                reportType = "Quest Analytics Report";
                runWith = "Quest Analytics Suite 2016";
            } else if (workInstructions.Contains(LINK) || workInstructions.Contains(ERS))
            {
                reportType = "ERS Report";
                runWith = "Enterprise Reporting";
            }
            SetCell(row, (int)COLUMN.REPORT_TYPE, reportType);
            SetCell(row, (int)COLUMN.RUN_WITH, runWith);
        }

        public void UpdateReportName(int row)
        {
            string reportName = GetCell(row, (int)COLUMN.REPORT_NAME);
            string groupName = GetCell(row, (int)COLUMN.GROUP_NAME);
            string ERS_ReportName = GetCell(row, (int)COLUMN.ERS_REPORT_NAME);
            string otherReportName = GetCell(row, (int)COLUMN.OTHER_REPORT_NAME);
            try
            {
                if (reportName.Equals(ERS_ReportName))
                {
                    reportName = $"{groupName} {ERS_ReportName}";
                }
                else if (reportName.Equals(otherReportName))
                {
                    reportName = $"{groupName} {otherReportName}";
                }
                SetCell(row, (int)COLUMN.REPORT_NAME, reportName);
            }
            catch (NullReferenceException ex)
            {
                log.Log(ex.Message);
                log.Log(ex.Data.ToString());
            }
        }
        public void UpdateOtherDepartment(int row)
        {
            string sourceDepartment = GetCell(row, (int)COLUMN.SOURCE_DEPARTMENT);
            string otherDepartment;
            if (sourceDepartment == null)
            {
                return;
            }
            if (sourceDepartment.Equals(departments[0]))
            {
                otherDepartment = "N";
            }
            else otherDepartment = "Y";
            SetCell(row, (int)COLUMN.OTHER_DEPARTMENT, otherDepartment);
            entryCountUpdated++;
        }
        public void UpdateNotes(int row)
        {
            string workInstructions = GetCell(row, (int)COLUMN.WORK_INSTRUCTIONS);
            string notes = GetCell(row, (int)COLUMN.NOTES);
            string temp = "";
            if (workInstructions == null || notes == null) return;
            if (workInstructions.Contains(PARAMETERS))
            {
                temp = workInstructions.Substring(workInstructions.IndexOf(PARAMETERS)).Trim();
                workInstructions = workInstructions.Substring(0, workInstructions.IndexOf(PARAMETERS)).Trim();
            }
            else
            {
                log.Log($"No Parameters found from {workInstructions}");
                return;
            }
            if (notes == null || notes.Equals("REMOVED!"))
            {
                SetCell(row, (int)COLUMN.NOTES, temp);
                notesUpdated++;
            }
            else
            {
                SetCell(row, (int)COLUMN.NOTES, $"{temp} {notes}");
            }
            SetCell(row, (int)COLUMN.WORK_INSTRUCTIONS, workInstructions.Equals("") ? "" : workInstructions);
            log.Log($"Moving to Notes: {temp}");
        }
        public void RemoveDuplicateInstructions(int row, int col_1, int col_2)
        {
            string firstCell = this.GetCell(row, col_1);
            string secondCell = this.GetCell(row, col_2);
            log.Log($"{firstCell} ? {secondCell}");
            if (firstCell == null || secondCell == null) return;
            if (firstCell.Equals(secondCell))
            {
                SetCell(row, col_2, "REMOVED!");
                duplicateEntriesRemoved++;
            }
        }
        public void UpdateSourceDepartment(int row)
        {
            string foundDepartment = GetCell(row, (int)COLUMN.WORK_INSTRUCTIONS);
            //foundDepartment = foundDepartment.Contains(REPORT) ? foundDepartment : GetCell(row, (int)COLUMN.WORK_INSTRUCTIONS);
            string sourceDepartment;
            if (foundDepartment == null)
            {
                //log.Log("No Department Found, No Change Made");
                //return;
                foundDepartment = GetCell(row, (int)COLUMN.ERS_LOCATION);
                if (foundDepartment == null) return;
            }
            if (foundDepartment.Contains(REPORT))
            {
                sourceDepartment = foundDepartment.Substring(REPORT.Length).Trim();
                if (sourceDepartment.Contains(";")) sourceDepartment = sourceDepartment.Split(';')[0].Trim();
                //SetCell(row, (int)COLUMN.WORK_INSTRUCTIONS, $"{this.GetCell(row, (int)COLUMN.WORK_INSTRUCTIONS)};  REPORT => {sourceDepartment};");
            }
            else sourceDepartment = "BI Reporting";
            SetCell(row, (int)COLUMN.SOURCE_DEPARTMENT, sourceDepartment);
            log.Log($"Found: {foundDepartment}, Source Department Updated To: {sourceDepartment}");
        }
        public void Log(string message) => log?.Log(message);

        public void SaveLog()
        {
            
        }
        public void SaveLog(string path)
        {
            
        }
        public void SaveExcel()
        {
            workbook.Save();
        }
        public void SaveExcel(string path)
        {
            workbook.SaveAs(path);
        }
    }
}
