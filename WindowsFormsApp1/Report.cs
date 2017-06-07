using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class Report
    {
        private string reportName, businessContact, businessOwner;
        private DateTime dueDate1, dueDate2, dueDate3, dueDate4;
        private DateTime dayDue;
        public enum Frequency
        {
            Weekly, Biweekly, Monthly, Quarterly, Semianually, Annually
        };
        private string deliveryFunction;
        private string workInstructions, notes;
        private int daysAfterQuarter;
        private string folderLocation;
        private string reportType, runWith, deliveryMethod, deliverTo;
        private DateTime effectiveDate, terminationDate;
        private string groupName;
        private int wwGroupID;
        private string state;
        private string reportPath;
        private bool otherDepartment;
        private enum SOURCE_DEPARTMENT
        {
            BI_REPORTING, CREDENTIALING, QUALITY, CLAIMS_AND_GREVIANCES, FRAUD, CUSTOMER_SERVICE, UM, AE, ED, CLIENT_ENGAGEMENT
        };
        private string sourceDepartment;
        private bool qualityIndicator;
        private int errStatus;
        private DateTime dateAdded;
        private DateTime systemRefreshDate;
        private int legacyReportID, legacyReportID_R2;
        private string ERS_ReportName;
        private string otherReportLocation;
        private string otherReportName;
        public Report() { }
        public Report(String reportName, String businessContact, String businessOwner)
        {
            this.reportName = reportName;
            this.businessContact = businessContact;
            this.businessOwner = businessOwner;
        }
        public String ReportName
        {
            get { return reportName; }
            set { reportName = value; }
        }
        public String BusinessOwner
        {
            get { return businessOwner; }
            set { businessOwner = value; }
        }
        public String BusinessContact
        {
            get { return businessContact; }
            set { businessContact = value; }
        }
        public DateTime DueDate1
        {
            get { return dueDate1; }
            set { dueDate1 = value; }
        }
        //private Dictionary<String, String> fields;
        //public void Add(String key, String val) => fields.Add(key, val);
    }
}
