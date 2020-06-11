using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using Gizmox.CSharp;
using System.Linq.Expressions;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
            //            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
            //            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {

            string sql = "";

            UpdateStatus("Updating TimeEntry...", 1, 7);
            sql = "update TimeEntry set BillableFlag = 'N' " +
                "where MatterSysNbr in (SELECT MatSysNbr FROM Matter where MatBillAgreeCode in ('N', 'B'))  and entrysource = 'JurisGo'";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating TimeBatch...", 2, 7);
            //timebatchdetail
            sql = "  update timebatchdetail set TBDBillableFlg = 'N' " +
                    " where timebatchdetail.tbdid in (" +
                    " select timebatchdetail.tbdid FROM  [TimeEntry] " +
                    " inner join TimeEntryLink on timeentrylink.entryid = timeentry.entryid " +
                    " inner join timebatchdetail on timeentrylink.tbdid = timebatchdetail.tbdid " +
                    " inner join matter on matsysnbr = mattersysnbr " +
                    " where MatBillAgreeCode in ('N', 'B')  and entrysource = 'JurisGo' and tbdrectype in (1,2)) ";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating UnbilledTime...", 3, 7);
            //unbilledtime
            sql = "update unbilledtime set UTBillableFlg = 'N' "
                    + "from unbilledtime "
                    + "inner join timeentrylink aa on utid = aa.tbdid "
                    + "inner join timeentry bb on aa.entryid = bb.entryid "
                    + "where MatterSysNbr in (SELECT MatSysNbr FROM Matter where MatBillAgreeCode in ('N', 'B'))  and entrysource = 'JurisGo'";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating FeeSumByPrd...", 5, 7);
            //feesumbyprd
            sql = " update FeeSumByPrd set [FSPNonBilHrsEntered] = hhb.FSPBilHrsEntered,[FSPBilHrsEntered] = hhb.FSPBilHrsEntered " +
            " ,[FSPFeeEnteredActualValue] = hhb.FSPFeeEnteredActualValue from " +
            " ( " +
            " select matter, yr, prd, tkpr, task, act, sum(FSPBilHrsEntered) as FSPBilHrsEntered, sum(FSPNonBilHrsEntered) as FSPNonBilHrsEntered, " +
            " sum(FSPFeeEnteredActualValue) as FSPFeeEnteredActualValue " +
            " from ( " +
            " SELECT  [UTMatter] as matter,[UTPrdYear] as yr,[UTPrdNbr] as prd,[UTTkpr]  as tkpr,[UTTaskCd] as task ,[UTActivityCd] as act " +
            " 	  ,sum(case when UTBillableFlg = 'Y' then [UTHoursToBill] else 0 end) as FSPBilHrsEntered " +
            " 	  ,sum(case when UTBillableFlg = 'N' then [UTActualHrsWrk] else 0 end) as FSPNonBilHrsEntered " +
            "       ,sum([UTAmount]) as FSPFeeEnteredActualValue " +
            "   FROM [UnbilledTime] " +
            " group by [UTMatter],[UTPrdYear],[UTPrdNbr] ,[UTTkpr] ,[UTTaskCd] ,[UTActivityCd] " +
            " union all " +
            " SELECT  [BTMatter] as matter,[BTPrdYear] as yr,[BTPrdNbr] as prd,[BTWrkTkpr] as tkpr ,[BTTaskCd] as task ,[BTActivityCd] as act " +
            "       ,sum(case when BTBillableFlg = 'Y' then [BTHoursToBill] else 0 end) as FSPBilHrsEntered " +
            " 	  ,sum(case when BTBillableFlg = 'N' then [BTActualHrsWrk] else 0 end) as FSPNonBilHrsEntered " +
            "       ,sum([BTAmount]) as FSPFeeEnteredActualValue " +
            "   FROM [BilledTime] " +
            "   group by [BTMatter] ,[BTPrdYear],[BTPrdNbr] ,[BTWrkTkpr] ,[BTTaskCd] ,[BTActivityCd] ) llk " +
            "   where yr >= 2019 and matter in (SELECT MatSysNbr FROM Matter where MatBillAgreeCode in ('N', 'B')) " +
            "   group by matter, yr, prd, tkpr, task, act) hhb " +
            "   where [FSPMatter]= hhb.matter and [FSPPrdYear] = hhb.yr and [FSPPrdNbr] = hhb.prd and [FSPTkpr] = hhb.tkpr  " +
            "   and [FSPTaskCd] = hhb.task and [FSPActivityCd] = hhb.act ";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating FeeSumITD...", 6, 7);
            //feesumitd
            sql = " update FeeSumITD set [FSINonBilHrsEntered] = hhb.FSPBilHrsEntered,[FSIBilHrsEntered] = hhb.FSPBilHrsEntered " +
            " ,[FSIFeeEnteredActualValue] = hhb.FSPFeeEnteredActualValue from " +
            " ( " +
            " select matter,  tkpr, sum(FSPBilHrsEntered) as FSPBilHrsEntered, sum(FSPNonBilHrsEntered) as FSPNonBilHrsEntered, " +
            " sum(FSPFeeEnteredActualValue) as FSPFeeEnteredActualValue " +
            " from ( " +
            " SELECT  [UTMatter] as matter,[UTTkpr]  as tkpr " +
            " 	  ,sum(case when UTBillableFlg = 'Y' then [UTHoursToBill] else 0 end) as FSPBilHrsEntered " +
            " 	  ,sum(case when UTBillableFlg = 'N' then [UTActualHrsWrk] else 0 end) as FSPNonBilHrsEntered " +
            "       ,sum([UTAmount]) as FSPFeeEnteredActualValue " +
            "     , sum([UTHoursToBill]) as WIPHrs " +
            "     , sum([UTAmount]) as WIPBal " +
            "   FROM [UnbilledTime] " +
            " group by [UTMatter],[UTTkpr]  " +
            " union all " +
            " SELECT  [BTMatter] as matter, [BTWrkTkpr] as tkpr  " +
            "       ,sum(case when BTBillableFlg = 'Y' then [BTHoursToBill] else 0 end) as FSPBilHrsEntered " +
            " 	  ,sum(case when BTBillableFlg = 'N' then [BTActualHrsWrk] else 0 end) as FSPNonBilHrsEntered " +
            "       ,sum([BTAmount]) as FSPFeeEnteredActualValue " +
            "      , 0 as WIPHrs, 0 as WIPBal " +
            "   FROM [BilledTime] " +
            "   group by [BTMatter] ,[BTPrdYear],[BTPrdNbr] ,[BTWrkTkpr] ,[BTTaskCd] ,[BTActivityCd] ) llk " +
            "   where matter in (SELECT MatSysNbr FROM Matter where MatBillAgreeCode in ('N', 'B')) " +
            "   group by matter, tkpr) hhb " +
            "   where [FSIMatter]= hhb.matter  and [FSITkpr] = hhb.tkpr  ";

            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("All tables updated.", 7, 7);

            MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step / steps) * 100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }



        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            try
            {
                string SQL = " DROP TABLE ##TempGo";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }
            catch(Exception ex1)
            { 
            }
            try
            {
                string SQL = " DROP TABLE ##TempGo2";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }
            catch (Exception ex1)
            {
            }

            System.Environment.Exit(0);
        }

        private void button2_Click(object sender, EventArgs e) //add number of time entries on each bill/matter
        {
            string SQL = "select distinct dbo.jfn_FormatClientCode(clicode) as ClientCode, dbo.jfn_FormatMatterCode(MatCode) as MatterCode, "+
" arbillnbr as BillNo, convert(varchar, arbilldate, 101) as BillDate, count(bb.EntryID) as TotalEntries from billedtime " +
"inner join timeentrylink aa on btid = aa.tbdid " +
"inner join timeentry bb on aa.entryid = bb.entryid " +
"inner join ARBill on arbillnbr = btbillnbr " + 
"Inner join matter on matsysnbr = btmatter " +
"inner join Client on clisysnbr = matclinbr " +
"where MatterSysNbr in (SELECT MatSysNbr FROM Matter where MatBillAgreeCode in ('N', 'B')) and entrysource = 'JurisGo' " +
" and (billedtime.btbillnbr not in (select btbillnbr from ##TempGo where ##TempGo.btbillnbr in (select btbillnbr from ##TempGo2) and ##TempGo.BillableFlag = 'N')) " +
" group by clicode, matcode, arbillnbr, arbilldate";

            DataSet ff = _jurisUtility.RecordsetFromSQL(SQL);

            ReportDisplay rp = new ReportDisplay(ff);
            rp.ShowDialog();


        }
    }
}
