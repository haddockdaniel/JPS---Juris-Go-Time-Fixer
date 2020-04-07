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
            string sql;
            string matsys;
            string period;
            string yr;
            string hours;
            string h2b;
            string amount;
            string std;
            string tkpr;
            string taskcd;
            string actcd;

            string SQLFSP = "select mattersysnbr,periodnumber, periodyear,timekeepersysnbr, isnull(TaskCode,'') as TaskCd, isnull(activitycode,'') as ActivityCd, sum(actualhourswork) as Hours,sum(hourstobill) as H2B, sum(amount) as Amount, sum(AmountAtStandardRate) as StdAmt from timeentry  where entrysource='JurisGo' and entrystatus>=7 and billableflag='Y' and MatterSysNbr in (select matsysnbr from matter where matbillagreecode in ('N','B')) " +
                " group by mattersysnbr,periodnumber, periodyear,timekeepersysnbr, isnull(TaskCode,''), isnull(activitycode,'') ";
            DataSet myRSFSP = _jurisUtility.RecordsetFromSQL(SQLFSP);



            UpdateStatus("Updating Fee Sum by Period and ITD...", 1, 7);
            //timebatchdetail

            if (myRSFSP.Tables[0].Rows.Count == 0)
                sql = "no records";
            else
            {
                foreach (DataTable table in myRSFSP.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                         matsys = dr["mattersysnbr"].ToString();
                        period = dr["periodnumber"].ToString();
                        yr = dr["periodyear"].ToString();
                        hours = dr["Hours"].ToString();
                        h2b = dr["H2B"].ToString();
                        amount = dr["Amount"].ToString();
                        std = dr["StdAmt"].ToString();
                        tkpr = dr["timekeepersysnbr"].ToString();
                        taskcd = dr["TaskCd"].ToString();
                        actcd = dr["ActivityCd"].ToString();

                        sql = "update feesumbyprd set fspnonbilhrsentered=fspnonbilhrsentered + cast('" + hours + "' as decimal(12,2)),FSPBilHrsEntered=FSPBilHrsEntered - cast('" + hours + "' as decimal(12,2)),FSPFeeEnteredActualValue=FSPFeeEnteredActualValue - cast('" + amount + "' as decimal(12,2)) " +
         " where fspmatter=" + matsys + " and fspprdyear=" + yr + " and fspprdnbr=" + period + " and fsptkpr=" + tkpr + " and isnull(fsptaskcd,'')='" + taskcd + "' and isnull(fspactivitycd,'')='" + actcd + "' ";
                   _jurisUtility.ExecuteNonQueryCommand(0,sql);



                        sql = "update feesumitd  set fsinonbilhrsentered=fsinonbilhrsentered + cast('" + hours + "' as decimal(12,2)),FSiBilHrsEntered=FSiBilHrsEntered - cast('" + hours + "' as decimal(12,2)),FSiFeeEnteredActualValue=FSiFeeEnteredActualValue - cast('" + amount + "' as decimal(12,2)) " +
         " where fsimatter=" + matsys +  " and fsitkpr=" + tkpr ;
                        _jurisUtility.ExecuteNonQueryCommand(0, sql);
                    }
                }

            }





            UpdateStatus("Updating TimeBatch...", 3,7);
            //timebatchdetail
            sql = "update TimeBatchDetail set TBDBillableFlg = 'N' where tbdid in (select tbdid from timeentrylink inner join timeentry on timeentrylink.entryid=timeentry.entryid where entrysource='JurisGo' and entrystatus>=7 and billableflag='Y'  and MatterSysNbr in (select matsysnbr from matter where matbillagreecode in ('N','B')))";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating UnbilledTime...", 4,7);
            //unbilledtime
            sql = "update unbilledtime set UTBillableFlg = 'N'  where utid in (select tbdid from timeentrylink inner join timeentry on timeentrylink.entryid=timeentry.entryid where entrysource='JurisGo' and entrystatus>=7 and billableflag='Y'  and MatterSysNbr in (select matsysnbr from matter where matbillagreecode in ('N','B')))";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);

            UpdateStatus("Updating BilledTime...", 5,7);
            //billedtime
            sql = "update billedtime set BTBillableFlg = 'N'   where btid in (select tbdid from timeentrylink inner join timeentry on timeentrylink.entryid=timeentry.entryid where entrysource='JurisGo' and entrystatus>=7 and billableflag='Y'  and MatterSysNbr in (select matsysnbr from matter where matbillagreecode in ('N','B')))";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);


            UpdateStatus("Updating TimeEntry...", 6, 7);
            sql = "update TimeEntry set BillableFlag = 'N' where entrysource='JurisGo' and entrystatus>=7 and billableflag='Y'  and MatterSysNbr in (select matsysnbr from matter where matbillagreecode in ('N','B'))";
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
                double pctLong = Math.Round(((double)step/steps)*100.0);
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
            if (File.Exists(filePathName ))
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

            System.Environment.Exit(0);
          
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            //if matter and billing timekeeper
            if (true)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentBillingTimekeeper, 'DEF' as NewBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee on empsysnbr=billtobillingatty" +
                        " where empinitials<>'ABC'";


            //if matter and originating timekeeper
            else if (false)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentOriginatingTimekeeper, 'DEF' as NewOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join employee on empsysnbr=morigatty" +
                    " where empinitials<>'ABC'";


            return reportSQL;
        }


    }
}
