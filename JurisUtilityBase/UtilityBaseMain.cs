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
using System.Runtime.InteropServices;
using Microsoft.VisualBasic.FileIO;

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

        string pathToExcelFile = "";

        int codeLength = 0;

        bool codeIsNumeric = false;

        int glAcct = 0;

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
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

            if (!string.IsNullOrEmpty(pathToExcelFile))
            {
                toolStripStatusLabel.Text = "Running. Please Wait...";
                getNumberSettings();
                glAcct = testGLAcct();
                if (glAcct == 0)
                {
                    MessageBox.Show("No GL Account can be found with number 9313-000. The tool cannot continue", "GL Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    System.Environment.Exit(0);
                }
                else
                {

                    TextFieldParser parser = new TextFieldParser(pathToExcelFile);

                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.SetDelimiters(",");

                    string[] fields;
                    List<Vendor> vendList = new List<Vendor>();
                    Vendor vend = null;
                    int counter = 0;
                    while (!parser.EndOfData)
                    {

                        fields = parser.ReadFields();
                        if (counter != 0) //skip first line (headers)
                        {
                            vend = new Vendor();
                            vend.num = fields[0].Replace("\"", "").Replace("\"", "").Replace("'", "").Replace("\\", "/").Replace("*", "").Replace("%", "").Replace("`", "").Replace("!", "").Replace("@", "").Trim();
                            vend.name = fields[1].Replace("\"", "").Replace("\"", "").Replace("'", "").Replace("\\", "/").Replace("*", "").Replace("%", "").Replace("`", "").Replace("!", "").Replace("@", "").Trim();
                            vend.addy = fields[2].Replace("\"", "").Replace("\"", "").Replace("'", "").Replace("\\", "/").Replace("*", "").Replace("%", "").Replace("`", "").Replace("!", "").Replace("@", "").Trim();
                            vend.city = fields[3].Replace("\"", "").Replace("\"", "").Replace("'", "").Replace("\\", "/").Replace("*", "").Replace("%", "").Replace("`", "").Replace("!", "").Replace("@", "").Trim();
                            vend.state = fields[4].Replace("\"", "").Replace("\"", "").Replace("'", "").Replace("\\", "/").Replace("*", "").Replace("%", "").Replace("`", "").Replace("!", "").Replace("@", "").Trim();
                            vend.zip = fields[5].Replace("\"", "").Replace("\"", "").Replace("'", "").Replace("\\", "/").Replace("*", "").Replace("%", "").Replace("`", "").Replace("!", "").Replace("@", "").Trim();
                            vend.fein = fields[6].Replace("\"", "").Replace("\"", "").Replace("'", "").Replace("\\", "/").Replace("*", "").Replace("%", "").Replace("`", "").Replace("!", "").Replace("@", "").Trim();
                            if (!string.IsNullOrEmpty(vend.fein))
                                vend.hasFEIN = true;
                            vendList.Add(vend);
                        }

                        counter++;
                    }

                    parser.Close();

                    //ensure vendor doesnt exist and required data is there
                    List<Vendor> verifiedList = verifyVendor(vendList).ToList();
                    int errors = 0;
                    foreach (Vendor zz in verifiedList)
                    {
                        if (zz.hasError)
                            errors++;
                        else
                        {
                            string sql = "";
                            if (zz.hasFEIN)

                                sql = @"insert into vendor ([VenSysNbr],[VenCode],[VenIsTemp],[VenHoldPayments],[VenNoVouchers],[VenSeparateCheck],
                            [VenName],[VenAddress],[VenCity],[VenState],[VenZip],[VenCountry],[VenContact],[VenPhone],[VenFax],[VenRefNbr],
                            [VenGets1099],[VenFedIDType],[VenFedIDNbr],[VenPaymentGroup],[VenTermsDesc],[VenTermsType],[VenAlwaysTakeDisc],
                            [VenDueDays],[VenDiscDays],[VenDiscPcnt] ,[VenDiscCutOff] ,[VenLastPurchaseDate],[VenDefaultExpCode] ,
                            [VenDefaultAPAcct] ,[VenDefaultDistAcct] ,[VenDiscAcct] ,[Ven1099Box] ,[VenActive])
                            values((select max(vensysnbr) from vendor) + 1,'" + formatVendorCode(zz.num) + "', 'N','N','N','N', " +
                                "left('" + zz.name + "', 50), left('" + zz.addy + "', 250), left('" + zz.city + "',20), left('" + zz.state + "',2), left('" + zz.zip + "',9), '', '', '','', '', " +
                                "'Y', 'F', left('" + zz.fein + "',9), '','','D','N',  " +
                                "0,0,0.00,0,'01/01/1900', null, null, null,  " +
                                glAcct + ", 0, 'Y')";
                            else
                                sql = @"insert into vendor ([VenSysNbr],[VenCode],[VenIsTemp],[VenHoldPayments],[VenNoVouchers],[VenSeparateCheck],
                            [VenName],[VenAddress],[VenCity],[VenState],[VenZip],[VenCountry],[VenContact],[VenPhone],[VenFax],[VenRefNbr],
                            [VenGets1099],[VenFedIDType],[VenFedIDNbr],[VenPaymentGroup],[VenTermsDesc],[VenTermsType],[VenAlwaysTakeDisc],
                            [VenDueDays],[VenDiscDays],[VenDiscPcnt] ,[VenDiscCutOff] ,[VenLastPurchaseDate],[VenDefaultExpCode] ,
                            [VenDefaultAPAcct] ,[VenDefaultDistAcct] ,[VenDiscAcct] ,[Ven1099Box] ,[VenActive])
                            values((select max(vensysnbr) from vendor) + 1,'" + formatVendorCode(zz.num) + "', 'N','N','N','N', " +
                                "left('" + zz.name + "',50), left('" + zz.addy + "', 250), left('" + zz.city + "', 20), left('" + zz.state + "', 2), left('" + zz.zip + "', 9), '', '', '','', '', " +
                                "'N', 'F', '', '','','D','N',  " +
                                "0,0,0.00,0,'01/01/1900', null, null, null,  " +
                                glAcct + ", 0, 'Y')";
                            _jurisUtility.ExecuteSql(0, sql);

                            //(select chtsysnbr from chartofaccounts where dbo.jfn_FormatChartOfAccount(ChartOfAccounts.ChtSysNbr) = '9313-000')
                            sql = @"  insert into [DocumentTree] ([DTDocID] ,[DTSystemCreated] ,[DTDocClass] ,[DTDocType] ,[DTParentID] ,[DTTitle] ,[DTKeyL], [dtkeyt])
                          select (select max(dtdocid) from documenttree) + 1, 'Y', 7000,'R',66, left(venname, 30), vensysnbr, null
                          from vendor where vencode = '" + formatVendorCode(zz.num) + "'";
                            _jurisUtility.ExecuteSql(0, sql);
                        }

                    }

                    //update sysparam
                    string ss = @"  update [SysParam] set SpNbrValue = (select max(dtdocid) from documenttree)
                                where SpName = 'LastSysNbrVendor '";
                    _jurisUtility.ExecuteSql(0, ss);

                    UpdateStatus("All Vendors created.", 1, 1);
                    toolStripStatusLabel.Text = "Status: Ready to Execute";

                    if (errors == 0)
                        MessageBox.Show("The process is complete without error", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                    else
                    {
                        DialogResult rr = MessageBox.Show("The process is complete but there were" + "\r\n" + "errors. Would you like to see them?", "Errors", MessageBoxButtons.YesNo, MessageBoxIcon.None);
                        if (rr == DialogResult.Yes)
                        {
                            DataSet ds = displayErrors(verifiedList);
                            ReportDisplay rpds = new ReportDisplay(ds);
                            rpds.ShowDialog();
                        }
                    }

                    System.Environment.Exit(0);
                }
            }
            else
                MessageBox.Show("Please browse to your Excel file first", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            
        }

        private void getNumberSettings()
        {
            string sql = "  select SpTxtValue from sysparam where SpName = 'FldVendor'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            string cell = "";
            if (dds != null && dds.Tables.Count > 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                    cell = dr[0].ToString();
            }

            string[] test = cell.Split(',');

            if (test[1].Equals("C"))
                codeIsNumeric = false;
            else
                codeIsNumeric = true;
            codeLength = Convert.ToInt32(test[2].ToString());



        }

        private int testGLAcct()
        {
            int ret = 0;
            string sql = "select chtsysnbr from chartofaccounts where dbo.jfn_FormatChartOfAccount(ChartOfAccounts.ChtSysNbr) = '9313-000'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            string cell = "";
            if (dds != null && dds.Tables.Count > 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                    ret = Convert.ToInt32(dr[0].ToString());
            }

            return ret;
        }

        private string formatVendorCode(string code)
        {
            
            string formattedCode = "";
            if (codeIsNumeric)
            {
                formattedCode = "000000000000" + code;
                formattedCode = formattedCode.Substring(formattedCode.Length - 12, 12);
            }
            else
                formattedCode = code;
            return formattedCode;

        }

        //returns false if EID exists in timeentry, timebatchdetail and unbilledtime table as well as the taskcode existing in taskcode, otherwise returns true
        //which means at least one of these tests were failed and they need to be fixed
        private List<Vendor> verifyVendor(List<Vendor> vList)
        {

            foreach (Vendor vv in vList)
            {

                DataSet ds1;
                string SQL = "Select * from vendor where vencode = '" + formatVendorCode(vv.num) + "'";
                ds1 = _jurisUtility.ExecuteSqlCommand(0, SQL);
                if (ds1.Tables[0].Rows.Count > 0)
                {
                    vv.errorMessage = "Vendor number is already in Juris";
                    vv.hasError = true;
                }


                if (string.IsNullOrEmpty(vv.num))
                {
                    vv.errorMessage = "Vendor number cannot be blank";
                    vv.hasError = true;
                }

                if (string.IsNullOrEmpty(vv.name))
                {
                    vv.errorMessage = "Vendor name cannot be blank";
                    vv.hasError = true;
                }

                if (string.IsNullOrEmpty(vv.city))
                {
                    vv.errorMessage = "Vendor city cannot be blank";
                    vv.hasError = true;
                }

                if (string.IsNullOrEmpty(vv.state))
                {
                    vv.errorMessage = "Vendor state cannot be blank";
                    vv.hasError = true;
                }

                if (string.IsNullOrEmpty(vv.zip))
                {
                    vv.errorMessage = "Vendor zip cannot be blank";
                    vv.hasError = true;
                }

            }

            return vList;
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
            string AppDir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
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
            string AppDir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
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

        private DataSet displayErrors(List<Vendor> vlist)
        {
            DataSet ds = new DataSet();
            System.Data.DataTable errorTable = ds.Tables.Add("Errors");
            errorTable.Columns.Add("Code");
            errorTable.Columns.Add("Name");
            errorTable.Columns.Add("Address");
            errorTable.Columns.Add("City");
            errorTable.Columns.Add("State");
            errorTable.Columns.Add("Zip");
            errorTable.Columns.Add("FEIN");
            errorTable.Columns.Add("Error");


            foreach (Vendor ff in vlist)
            {
                if (ff.hasError)
                {
                    DataRow errorRow = ds.Tables["Errors"].NewRow();
                    errorRow["Code"] = ff.num;
                    errorRow["Name"] = ff.name;
                    errorRow["Address"] = ff.addy;
                    errorRow["City"] = ff.city;
                    errorRow["State"] = ff.state;
                    errorRow["Zip"] = ff.zip;
                    errorRow["FEIN"] = ff.fein;
                    errorRow["Error"] = ff.errorMessage;
                    ds.Tables["Errors"].Rows.Add(errorRow);
                }
            }

            return ds;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(openFileDialog1.FileName).ToLower().Trim() == ".csv")
                {
                    pathToExcelFile = openFileDialog1.FileName;
                    label2.Text = "File Chosen: " + Path.GetFileName(pathToExcelFile);

                }
                else
                    MessageBox.Show("Only valid csv files can be seleced (.csv)", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            } 
        }

        private void JurisLogoImageBox_Click(object sender, EventArgs e)
        {
           DialogResult dr =  MessageBox.Show("This will remove the problem vendors. Continue?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                getNumberSettings();
                List<string> vens = new List<string>();
                for (int i = 730001; i < 730106; i++)
                    vens.Add(i.ToString());
                string sql = "";
                foreach (string ven in vens)
                {
                    sql = @"delete from vendor where vencode = '" + formatVendorCode(ven) + "'";
                    _jurisUtility.ExecuteSql(0, sql);
                }
                MessageBox.Show("Done. Use the tool like normal now.");
             }
        }
    }
}
