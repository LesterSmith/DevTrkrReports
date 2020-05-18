#define DynSQL
/*
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO;
using System.Text;

namespace ClientMigrationVetting.BusinessObjects
{
    public class ExcelBuilderEP
    {
        #region public and private members
        const string yes = "Yes";
        const string no = "No";
        public ExcelWorksheet WorkSheet { get; set; }
        public ExcelPackage Excel { get; set; }
        private string FileName { get; set; }
        private FileInfo ExcelFile { get; set; }
        private int RowId { get; set; }
        private int ValidCell { get; set; }
        public List<Column> Columns = new List<Column>();
        private TextBox txtProcessing { get; set; }
        private TextBox txtCustid { get; set; }
        private TextBox txtField { get; set; }

       // private Forms.frmMain frmMain { get; set; }
        private List<ColHdr> ColHdrs { get; set; }
        //private string[] Headers =
        //    new string[]
        //               {"Client Friendly Name",
        //                "Acct Name",
        //                "CustID",
        //                "Account Grading - On or Off",
        //                "Adjudication (No Grading, System or Group Name)",
        //                "MVR Ordered (13 months)",
        //                "MVR PURPOSE",
        //                "MVR Scoring",
        //                "Credit Ordered  (13 months)",
        //                "Credit (No Grade, Generic, Generic 2 or Group Name)",
        //                "CREDIT TYPE (EP/OTHER)",
        //                "FACIS Ordered (13 months)",
        //                "FACIS LEVEL",
        //                "STATEWIDE Ordered (13 months)",
        //                "STATEWIDE - STD OR REPOSITORY",
        //                "Drug Testing Ordered (13 months)",
        //                "Panel By Package?",
        //                "PRE AA",
        //                "AA",
        //                "CA",
        //                "MN",
        //                "OK",
        //                "MA",
        //                "NY",
        //                "613",
        //                "Individualized Assessment",
        //                "CONSUMER NOTICE (NYC & LA)",
        //                "CA-1008",
        //                "Completion Email",
        //                "Completion Email Recipient (Recruiter or Specified Email Address)",
        //                "Pass Completion Email",
        //                "Pass Email Recipient (Recruiter or Specified Email Address)",
        //                "Review Completion Email",
        //                "Review Email Recipient (Recruiter or Specified Email Address)",
        //                "Fail Completion Email",
        //                "Fail Email Recipient (Recruiter or Specified Email Address)",
        //                "Change Grade Email",
        //                "Change Grade Recipient (Recruiter or Specified Email Address)",
        //                "Fail Change Grade Email",
        //                "Fail Change Grade Recipient (Recruiter or Specified Email Address)",
        //                "Pass Change Grade Email",
        //                "Pass Change Grade Recipient (Recruiter or Specified Email Address)",
        //                "Progression Yes/No",
        //                "Progression Email",
        //                "Progression Recipient",
        //                "Progression Tiers",
        //                "Elink (Yes/No)",
        //                "Elink Consent Days",
        //                "Elink Email",
        //                "Elink Reminder Days",
        //                "Billing Pass Thru Fees"};

        #endregion

        #region ..ctor
        public ExcelBuilderEP(string fileName, TextBox custid, TextBox processing, TextBox field, Forms.frmMain f)
        {
            FileName = fileName;
            txtCustid = custid;
            txtProcessing = processing;
            txtField = field;
            //frmMain = f;
            CreateColumnHeaders();
            CreateColumnObjects();
            Excel = new ExcelPackage();
            Excel.Workbook.Worksheets.Add("Vetting");
            Excel.Workbook.Worksheets.Add("Custid Errors");
            Excel.Workbook.Worksheets.Add("User Errors");
            ExcelFile = new FileInfo(fileName);
            WorkSheet = GetNewWorkBook();
        }
        #endregion

        #region public methods
        /// <summary>
        /// Populate the spreadsheet from database
        /// </summary>
        /// <param name="clients"></param>
        public void PopulateSheet(List<Client> clients)
        {
            var errCnt = 0;
        TryAgain:
            try
            {
                int ValidCustidRow = 2;
                int ValidUserRow = 2;
                var rowId = 2;
                var ws = Excel.Workbook.Worksheets[1];
                var ws2 = Excel.Workbook.Worksheets[2];
                var ws3 = Excel.Workbook.Worksheets[3];
                ws2.Cells[1, 1].Value = "Custid Validation Errors";
                ws2.Cells[1, 1].Style.Font.Bold = true;
                ws2.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                ws2.Cells[1,1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws2.Cells[1,1].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
                ws3.Cells[1, 1].Value = "Custid Validation Errors";
                ws3.Cells[1, 1].Style.Font.Bold = true;
                ws3.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                ws3.Cells[1,1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws3.Cells[1,1].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
                var hlpr = new DHMisc(ApplicationWrapper.ProdReportConnString);
                frmMain.ProgressBarInit(0, clients.Count, 0);
                Cursor.Current = Cursors.WaitCursor;
                var count = 1;
                foreach (Client client in clients)
                {
                    txtProcessing.Text = $"{count} of {clients.Count}";
                    txtCustid.Text = client.Custid;
                    var gisGrade = string.Empty;
                    txtField.Text = "SvcCd Order Counts";
                    frmMain.ProgressBarUpdate(count);
                    count++;
                    Application.DoEvents();
                    var ceSql = "select * from OldGIS.dbo.CIFExtension WITH (NOLOCK) where CustID = @Custid ";
                    var dtCifExt = hlpr.GetDatatableFromSQL(client.Custid, ceSql);
                    var orderSql = "Select eo.Svccd from OldGIS.dbo.personaldata as pd with(nolock) inner join EquestPlus.dbo.equestorders as eo with(nolock) on pd.ID = eo.CaseId where pd.Custid = @CustID and pd.DateRecvd > getdate() - 400 Group by Svccd";
                    var dtOrders = hlpr.GetDatatableFromSQL(client.Custid, orderSql);
                    List<string> svcCds = new List<string>();
                    foreach (DataRow dr in dtOrders.Rows)
                    {
                        svcCds.Add(dr["SvcCd"].ToString());
                    }

                    ws.Cells[rowId, 3].Value = client.Custid;

                    foreach (Column col in Columns)
                    {
                        var errCnt2 = 0;
                    TryAgain2:
                        try
                        {
                            txtField.Text = col.Description;
                            Application.DoEvents();

                            DataTable dt = null;
                            if (!string.IsNullOrWhiteSpace(col.SqlText))
                            {
                                dt = hlpr.GetDatatableFromSQL(client.Custid, col.SqlText);
                            }

                            switch (col.ColNbr)
                            {
                                case 1: // Cust Name
                                    ws.Cells[rowId, 2].Value = dt.Rows.Count > 0 ? dt.Rows[0]["Name"].GetNotDBNull() : "Missing Customer Name";
                                    break;
                                case 2: // Custid
                                    ws.Cells[rowId, 3].Value = client.Custid;
                                    break;
                                case 3:
                                    if (dtCifExt.Rows.Count < 1)
                                        gisGrade = na;
                                    else
                                        gisGrade = dtCifExt.Rows[0]["GisGrade"].IsDBNull() || dtCifExt.Rows[0]["GisGrade"].GetNotDBNullInt() == 0 ? "No" : yes;
                                    ws.Cells[rowId, 4].Value = gisGrade;
                                    break;
                                case 4:
                                    if (dtCifExt.Rows.Count < 1)
                                        ws.Cells[rowId, 5].Value = na;
                                    else
                                        ws.Cells[rowId, 5].Value = GetGradingType(gisGrade, dtCifExt, client.Custid);
                                    break;
                                case 5: // mvr ordering
                                    ws.Cells[rowId, 6].Value = IsOrdered(svcCds, "MVR");
                                    break;
                                case 6:
                                    if (dtCifExt.Rows.Count < 1)
                                        ws.Cells[rowId, 7].Value = na;
                                    else
                                        ws.Cells[rowId, 7].Value = dtCifExt.Rows[0]["MVRPurpose"].DBNullStringTest() == "E" ? "Employment" : !string.IsNullOrWhiteSpace(dtCifExt.Rows[0]["MVRPurpose"].GetNotDBNull()) ? dtCifExt.Rows[0]["MVRPurpose"].DBNullStringTest() : na;
                                    break;
                                case 7:
                                    ws.Cells[rowId, 8].Value = dtCifExt.Rows.Count > 0 && dtCifExt.Rows[0]["IIXCustAcct"].GetNotDBNull().ToUpper() == "SOFTTECH" && dt.Rows.Count > 0 ? dt.Rows[0]["MVRScoring"].GetNotDBNull() : na;
                                    break;
                                case 8:
                                    ws.Cells[rowId, 9].Value = IsOrdered(svcCds, "CH");
                                    break;
                                case 9:
                                    ws.Cells[rowId, 10].Value = dt.Rows.Count > 0 ? dt.Rows[0]["CreditGroup"].GetNotDBNull() : na;
                                    break;
                                case 10:
                                    ws.Cells[rowId, 11].Value = dt.Rows.Count > 0 ? dt.Rows[0]["CreditType"].GetNotDBNull() : na;
                                    break;
                                case 11:
                                    ws.Cells[rowId, 12].Value = IsOrdered(svcCds, "FACISI");
                                    break;
                                case 12:
                                    ws.Cells[rowId, 13].Value = dt.Rows.Count > 0 ? dt.Rows[0]["FacisLevel"].GetNotDBNull() : na;
                                    break;
                                case 13:
                                    ws.Cells[rowId, 14].Value = IsOrdered(svcCds, "CRCSTA");
                                    break;
                                case 14:
                                    ws.Cells[rowId, 15].Value = dtCifExt.Rows.Count > 0 && dtCifExt.Rows[0]["eQuestFieldExt"].GetNotDBNull().ToLower().Contains("-STRESP") ? "Repository" : "Standard";
                                    break;
                                case 15:
                                    ws.Cells[rowId, 16].Value = IsOrdered(svcCds, "DRGTST");
                                    break;
                                case 17: // letter cols R-Y(17-24)
                                    List<Letter> ltrs = GetLetterData(dt);
                                    var preAA = no;
                                    var AA = no;
                                    var CA = no;
                                    var MN = no;
                                    var OK = no;
                                    var MA = no;
                                    var NY = no;
                                    var six13 = no;
                                    foreach (var ltr in ltrs)
                                    {
                                        if (ltr.LetterType.Equals("3") || ltr.LetterType.Equals("5")) preAA = yes;
                                        if (ltr.LetterType.Equals("4") || ltr.LetterType.Equals("6")) AA = yes;
                                        if (ltr.LetterType.Equals("D")) NY = yes;
                                        if (ltr.LetterType.Equals("S")) MA = yes;
                                        if (ltr.LetterType.Equals("B"))
                                        {
                                            if (ltr.StateCode.Equals("CA")) CA = yes;
                                            if (ltr.StateCode.Equals("MN")) MN = yes;
                                            if (ltr.StateCode.Equals("OK")) OK = yes;
                                        }
                                        if (ltr.LetterType.Equals("613")) six13 = yes;
                                    }
                                    ws.Cells[rowId, 18].Value = preAA;
                                    ws.Cells[rowId, 19].Value = AA;
                                    ws.Cells[rowId, 20].Value = CA;
                                    ws.Cells[rowId, 21].Value = MN;
                                    ws.Cells[rowId, 22].Value = OK;
                                    ws.Cells[rowId, 23].Value = MA;
                                    ws.Cells[rowId, 24].Value = NY;
                                    ws.Cells[rowId, 25].Value = six13;
                                    break;
                                case 26:
                                    ws.Cells[rowId, 27].Value = dt.Rows.Count > 0 ? yes : no;
                                    break;
                                case 27:
                                    ws.Cells[rowId, 28].Value = dt.Rows.Count > 0 && dt.Rows[0]["CustOptionValue"].GetNotDBNull().ToUpper().StartsWith("Y") ? yes : no;
                                    break;
                                case 28:
                                    PutEmailDataInSheet(client.Custid, ws, hlpr, rowId);
                                    break;
                                case 42:
                                    var prgOnOff = string.Empty;
                                    var tiers = string.Empty;
                                    GetProgression(client.Custid, hlpr, out prgOnOff, out tiers);
                                    ws.Cells[rowId, 43].Value = prgOnOff;
                                    ws.Cells[rowId, 46].Value = tiers;
                                    break;
                                case 46:
                                    ws.Cells[rowId, 47].Value = dt.Rows.Count > 0 ? yes : no;
                                    break;
                                case 47:
                                    ws.Cells[rowId, 48].Value = dt.Rows.Count > 0 ? dt.Rows[0]["CustOptionValue"] : na;
                                    break;
                                case 48:
                                    ws.Cells[rowId, 49].Value = GetElinkEmail(client.Custid, hlpr);
                                    break;
                                case 49:
                                    ws.Cells[rowId, 50].Value = dt.Rows.Count > 0 ? dt.Rows[0]["CustOptionValue"].GetNotDBNull() : na;
                                    break;
                                case 50:
                                    ws.Cells[rowId, 51].Value = dtCifExt.Rows.Count < 1 ? na : dtCifExt.Rows[0]["BillCourtFees"].GetNotDBNull()  == "-County" ? "Long" : "Short";
                                    break;
                                case 51:
                                    var errCustid = CheckForWarnings(client.Custid);

                                    // get the errors if any, put them in the Validation Errors sheet
                                    // and put a link in cell in main sheet to go to the errors

                                    if (!string.IsNullOrWhiteSpace(errCustid))
                                    {
                                        ws.Cells[rowId, 52].Value = "See Errors";
                                        //Uri url = new Uri($"#'Custid Errors'A{ValidCustidRow}", UriKind.Relative);
                                        ws.Cells[rowId, 52].Hyperlink = new ExcelHyperLink((char)39 + "Custid Errors" + (char)39 + "!A" + (ValidCustidRow+1).ToString(), "See Errors");
                                        ws.Cells[rowId, 52].Style.Font.UnderLine = true;
                                        ws.Cells[rowId, 52].Style.Font.Bold = true;
                                        ws.Cells[rowId, 52].Style.Font.Color.SetColor(Color.Red);
                                        ValidCustidRow++; // insert blank line before new custid
                                        ws2.Cells[ValidCustidRow, 1].Style.Font.Bold = true;
                                        ws2.Cells[ValidCustidRow++, 1].Value = client.Custid;
                                        using (var sr = new StringReader(errCustid))
                                        {
                                            var line = string.Empty;
                                            while (line != null)
                                            {
                                                line = sr.ReadLine();
                                                if (line == null) break;
                                                if (string.IsNullOrWhiteSpace(line)) continue;
                                                ws2.Cells[ValidCustidRow++, 1].Value = line;
                                            }
                                        }
                                    }
                                    break;
                                case 52:
                                    var errUser = GetUserErrors(client.Custid);
                                    if (!string.IsNullOrWhiteSpace(errUser))
                                    {
                                        ws.Cells[rowId, 53].Value = "See Errors";
                                        //Uri url = new Uri($"#'User Errors'A{ValidUserRow}", UriKind.Relative);
                                        //ws.Cells[rowId, 53].Hyperlink = url;
                                        ws.Cells[rowId, 53].Hyperlink = new ExcelHyperLink((char)39 + "User Errors" + (char)39 + "!A" + (ValidUserRow+1).ToString(), "See Errors");
                                        ws.Cells[rowId, 53].Style.Font.UnderLine = true;
                                        ws.Cells[rowId, 53].Style.Font.Bold = true;
                                        ws.Cells[rowId, 53].Style.Font.Color.SetColor(Color.Red);
                                        ValidUserRow++; // insert blank line before new Custid
                                        ws3.Cells[ValidUserRow, 1].Style.Font.Bold = true;
                                        ws3.Cells[ValidUserRow++, 1].Value = client.Custid;
                                        using (var sr = new StringReader(errUser))
                                        {
                                            var line = string.Empty;
                                            while (line != null)
                                            {
                                                line = sr.ReadLine();
                                                if (line == null) break;
                                                if (string.IsNullOrWhiteSpace(line)) continue;
                                                ws3.Cells[ValidUserRow++, 1].Value = line;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            if ((ex.Message.ToLower().IndexOf("transport") > -1 || ex.Message.ToLower().IndexOf("connection") > -1 || ex.Message.ToLower().Contains("timeout") || ex.Message.ToLower().Contains("network")) && ++errCnt2 < 3)
                                goto TryAgain2;
                            MessageBox.Show($"{client.Custid} had a System Error, {ex.Message}", "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    rowId++;
                    //Excel.SaveAs(ExcelFile);
                }
                Cursor.Current = Cursors.Default;
                frmMain.ProgressBarClear();
                Excel.SaveAs(ExcelFile);
            }
            catch (Exception ex)
            {
                if ((ex.Message.ToLower().IndexOf("transport") > -1 || ex.Message.ToLower().IndexOf("connection") > -1 || ex.Message.ToLower().Contains("timeout") || ex.Message.ToLower().Contains("network")) && ++errCnt < 3)
                    goto TryAgain;
                throw;
            }
        }
        #endregion

        #region private methods
        private string GetElinkEmail(string custId, DHMisc hlpr)
        {
            var retValue = no;
            var sql = "Select GroupIDDef from OLDGIS.dbo.EmailGroup with(nolock) where GroupIDDef Like '%" + custId + "%' ";
            var dt = hlpr.GetDatatableFromSQL(custId, sql);
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["GroupIdDef"].GetNotDBNull().ToUpper().Contains("NOTICE"))
                {
                    retValue = yes;
                    break;
                }
            }
            return retValue;
        }
        private void GetProgression(string custId, DHMisc hlpr, out string prgOnOff, out string tiers)
        {
            var dt = hlpr.GetProgression(custId);
            if (dt.Rows.Count < 1)
            {
                prgOnOff = no;
                tiers = "None";
                return;
            }
            prgOnOff = yes;
            tiers = dt.Rows[dt.Rows.Count - 1]["Tier"].GetNotDBNullInt().ToString();
            return;
        }
        private List<Letter> GetLetterData(DataTable dt)
        {
            List<Letter> list = (from DataRow dr in dt.Rows
                                 select new Letter
                                 {
                                     LetterType = dr["LetterType"].GetNotDBNull(),
                                     LetterName = dr["LetterName"].GetNotDBNull(),
                                     StateCode = dr["StateCd"].GetNotDBNull(),
                                     SumOfRightsFile = dr["SumOfRightsFile"].GetNotDBNull(),
                                     DisputeRequestFile = dr["DisputeRequestFile"].GetNotDBNull()
                                 }).ToList();
            return list;
        }

        const string comp = "_comp";
        const string pcomp = "_pcomp";
        const string rcomp = "_rcomp";
        const string fcomp = "_fcomp";
        const string fcg = "_fcg";
        const string pcg = "_pcg";
        const string prog = "_prog";
        const string qaComplete = "QACOMPLETE";
        const string custGrade = "CUSTGRADE";
        const string tierComplete = "TIERCOMPLETE";
        const string na = "Not Available";

        /// <summary>
        /// Column AC is col 28
        /// </summary>
        /// <param name="custId"></param>
        /// <param name="ws"></param>
        /// <param name="hlpr"></param>
        private void PutEmailDataInSheet(string custId, ExcelWorksheet ws, DHMisc hlpr, int rowId)
        {
            var sql = "Select Custid, EventId, SvcCd, SvcParameterList from EquestPlus.dbo.CustomerEvents with (nolock) where CustID = @Custid and SvcCd = 'GTETMP' and EventID in ('CUSTGRADE','QACOMPLETE', 'TIERCOMPLETE') Order by Custid ";
            var dt = hlpr.GetDatatableFromSQL(custId, sql);
            //if (dt.Rows.Count > 0)
            //    return; // nothing else to do

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    var eventId = dr["EventID"].GetNotDBNull().ToUpper();
                    var svcParList = dr["SvcParameterList"].GetNotDBNull().ToLower();

                    ws.Cells[rowId, 29].Value = eventId.Equals("QACOMPLETE") && svcParList.Contains(comp) ? yes : no;
                    ws.Cells[rowId, 31].Value = eventId.Equals("QACOMPLETE") && svcParList.Contains(pcomp) ? yes : no;
                    ws.Cells[rowId, 33].Value = eventId.Equals("QACOMPLETE") && svcParList.Contains(rcomp) ? yes : "No";
                    ws.Cells[rowId, 35].Value = eventId.Equals("QACOMPLETE") && svcParList.Contains(fcomp) ? yes : "No";
                    ws.Cells[rowId, 37].Value = eventId.Equals("CUSTGRADE") && svcParList.Contains(pcomp) ? yes : "No";
                    ws.Cells[rowId, 39].Value = eventId.Equals("CUSTGRADE") && svcParList.Contains(fcg) ? yes : "No";
                    ws.Cells[rowId, 41].Value = eventId.Equals("CUSTGRADE") && svcParList.Contains(pcg) ? yes : "No";
                    ws.Cells[rowId, 44].Value = eventId.Equals(tierComplete) ? yes : "No";
                }
            }
            else
            {
                ws.Cells[rowId, 29].Value = na;
                ws.Cells[rowId, 31].Value = na;
                ws.Cells[rowId, 33].Value = na;
                ws.Cells[rowId, 35/].Value = na;
                ws.Cells[rowId, 37].Value = na;
                ws.Cells[rowId, 39].Value = na;
                ws.Cells[rowId, 41].Value = na;
                ws.Cells[rowId, 44].Value = na;
            }
            sql = "Select EC.CustID, CE.EventID, EG.GroupIDDef, EG.EmailRequestor, EG.DistributionList from OLDGIS.dbo.EmailGroup EG with(nolock) Join OLDGIS.dbo.EmailCustomers EC with(nolock) on EC.GroupIDDef = EG.GroupIDDef Join EquestPlus.dbo.CustomerEvents CE with(nolock)  on CE.CustID = EC.CustID where EC.CustID = @custId Order by EC.CustID";
            dt = hlpr.GetDatatableFromSQL(custId, sql);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    var eventId = dr["EventID"].GetNotDBNull().ToUpper();
                    var groupDefId = dr["GroupIDDef"].GetNotDBNull().ToLower();
                    var emailRequestor = dr["EmailRequestor"].GetNotDBNullInt();
                    var distList = dr["DistributionList"].GetNotDBNull();

                    if (eventId.Equals(qaComplete) && groupDefId.Contains(comp) && emailRequestor.Equals(1) && string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 30].Value = emailRequestor;
                    else if (eventId.Equals(qaComplete) && groupDefId.Contains(comp) && (emailRequestor.Equals(0) || emailRequestor.Equals("255")))
                        ws.Cells[rowId, 30].Value = distList;
                    else if (eventId.Equals(qaComplete) && groupDefId.Contains(comp) && emailRequestor.Equals(1) && !string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 30].Value = $"({emailRequestor} {distList})";
                    else
                        ws.Cells[rowId, 30 ].Value = na;

                    if (eventId.Equals(qaComplete) && groupDefId.Contains(pcomp) && emailRequestor.Equals(1) && string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 32].Value = emailRequestor;
                    else if (eventId.Equals(qaComplete) && groupDefId.Contains(pcomp) && (emailRequestor.Equals(0) || emailRequestor.Equals(255)))
                        ws.Cells[rowId, 32].Value = distList;
                    else if (eventId.Equals(qaComplete) && groupDefId.Contains(pcomp) && emailRequestor.Equals(1) && !string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 32].Value = $"({emailRequestor} {distList})"; 
                    else
                        ws.Cells[rowId, 32 ].Value = na; 

                    if (eventId.Equals(qaComplete) && groupDefId.Contains(rcomp) && emailRequestor.Equals(1) && string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 34].Value = emailRequestor; 
                    else if (eventId.Equals(qaComplete) && groupDefId.Contains(rcomp) && (emailRequestor.Equals(0) || emailRequestor.Equals(255)))
                        ws.Cells[rowId, 34].Value = distList; 
                    else if (eventId.Equals(qaComplete) && groupDefId.Contains(rcomp) && emailRequestor.Equals(1) && !string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 34].Value = $"({emailRequestor} {distList})"; 
                    else
                        ws.Cells[rowId, 34 ].Value = na; 

                    if (eventId.Equals(qaComplete) && groupDefId.Contains(fcomp) && emailRequestor.Equals(1) && string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 36].Value = emailRequestor; 
                    else if (eventId.Equals(qaComplete) && groupDefId.Contains(fcomp) && (emailRequestor.Equals(0) || emailRequestor.Equals(255)))
                        ws.Cells[rowId, 36].Value = distList; 
                    else if (eventId.Equals(qaComplete) && groupDefId.Contains(fcomp) && emailRequestor.Equals(1) && !string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 36].Value = $"({emailRequestor} {distList})"; 
                    else
                        ws.Cells[rowId, 36 ].Value = na; 

                    if (eventId.Equals(custGrade) && groupDefId.Contains(comp) && emailRequestor.Equals(1) && string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 38].Value = emailRequestor; 
                    else if (eventId.Equals(custGrade) && groupDefId.Contains(comp) && (emailRequestor.Equals(0) || emailRequestor.Equals(255)))
                        ws.Cells[rowId, 38].Value = distList; 
                    else if (eventId.Equals(custGrade) && groupDefId.Contains(comp) && emailRequestor.Equals(1) && !string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 38].Value = $"({emailRequestor} {distList})"; 
                    else
                        ws.Cells[rowId, 38 ].Value = na; 

                    if (eventId.Equals(custGrade) && groupDefId.Contains(fcg) && emailRequestor.Equals(1) && string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 40].Value = emailRequestor; 
                    else if (eventId.Equals(custGrade) && groupDefId.Contains(fcg) && (emailRequestor.Equals(0) || emailRequestor.Equals(255)))
                        ws.Cells[rowId, 40].Value = distList; 
                    else if (eventId.Equals(custGrade) && groupDefId.Contains(fcg) && emailRequestor.Equals(1) && !string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 49].Value = $"({emailRequestor} {distList})"; 
                    else
                        ws.Cells[rowId, 40].Value = na; 

                    if (eventId.Equals(custGrade) && groupDefId.Contains(pcg) && emailRequestor.Equals(1) && string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 42].Value = emailRequestor; 
                    else if (eventId.Equals(custGrade) && groupDefId.Contains(pcg) && (emailRequestor.Equals(0) || emailRequestor.Equals(255)))
                        ws.Cells[rowId, 42].Value = distList; 
                    else if (eventId.Equals(custGrade) && groupDefId.Contains(pcg) && emailRequestor.Equals(1) && !string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 42].Value = $"({emailRequestor} {distList})"; 
                    else
                        ws.Cells[rowId, 42 ].Value = na; 

                    if (eventId.Equals(custGrade) && groupDefId.Contains(prog) && emailRequestor.Equals(1) && string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 45].Value = emailRequestor; 
                    else if (eventId.Equals(custGrade) && groupDefId.Contains(prog) && (emailRequestor.Equals(0) || emailRequestor.Equals(255)))
                        ws.Cells[rowId, 45].Value = distList; 
                    else if (eventId.Equals(custGrade) && groupDefId.Contains(prog) && emailRequestor.Equals(1) && !string.IsNullOrWhiteSpace(distList))
                        ws.Cells[rowId, 45].Value = $"({emailRequestor} {distList})"; 
                    else
                        ws.Cells[rowId, 45 ].Value = na; 
                }
            }
            else
            {
                ws.Cells[rowId, 30 ].Value = na; 
                ws.Cells[rowId, 32 ].Value = na; 
                ws.Cells[rowId, 34 ].Value = na; 
                ws.Cells[rowId, 36 ].Value = na; 
                ws.Cells[rowId, 38 ].Value = na; 
                ws.Cells[rowId, 30 ].Value = na; 
                ws.Cells[rowId, 42 ].Value = na; 
                ws.Cells[rowId, 45 ].Value = na; 
            }
        }


        private string GetGradingType(string gisGrade, DataTable cifExt, string custid)
        {
            if (cifExt.Rows[0]["GisGrade"].GetNotDBNullInt() != 255)
                return "No Grading";

            var hlpr = new DHMisc(Classes.ApplicationWrapper.ProdReportConnString);
            var sqlGT = "SELECT GroupId FROM OLDGIS.dbo.ClientProcessGroup WITH(nolock) WHERE CustId = @CustId AND ProcessId = 'ADJUD_RULES'";
            var dtGrdGrp = hlpr.GetDatatableFromSQL(custid, sqlGT);
            var grdGrp = dtGrdGrp.Rows.Count > 0 ? dtGrdGrp.Rows[0]["GroupId"].GetNotDBNull() : "Not Available";
            var retval = string.Empty;

            var dt = hlpr.GetGradingData(custid);
            if (dt.Rows.Count == 0)
                return "Not Available";
            var custTrueGrd = dt.Rows.Count > 0 ? dt.Rows[0]["CustomGrading"].GetNotDBNullBool() : false;

            //var eLink = dt.Rows.Count > 0 ? dt.Rows[0]["Elink"].GetNotDBNullBool() : false;
            var eLink = hlpr.GetElinkSettings(custid, DateTime.Parse("12/31/9999"));
            // calc the grading value CustomGrading, ELink
            if (gisGrade == "No")
                return "NoGrade";
            else if (eLink && custTrueGrd)
                return $"TrueCustomAndElink/{grdGrp}";
            else if (custTrueGrd)
                return $"TrueCustom/{grdGrp}";
            else if (eLink && !custTrueGrd)
                return $"ElinkCustom/{grdGrp}";
            else
                return "System Grading";

        }

        private string IsOrdered(List<string> svcCds, string svcCd)
        {
            var sc = svcCds.Find(s => s == svcCd);
            return !string.IsNullOrWhiteSpace(sc) ? yes : "No";
        }
        private ExcelWorksheet GetNewWorkBook()
        {
            try
            {
                var ws = Excel.Workbook.Worksheets[1];
                CreateHdrBoldStyle(ws, "A1:BA1");
                for (var i = 0; i < ColHdrs.Count; i++)
                {
                    var hdr = ColHdrs[i];
                    ws.Cells[1, i + 1].Value = hdr.Header;
                    ws.Column(i + 1).Width = hdr.Width;   
                }
                Excel.SaveAs(ExcelFile);
                return ws;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void CreateHdrBoldStyle(ExcelWorksheet ws, string hdrRange, string style = "@")
        {
            try
            {
                ws.Cells[hdrRange].Style.Numberformat.Format = style;
                ws.Cells[hdrRange].Style.Font.Bold = true;
                ws.Row(1).Height = 55;
                ws.Cells[hdrRange].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[hdrRange].Style.WrapText = true;
                ws.Cells[hdrRange].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[hdrRange].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
                ws.Cells[hdrRange].Style.Font.Color.SetColor(Color.White);
                ws.View.FreezePanes(2, 1);
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        private void CreateColumnHeaders()
        {
            ColHdrs = new List<ColHdr>();
            ColHdrs.Add(new ColHdr { Header = "Client Friendly Name", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "Acct Name", Width= 40 });
            ColHdrs.Add(new ColHdr { Header = "CustID", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Account Grading - On or Off", Width = 10 });
            ColHdrs.Add(new ColHdr { Header = "Adjudication (No Grading, System or Group Name)", Width = 35 });
            ColHdrs.Add(new ColHdr { Header = "MVR Ordered (13 months)", Width = 10 });
            ColHdrs.Add(new ColHdr { Header = "MVR PURPOSE", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "MVR Scoring", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "Credit Ordered (13 months)", Width = 10 });
            ColHdrs.Add(new ColHdr { Header = "Credit (No Grade, Generic, Generic 2 or Group Name)", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "CREDIT TYPE (EP/OTHER)", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "FACIS Ordered (13 months)", Width = 10 });
            ColHdrs.Add(new ColHdr { Header = "FACIS LEVEL", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "STATEWIDE Ordered (13 months)", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "STATEWIDE - STD OR REPOSITORY", Width = 14 });
            ColHdrs.Add(new ColHdr { Header = "Drug Testing Ordered (13 months)", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Panel By Package?", Width = 10 });
            ColHdrs.Add(new ColHdr { Header = "PRE AA", Width = 8 });
            ColHdrs.Add(new ColHdr { Header = "AA", Width = 5 });
            ColHdrs.Add(new ColHdr { Header = "CA", Width = 5 });
            ColHdrs.Add(new ColHdr { Header = "MN", Width = 5 });
            ColHdrs.Add(new ColHdr { Header = "OK", Width = 5 });
            ColHdrs.Add(new ColHdr { Header = "MA", Width = 5 });
            ColHdrs.Add(new ColHdr { Header = "NY", Width = 5 });
            ColHdrs.Add(new ColHdr { Header = "613", Width = 5 });
            ColHdrs.Add(new ColHdr { Header = "Individualized Assessment", Width = 14 });
            ColHdrs.Add(new ColHdr { Header = "CONSUMER NOTICE (NYC & LA)", Width = 14 });
            ColHdrs.Add(new ColHdr { Header = "CA-1008", Width = 5 });
            ColHdrs.Add(new ColHdr { Header = "Completion Email", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Completion Email Recipient (Recruiter or Specified Email Address)", Width = 30 });
            ColHdrs.Add(new ColHdr { Header = "Pass Completion Email", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Pass Email Recipient (Recruiter or Specified Email Address)", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "Review Completion Email", Width = 14 });
            ColHdrs.Add(new ColHdr { Header = "Review Email Recipient (Recruiter or Specified Email Address)", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "Fail Completion Email", Width = 14 });
            ColHdrs.Add(new ColHdr { Header = "Fail Email Recipient (Recruiter or Specified Email Address)", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "Change Grade Email", Width = 14 });
            ColHdrs.Add(new ColHdr { Header = "Change Grade Recipient (Recruiter or Specified Email Address)", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "Fail Change Grade Email", Width = 14 });
            ColHdrs.Add(new ColHdr { Header = "Fail Change Grade Recipient (Recruiter or Specified Email Address)", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "Pass Change Grade Email", Width = 14 });
            ColHdrs.Add(new ColHdr { Header = "Pass Change Grade Recipient (Recruiter or Specified Email Address)", Width = 20 });
            ColHdrs.Add(new ColHdr { Header = "Progression Yes/No", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Progression Email", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Progression Recipient", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Progression Tiers", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Elink (Yes/No)", Width = 10 });
            ColHdrs.Add(new ColHdr { Header = "Elink Consent Days", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Elink Email", Width = 10 });
            ColHdrs.Add(new ColHdr { Header = "Elink Reminder Days", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Billing Pass Thru Fees", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "Custid Errors", Width = 12 });
            ColHdrs.Add(new ColHdr { Header = "User Errors", Width = 12 });
        }

        private void CreateColumnObjects()
        {
            Columns.Add(new Column
            {
                ColNbr = 0
            });

            Columns.Add(new Column
            {
                SqlText = "Select Name From Oldgis..Cif_Base with (nolock) Where Custid = @Custid ",
                ColNbr = 1
            });

            Columns.Add(new Column
            {
                Description = "Grading On/Off",
                SqlText = string.Empty,
                ColNbr = 3
            });

            Columns.Add(new Column
            {
                Description = "Grading Type",
                SqlText = "",
                ColNbr = 4
            });

            Columns.Add(new Column
            {
                SqlText = "", //"Select MVROrdering = Case When mo.MvrCount > 0 Then 'Yes' Else 'No' End From (select count(*) as MvrCount from equestplus.dbo.eQuestOrders EO with(nolock) join OLDGIS.dbo.PersonalData PD with(nolock) on PD.ID = EO.CaseID Where PD.CustID = @Custid and EO.SvcCd = 'MVR' and EO.RequestedDtm between dateadd(month, -13, getdate()) and getdate()) mo ",
                ColNbr = 5
            });

            Columns.Add(new Column
            {
                Description = "MVR Purpose",
                SqlText = "select MVROrderMode, MVRGradingInd, MVRPriority, Case When MVRPurpose = 'E' Then 'Employment' Else MvrPurpose End as MVRPurpose, IIXCustAcct from OldGIS.dbo.CIFExtension WITH (NOLOCK) where CustID = @Custid ",
                ColNbr = 6
            });

            Columns.Add(new Column
            {
                Description = "MVR Scoring",
                SqlText = "Select RuleGroupId as MVRScoring from GISMVR.dbo.MVRGradingGroupCustomerList with (nolock) where CustID = @Custid ",
                ColNbr = 7
            });

            Columns.Add(new Column
            {
                Description = "Credit Ordered",
                SqlText = "", //"Select CreditOrdered = Case When sq.CreditCount > 0 Then 'Yes' Else 'No' End From (select count(*) as CreditCount from equestplus.dbo.eQuestOrders EO with(nolock) join OLDGIS.dbo.PersonalData PD with(nolock) on PD.ID = EO.CaseID Where PD.CustID = @Custid and EO.SvcCd = 'CH' and EO.RequestedDtm between dateadd(month, -13, getdate()) and getdate())  sq ",
                ColNbr = 8
            });

            Columns.Add(new Column
            {
                Description = "Credit Group",
                SqlText = "select CreditGroup = Case When RptGrdDefCd is not null Then RptGrdDefCd Else 'Not Ordering Credit' End from OLDGIS..DBCLUSTER2_GiS_dbo_ReportGradingDefGroup WITH (NOLOCK) where CustAcctCd = @Custid ",
                ColNbr = 9
            });

            Columns.Add(new Column
            {
                Description = "Credit Type",
                SqlText = "Declare @GroupId varchar (50) Set @GroupId = (Select GroupId from oldgis.dbo.providergroups with (nolock) where CustID = @Custid) Select ReasonCode as CreditType from oldgis.dbo.ProviderPasswords with(Nolock) where GroupID = @GroupId ",
                ColNbr = 10
            });

            Columns.Add(new Column
            {
                Description = "Facis Ordered",
                SqlText = "", //"Select FacisOrdered = Case When fo.FacisCnt > 0 Then 'Yes' Else 'No' End From (select count(*) as FacisCnt from equestplus.dbo.eQuestOrders EO with(nolock) join OLDGIS.dbo.PersonalData PD with(nolock) on PD.ID = EO.CaseID Where PD.CustID = @Custid and EO.SvcCd = 'FACISI' and EO.RequestedDtm between dateadd(month, -13, getdate()) and getdate()) fo ",
                ColNbr = 11
            });

            Columns.Add(new Column
            {
                Description = "Facis Level",
                SqlText = "select case when CustOptionValue is not null then 'FacisLevel ' + CustOptionValue Else 'No' End FacisLevel from EquestPlus.dbo.customeroptions with (nolock) where CustID = @Custid and CustOptionName = 'FacisILevel' ",
                ColNbr = 12
            });

            Columns.Add(new Column
            {
                Description = "Statewide Ordered",
                SqlText = "", //"Select StateWide = Case When sw.StateWideCnt > 0 Then 'Yes' Else 'No' End From (select count(*) as StateWideCnt from equestplus.dbo.eQuestOrders EO with(nolock) join OLDGIS.dbo.PersonalData PD with(nolock) on PD.ID = EO.CaseID Where PD.CustID = @Custid and EO.SvcCd = 'CRCSTA' and EO.RequestedDtm between dateadd(month, -13, getdate()) and getdate()) sw ",
                ColNbr = 13
            });

            Columns.Add(new Column
            {
                Description = "Statewide Type",
                SqlText = string.Empty, // "select StateWideType = case when equestfieldext like '%-STRESP%' Then 'Repository' Else 'Standard' end from oldgis.dbo.CIFExtension with (nolock) where CustID = @Custid ",
                ColNbr = 14
            });

            Columns.Add(new Column
            {
                Description = "DrgTst Ordered",
                SqlText = "", //"Select DrgTstOrdered = Case When dt.DrgTstCnt > 0 Then 'Yes' Else 'No' End From (select count(*) as DrgTstCnt from equestplus.dbo.eQuestOrders EO with(nolock) join OLDGIS.dbo.PersonalData PD with(nolock) on PD.ID = EO.CaseID Where PD.CustID = @Custid and EO.SvcCd = 'DRGTST' and EO.RequestedDtm between dateadd(month, -13, getdate()) and getdate()) dt ",
                ColNbr = 15
            });

            Columns.Add(new Column
            {
                Description = "Letter Data",
                SqlText = "Select CustID, LetterType, StateCD, SumOfRightsFile, DisputeRequestFile, LetterName From OLDGIS.dbo.NotificationLetters WITH(NOLOCK) Where CustId = @Custid Order By LetterType, StateCd ",
                ColNbr = 17
            });

            Columns.Add(new Column
            {
                Description = "CA-1008/BanTheBox",
                SqlText = "Select CustOptionValue from equestplus.dbo.customeroptions with (nolock) where CustOptionName = 'UseBanTheBoxProcess' and CustID = @Custid and getdate() between EffDate and ExpDate",
                ColNbr = 27
            });

            Columns.Add(new Column
            {
                Description = "Consumer Notice",
                SqlText = "select CustOptionValue from equestplus.dbo.CustomerOptions with  (nolock) where CustOptionName in ('ConsumerNoticeTimeLimit', 'UseConsumerNoticeProcess', 'PrtltrUseStitchProcessing') and CustID = @CustID",
                ColNbr = 26
            });

            Columns.Add(new Column
            {
                Description = "Email Notifications",
                SqlText = string.Empty,
                ColNbr = 28
            });

            Columns.Add(new Column
            {
                Description = "Progression",
                SqlText = string.Empty,
                ColNbr = 42
            });

            Columns.Add(new Column
            {
                Description = "Elink",
                SqlText = "Select SvcCd from EquestPlus.dbo.CustomerEvents with (nolock) where CustID = @Custid and SvcCd = 'CR8ELK' ",
                ColNbr = 46
            });

            Columns.Add(new Column
            {
                Description = "Elink Consent Days",
                SqlText = "Select CustOptionValue from EquestPlus.dbo.CustomerOptions with (nolock) where CustID = @Custid and CustOptionName = 'ConsentAgingDays' ",
                ColNbr = 47
            });
            Columns.Add(new Column
            {
                Description = "Elink Email",
                SqlText = string.Empty, //"Select GroupIDDef from OLDGIS.dbo.EmailGroup with (nolock) where GroupIDDef Like '%' + @Custid + '%'",
                ColNbr = 48
            });

            Columns.Add(new Column
            {
                Description = "Elink Reminder Days",
                SqlText = "Select CustOptionValue from EquestPlus.dbo.CustomerOptions with (nolock) where CustID = @Custid and CustOptionName = 'ReminderGteDays' ",
                ColNbr = 49
            });

            Columns.Add(new Column
            { 
                Description="Billing Pass Thru Fees",
                SqlText="",
                ColNbr=50
            });

            Columns.Add(new Column
            {
                Description = "Custid Validation Errors",
                SqlText = "",
                ColNbr = 51
            });

            Columns.Add(new Column
            {
                Description = "User Validation Errors",
                SqlText = "",
                ColNbr = 52
            });
        }

        private void InsertLinkToValidation(ExcelWorksheet wsSrc, string targetCell)
        {
            wsSrc.Cells[1, 6].Value = "Validation Errors";
            Uri url = new Uri($"#'Validation Errors'!{targetCell}", UriKind.Relative);
            wsSrc.Cells[1, 6].Hyperlink = url;
        }

        private string CheckForWarnings(string custId)
        {
            var sb = new StringBuilder();

            DataTable dt;

            var hlpr2 = new DHMisc(ApplicationWrapper.OldGisConnString);
            dt = hlpr2.GetDatatableFromSQL(custId, ApplicationWrapper.ExportValidationSQL);

            //if (dt.Rows.Count == 0) return string.Empty;
            // list errors from william's sproc
            foreach (DataRow dr in dt.Rows)
            {
                var error = dr["Error"].GetNotDBNull().ToLower();
                if (!error.Contains("more than one email address or invalid email address") &&
                    !error.Contains("missing last name") &&
                    !error.Contains("missing first name"))
                    sb.AppendFormat("{0} - {1}{2}", custId, dr["Error"].GetNotDBNull(), Environment.NewLine);
            }

            // NOTE: do email validation here b/c sql does not have regexes
            // since we have not created the XML nor the client objects
            // get the header to get the billing page that matches the header
            var hdrHlpr = new DHMisc(ApplicationWrapper.OldGisConnString);
            CIFBase approvedCif;
#if DynSQL
            approvedCif = hdrHlpr.GetCifBase(custId, ApplicationWrapper.GetCifBase);
#else
            approvedCif = hdrHlpr.GetCifBase(custId);
#endif
            if (approvedCif != null)
            {
                var cifBaseId = approvedCif.ID;
                var invHlpr = new DHMisc(ApplicationWrapper.OldGisConnString);
                Cif_BillingAndAccounting billing;
#if DynSQL
                billing = invHlpr.GetCifBillingAndAccounting(custId, ApplicationWrapper.GetCifBillingAndAccounting);
#else
                billing = invHlpr.GetCifBillingAndAccounting(custId);
#endif
                // if we could not find billing record, stop crashing:)
                if (billing != null)
                {
                    const string emailPattern = @"^(([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)(\s*;\s*|\s*$))*$";
                    var re = new Regex(emailPattern, RegexOptions.IgnoreCase);
                    const string s = "  is invalid or invalid delimiter.  Multiple email addrs must be separated by ';'. Please correct in eCIF";

                    var m = re.Match(billing.Email);
                    if (!m.Success)
                        sb.AppendFormat("{0} - {1} ({2}) {3}{4}", custId, "Main email", billing.Email.Trim(), s, Environment.NewLine);

                    m = re.Match(billing.AltContactEmail);
                    if (!m.Success)
                        sb.AppendFormat("{0} - {1} ({2}) {3}{4}", custId, "Alternate contact email", billing.AltContactEmail.Trim(), s, Environment.NewLine);

                    m = re.Match(billing.BillToEMail);
                    if (!m.Success)
                        sb.AppendFormat("{0} - {1} ({2}) {3}{4}", custId, "BillTo email", billing.BillToEMail.Trim(), s, Environment.NewLine);
                }
            }
            return sb.ToString();
        }

        private string GetUserErrors(string custId)
        {
            var invHlpr = new DHMisc(ApplicationWrapper.OldGisConnString);
            DataTable usrDt;
#if DynSQL
            var usrHlpr = new DataHelpers.DHMisc(ApplicationWrapper.OldGisConnString);
            usrDt = usrHlpr.GetDatatableFromSQL(custId, ApplicationWrapper.GetUserEmailsForCustid);
#else
            usrDt = invHlpr.GetUsersEmailsForCustid(custId);
#endif
            const string emailPattern = @"^(([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)(\s*;\s*|\s*$))*$";
            var re = new Regex(emailPattern, RegexOptions.IgnoreCase);
            var sb = new StringBuilder();

            foreach (DataRow dr in usrDt.Rows)
            {
                var email = dr["EmailAddress"].GetNotDBNull().ToUpper();
                var userName = dr["UserName"].GetNotDBNull().ToUpper();

                if (string.IsNullOrWhiteSpace(email))
                    sb.AppendFormat("{0} - USER: {1} missing Email Address.{2}", custId, dr["UserName"].GetNotDBNull(), Environment.NewLine);
                else if ((!email.EndsWith("XML") && !email.EndsWith("_LTR") && !email.EndsWith("_QA")) && email == "HIDDEN@GENINFO.COM")
                    sb.AppendFormat("{0} - USER: {1}, 'Hidden@Geninfo.com' is an invalid address. ({2}){3}", custId, dr["UserName"].GetNotDBNull(), email, Environment.NewLine);
                else
                {
                    var m = re.Match(email);
                    if (!m.Success)
                        sb.AppendFormat("{0} - USER: {1}, has invalid email address[es] ({2}){3}", custId, dr["UserName"].GetNotDBNull(), email, Environment.NewLine);
                }

                if (string.IsNullOrWhiteSpace(dr["LastName"].GetNotDBNull()) || string.IsNullOrWhiteSpace(dr["FirstName"].GetNotDBNull()))
                    sb.AppendFormat("{0} - USER: {1} missing First and/or Last Name.{2}", custId, dr["UserName"].GetNotDBNull(), Environment.NewLine);
            }
            return sb.ToString();
        }

        #endregion
    }
    public class EmailNoticeEP
    {
        public ExcelWorksheet WS { get; set; }
        public DataTable DT { get; set; }
        public string Type { get; set; }
        public int RowId { get; set; }
        public int Col1 { get; set; }
        public int Col2 { get; set; }
        public int Col3 { get; set; }
    }
    public class ColHdr
    {
        public string Header { get; set; }
        public double Width { get; set; }
    }
}
*/

