using System;
using System.Collections.Generic;
using BusinessObjects;
using OfficeOpenXml;
using System.IO;
using DataHelpers;
using System.Windows.Forms;
using System.Data;
using System.Diagnostics;
using AppWrapper;
using DevTrackerLogging;
namespace DevTrkrReports
{
    internal class ProjectDetail : Reporter
    {
        #region ..ctor
        public bool Process(List<ProjectNameAndSync> projects, List<DeveloperNames> developers, List<NotableApplication> apps)
        {
            try
            {
                AppList = apps;
                Excel = new ExcelPackage();
                Excel.Workbook.Worksheets.Add("Project Detail Report");
                ExcelFile = new FileInfo(FileName);
                WorkSheet = GetNewWorkBook();
                PopulateSheet(projects, developers, StartTime, EndTime);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ProgramError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        #endregion
        #region private members
        List<NotableApplication> AppList { get; set; }
        private int subCodeLines = 0;
        private int subCommentLines = 0;
        private int subBlankLines = 0;
        private int subDesignerLines = 0;
        private int subTotalAllLines = 0;
        private int totCodeLines = 0;
        private int totCommentLines = 0;
        private int totBlankLines = 0;
        private int totDesignerLines = 0;
        private int totTotalAllLines = 0; 
        #endregion

        private void PopulateSheet(List<ProjectNameAndSync> projects, List<DeveloperNames> developers, DateTime? startTime, DateTime? endTime)
        {
            try
            {
                var lastProject = string.Empty;
                var hlpr = new DHReports(string.Empty);
                var ws = Excel.Workbook.Worksheets[1];
                var sql = CreateSQL(projects, developers, StartTime, EndTime);
                DataSet ds = hlpr.GetDataSetFromSQL(sql);

                var rowId = 0;
                // grand total ctrs
                var mins = 0;
                var hrs = 0;
                var secs = 0;

                // sub total counters
                var subMins = 0;
                var subHrs = 0;
                var subSecs = 0;

                // temp ctrs
                var tmpSecs = 0;
                var tmpMins = 0;
                var tmpHrs = 0;
                //Table 0 is detail
                var dt = ds.Tables[0];
                rowId = 3;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var dr = dt.Rows[i];
                    string appName = dr["appname"].GetNotDBNull();
                    var nameObj = AppList.Find(x => x.AppName == appName);
                    appName = nameObj != null && !string.IsNullOrWhiteSpace(nameObj.AppFriendlyName) ? nameObj.AppFriendlyName : appName;
                    if (i == 0)
                    {
                        lastProject = dt.Rows[i]["Project Name"].GetNotDBNull();
                        ws.Cells[rowId, 1].Value = $"Project: {lastProject}";
                        ws.Cells[rowId, 1].Style.Font.Bold = true;
                        ws.Cells[rowId, 1].Style.Font.UnderLine = true;
                        rowId++;
                    }

                    var thisRowSecs = dr["Seconds"].GetNotDBNullInt();
                    var thisRowMins = dr["Minutes"].GetNotDBNullInt();
                    var thisRowHours = dr["Hours"].GetNotDBNullInt();

                    if (lastProject != dr["Project Name"].GetNotDBNull())
                    {
                        // compute ctrs for sub total and put in the sheet
                        tmpSecs = 0;
                        tmpMins = subMins + Math.DivRem(subSecs, 60, out tmpSecs);
                        tmpHrs = subHrs + Math.DivRem(subMins, 60, out tmpMins);

                        ws.Cells[rowId, 1].Value = "Sub Total";
                        ws.Cells[rowId, 1].Style.Font.Bold = true;
                        ws.Cells[rowId, 2].Value = tmpHrs;
                        ws.Cells[rowId, 3].Value = tmpMins;
                        ws.Cells[rowId, 4].Value = tmpSecs;
                        rowId += 2;
                        lastProject = dr["Project Name"].GetNotDBNull();
                        rowId++;
                        ws.Cells[rowId, 1].Value = $"Project: {lastProject}";
                        ws.Cells[rowId, 1].Style.Font.Bold = true;
                        ws.Cells[rowId, 1].Style.Font.UnderLine = true;
                        rowId++;
                        subHrs = subMins = subSecs = 0; // reset ctrs for next user
                    }

                    // now deal with current row data
                    ws.Cells[rowId, 1].Value = dr["Project Name"].GetNotDBNull();
                    WriteNumberCell(ws, rowId, 2, thisRowHours);
                    ws.Cells[rowId, 3].Value = thisRowMins;
                    ws.Cells[rowId, 4].Value = thisRowSecs;
                    ws.Cells[rowId, 5].Value = dr["UserDisplayName"];
                    ws.Cells[rowId, 6].Value = appName; //dr["appname"].GetNotDBNull().ToUpper();
                    rowId++;
                    // accumulate ctrs this user
                    subMins += thisRowMins;
                    subHrs += thisRowHours;
                    subSecs += thisRowSecs;

                    // accumulate ctrs for grand total
                    hrs += thisRowHours;
                    mins += thisRowMins;
                    secs += thisRowSecs;
                }

                // running out of dt rows may have left us with time unreported for
                // last user
                if (subHrs != 0 || subMins != 0 || subSecs != 0)
                {
                    tmpSecs = 0;
                    tmpMins = subMins + Math.DivRem(subSecs, 60, out tmpSecs);
                    tmpHrs = subHrs + Math.DivRem(subMins, 60, out tmpMins);
                    rowId++;
                    ws.Cells[rowId, 1].Value = "Sub Total";
                    ws.Cells[rowId, 1].Style.Font.Bold = true;
                    WriteNumberCell(ws, rowId, 2, tmpHrs);
                    ws.Cells[rowId, 3].Value = tmpMins;
                    ws.Cells[rowId, 4].Value = tmpSecs;
                    rowId += 1;
                }

                tmpSecs = 0;
                tmpMins = mins + Math.DivRem(secs, 60, out tmpSecs);
                tmpHrs = hrs + Math.DivRem(mins, 60, out tmpMins);


                ws.Cells[rowId + 2, 1].Value = "Time Grand Total";
                ws.Cells[rowId + 2, 1].Style.Font.Bold = true;
                WriteBoldNumberCell(ws, rowId + 2, 2, tmpHrs);
                WriteBoldNumberCell(ws, rowId + 2, 3, tmpMins);
                WriteBoldNumberCell(ws, rowId + 2, 4, tmpSecs);

                // let's play with some number
                int totSlnSeconds = tmpSecs + (tmpMins * 60) + (tmpHrs * 3600);

                // the time portion of the report is complete, write out the files part
                // write the header Row
                rowId += 5;
                lastProject = string.Empty;
                dt = ds.Tables[1];
                if (dt.Rows.Count > 0)
                {
                    //CreateHdrBoldStyle(ws, HdrRange);
                    WriteBoldCell(ws, rowId, 1, "DevProjectName");
                    WriteBoldCell(ws, rowId, 2, "Relative FileName");
                    WriteBoldCell(ws, rowId, 3, "CodeLines");
                    WriteBoldCell(ws, rowId, 4, "CommentLines");
                    WriteBoldCell(ws, rowId, 5, "BlankLines");
                    WriteBoldCell(ws, rowId, 6, "DesignerLines");
                    WriteBoldCell(ws, rowId, 7, "AllLines");
                    WriteBoldCell(ws, rowId, 8, "UpdateCount (For Binary Files = Rebuilds)");
                    rowId++;

                    string currProject;
                    rowId++;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dr = dt.Rows[i];
                        if (i.Equals(0))
                        {
                            lastProject = dr["DevProjectName"].GetNotDBNull();
                            //WriteBoldCell(ws, rowId, 1, lastProject);
                            //rowId++;
                        }

                        currProject = (string)dr["DevProjectName"];
                        if (lastProject != currProject)
                        {
                            // compute ctrs for sub total and put in the sheet
                            lastProject = currProject;
                            WriteBoldCell(ws, rowId, 1, "Sub Total");
                            WriteNumberCell(ws, rowId, 3, subCodeLines);
                            WriteNumberCell(ws, rowId, 4, subCommentLines);
                            WriteNumberCell(ws, rowId, 5, subBlankLines);
                            WriteNumberCell(ws, rowId, 6, subDesignerLines);
                            WriteNumberCell(ws, rowId, 7, subTotalAllLines);
                            rowId += 2;
                            subCodeLines = subCommentLines = subBlankLines = subDesignerLines = subTotalAllLines = 0;
                        }

                        // now deal with current row
                        WriteLinesRow(dr, ws, rowId);
                        rowId++;
                    }

                    // running out of dt rows may have left us with time unreported for
                    // last user
                    if (subCodeLines != 0 || subCommentLines != 0 || subBlankLines != 0 || subDesignerLines != 0)
                    {
                        WriteBoldCell(ws, rowId, 1, "Sub Total");
                        WriteNumberCell(ws, rowId, 3, subCodeLines);
                        WriteNumberCell(ws, rowId, 4, subCommentLines);
                        WriteNumberCell(ws, rowId, 5, subBlankLines);
                        WriteNumberCell(ws, rowId, 6, subDesignerLines);
                        WriteNumberCell(ws, rowId, 7, subTotalAllLines);
                        rowId += 2;
                    }

                    // print grand totals
                    WriteBoldCell(ws, rowId, 1, "Code Lines Grand Totals");
                    WriteBoldNumberCell(ws, rowId, 3, totCodeLines);
                    WriteBoldNumberCell(ws, rowId, 4, totCommentLines);
                    WriteBoldNumberCell(ws, rowId, 5, totBlankLines);
                    WriteBoldNumberCell(ws, rowId, 6, totDesignerLines);
                    WriteBoldNumberCell(ws, rowId, 7, totTotalAllLines);

                    // now, continue playing with numbers
                    // designer generated lines are not free b/c the develope
                    // used the designer to build the form, etc. and the generated
                    // lines were a result of that work
                    var secsPerTotalLines = totSlnSeconds / totTotalAllLines;
                    var secsPerCodeLine = totSlnSeconds / totCodeLines;

                    rowId += 2;
                    WriteBoldCell(ws, rowId, 1, "Seconds Per All Type Lines:");
                    WriteBoldNumberCell(ws, rowId, 7, secsPerTotalLines);
                    rowId++;
                    WriteBoldCell(ws, rowId, 1, "Seconds Per Just Code Lines:");
                    WriteBoldNumberCell(ws, rowId, 3, secsPerCodeLine);

                    // get elapsed time from the project/solution
                    var stTime = (DateTime)ds.Tables[2].Rows[0]["StartTime"];
                    var eTime = (DateTime)ds.Tables[3].Rows[0]["EndTime"];
                    var elpTime = ((eTime - stTime).TotalDays) / 7 * 5 * 8 * 3600;
                    TimeSpan elapsed = eTime.Subtract(stTime);
                    // Get number of days ago.
                    double daysAgo = elapsed.TotalDays / 7 * 5;
                    var sd = stTime.ToString("MM/dd/yyyy");
                    var ed = eTime.ToString("MM/dd/yyyy");
                    WriteBoldCell(ws, ++rowId, 1, $"Coding Started: {sd}");
                    WriteBoldCell(ws, ++rowId, 1, $"Last Coding: {ed}");
                    WriteBoldCell(ws, ++rowId, 1, $"Elapsed Work Days: {Math.Truncate(daysAgo)}");
                }

                Excel.SaveAs(ExcelFile);
                Excel.Dispose();
            }
            catch (Exception ex)
            {
                _ = new LogError(ex.Message, false, "ProjectDetail.PopulateSheet");
            }
        }

        private void WriteLinesRow(DataRow dr, ExcelWorksheet ws, int rowId)
        {
            var tmpCodeLines = dr["CodeLines"].GetNotDBNullInt();
            var tmpCommLines = dr["CommentLines"].GetNotDBNullInt();
            var tmpBlankLines = dr["BlankLines"].GetNotDBNullInt();
            var tmpDesgLines = dr["DesignerLines"].GetNotDBNullInt();
            var tmpCodeTotal = tmpCodeLines + tmpCommLines + tmpBlankLines + tmpDesgLines;
            subCodeLines += tmpCodeLines;
            totCodeLines += tmpCodeLines;
            subCommentLines += tmpCommLines;
            totCommentLines += tmpCommLines;
            subBlankLines += tmpBlankLines;
            totBlankLines += tmpBlankLines;
            subDesignerLines += tmpDesgLines;
            totDesignerLines += tmpDesgLines;
            subTotalAllLines += tmpCodeTotal;
            totTotalAllLines += tmpCodeTotal;
            ws.Cells[rowId, 1].Value = dr["DevProjectName"].GetNotDBNull();
            ws.Cells[rowId, 2].Value = dr["RelativeFileName"].GetNotDBNull();
            WriteNumberCell(ws, rowId, 3, tmpCodeLines);
            WriteNumberCell(ws, rowId, 4, tmpCommLines);
            WriteNumberCell(ws, rowId, 5, tmpBlankLines);
            WriteNumberCell(ws, rowId, 6, tmpDesgLines);
            WriteNumberCell(ws, rowId, 7, tmpCodeTotal);
            WriteNumberCell(ws, rowId, 8, dr["UpdateCount"].GetNotDBNullInt());
        }

        /// <summary>
        /// Dynamic SQL either has to be built here or in a stored procedure.
        /// It is easier here and just as safe b/c there are no parameters.
        /// </summary>
        private string CreateSQL(List<ProjectNameAndSync> projects, List<DeveloperNames> developers, DateTime? startTime, DateTime? endTime)
        {
            string sql =
            "SET NOCOUNT ON; " +
            "select DevProjectName as [Project Name], " +
            "sum(TotalSeconds) / 3600 as Hours, " +
            "(sum(TotalSeconds) % 3600) / 60 as Minutes, " +
            "sum(TotalSeconds) % 60 as Seconds " +
            ", UserName, UserDisplayName, appname " +
            "from " +
            "(select DevProjectName, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            ", UserName, UserDisplayName, appname " +
            "from DevTrkr..WindowEvents w " +
            "where 1=1 " +
            GetListSQL(projects) +
            GetListSQL(developers, "UserName") +
            GetDateSQL(startTime, endTime, true) +
            ") x " +
            "Group by UserName, UserDisplayName, DevProjectName, appname " +
            "Order by DevProjectName, UserName "; // +

            sql += "select * from DevTrkr..ProjectFiles with (nolock) " +
                "where 1=1 " +
                GetListSQL(projects) +
                GetListSQL(developers, "UserName") +
                GetDateSQL(startTime, endTime, true) +
                "Order by DevProjectName, RelativeFileName ";

            sql += "select top 1 starttime from DevTrkr..WindowEvents with (nolock) " +
                   "where 1=1 " +
                    GetListSQL(projects) +
                    GetListSQL(developers, "UserName") +
                    GetDateSQL(startTime, endTime, true) +
                    "Order by starttime ";

            sql += "select top 1 endtime from DevTrkr..WindowEvents with (nolock) " +
                   "where 1=1 " +
                    GetListSQL(projects) +
                    GetListSQL(developers, "UserName") +
                    GetDateSQL(startTime, endTime, true) +
                    "Order by endtime desc ";

            return sql;
        }

    }
}
