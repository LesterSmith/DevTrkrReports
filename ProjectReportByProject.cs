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
namespace DevTrkrReports
{
    internal class ProjectReportByProject : Reporter
    {
        public bool Process(List<ReportProjects> projects, List<DeveloperNames> developers) //, DateTime startTime, DateTime endTime)
        {
            try
            {
                Excel = new ExcelPackage();
                Excel.Workbook.Worksheets.Add("Project Report");
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

        private void PopulateSheet(List<ReportProjects> projects, List<DeveloperNames> developers, DateTime? startTime, DateTime? endTime)
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

                    // mow deal with current row data
                    ws.Cells[rowId, 1].Value = dr["Project Name"].GetNotDBNull();
                    ws.Cells[rowId, 2].Value = thisRowHours;
                    ws.Cells[rowId, 3].Value = thisRowMins;
                    ws.Cells[rowId, 4].Value = thisRowSecs;
                    ws.Cells[rowId, 5].Value = dr["UserDisplayName"];
                    ws.Cells[rowId, 6].Value = GetPrjPath((string)ws.Cells[rowId, 1].Value, projects);
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
                    //secs += subSecs;
                    //mins += subMins;
                    //hrs += subHrs;
                    tmpSecs = 0;
                    tmpMins = subMins + Math.DivRem(subSecs, 60, out tmpSecs);
                    tmpHrs = subHrs + Math.DivRem(subMins, 60, out tmpMins);
                    rowId++;
                    ws.Cells[rowId, 1].Value = "Sub Total";
                    ws.Cells[rowId, 1].Style.Font.Bold = true;
                    ws.Cells[rowId, 2].Value = tmpHrs;
                    ws.Cells[rowId, 3].Value = tmpMins;
                    ws.Cells[rowId, 4].Value = tmpSecs;
                    rowId += 1;
                }

                tmpSecs = 0;
                tmpMins = mins + Math.DivRem(secs, 60, out tmpSecs);
                tmpHrs = hrs + Math.DivRem(mins, 60, out tmpMins);


                ws.Cells[rowId + 2, 1].Value = "Grand Total";
                ws.Cells[rowId + 2, 1].Style.Font.Bold = true;
                ws.Cells[rowId + 2, 2].Value = tmpHrs;
                ws.Cells[rowId + 2, 3].Value = tmpMins;
                ws.Cells[rowId + 2, 4].Value = tmpSecs;

                Excel.SaveAs(ExcelFile);
                Excel.Dispose();
            }
            catch (Exception ex)
            {
                Util.LogError(ex, true);
            }
        }

        /// <summary>
        /// 
        /// Dynamic SQL either has to be built here or in a stored procedure.
        /// It is easier here and just as safe b/c there are no parameters.
        /// </summary>
        private string CreateSQL(List<ReportProjects> projects, List<DeveloperNames> developers, DateTime? startTime, DateTime? endTime)
        {
            string sql =
            "SET NOCOUNT ON; " +
            "select DevProjectName as [Project Name], " +
            "sum(TotalSeconds) / 3600 as Hours, " +
            "(sum(TotalSeconds) % 3600) / 60 as Minutes, " +
            "sum(TotalSeconds) % 60 as Seconds " +
            ", UserName, UserDisplayName " +
            "from " +
            "(select DevProjectName, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            ", UserName, UserDisplayName " +
            "from DevTrkr..WindowEvents w " +
            "where 1=1 " +
            GetListSQL(projects, "DevProjectName") +
            GetListSQL(developers, "UserName") +
            GetDateSQL(startTime, endTime, true) +
            ") x " +
            "Group by UserName, UserDisplayName, DevProjectName " +
            "Order by DevProjectName, UserName " +
            "select 'Total' as [Total], " +
            "sum(TotalSeconds) / 3600 as Hours, " +
            "(sum(TotalSeconds) % 3600) / 60 as Minutes " +
            ", sum(TotalSeconds) % 60 as Seconds " +
            "from " +
            "( " +
            "select DevProjectName, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            "from DevTrkr..WindowEvents w " +
            "where 1=1 " +
            GetListSQL(projects, "DevProjectName") +
            GetListSQL(developers, "UserName") +
            GetDateSQL(startTime, endTime, true) +
            ") x ";
            return sql;
        }

        private string GetDevelopersSQL(List<string> developers)
        {
            return developers.Count.Equals(1) && developers[0].Equals("All") ? "\"All\" as Developer " : "UserName as Developer, UserDisplayName as DevFullName ";
        }
    }
}
