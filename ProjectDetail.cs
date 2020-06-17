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
    internal class ProjectDetail : Reporter
    {
        public bool Process(List<ProjectNameAndSync> projects, List<DeveloperNames> developers)
        {
			try
			{
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
                    ws.Cells[rowId, 2].Value = thisRowHours;
                    ws.Cells[rowId, 3].Value = thisRowMins;
                    ws.Cells[rowId, 4].Value = thisRowSecs;
                    ws.Cells[rowId, 5].Value = dr["UserDisplayName"];
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

                // let's play with some number
                int totSlnSeconds = tmpSecs + (tmpMins * 60) + (tmpHrs * 3600);

                // the time portion of the report is complete, write out the files part
                // write the header Row
                rowId += 5;
                lastProject = string.Empty;
                //CreateHdrBoldStyle(ws, HdrRange);
                WriteBoldCell(ws, rowId, 1, "DevProjectName");
                WriteBoldCell(ws, rowId, 2, "Relative FileName");
                WriteBoldCell(ws, rowId, 3, "CodeLines");
                WriteBoldCell(ws, rowId, 4, "CommentLines");
                WriteBoldCell(ws, rowId, 5, "BlankLines");
                WriteBoldCell(ws, rowId, 6, "DesignerLines");
                WriteBoldCell(ws, rowId, 7, "AllLines");
                WriteBoldCell(ws, rowId, 8, "UpdateCount");
                rowId++;

                string currProject;
                dt = ds.Tables[1];
                rowId++;
                for (int i= 0; i < dt.Rows.Count; i++)
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
                        ws.Cells[rowId, 2].Value = "Sub Total";
                        ws.Cells[rowId, 3].Value = subCodeLines;
                        ws.Cells[rowId, 4].Value = subCommentLines;
                        ws.Cells[rowId, 5].Value = subBlankLines;
                        ws.Cells[rowId, 6].Value = subDesignerLines;
                        ws.Cells[rowId, 7].Value = subTotalAllLines;
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
                    ws.Cells[rowId, 1].Value = "Sub Total";
                    ws.Cells[rowId, 3].Value = subCodeLines;
                    ws.Cells[rowId, 4].Value = subCommentLines;
                    ws.Cells[rowId, 5].Value = subBlankLines;
                    ws.Cells[rowId, 6].Value = subDesignerLines;
                    ws.Cells[rowId, 7].Value = subTotalAllLines;
                    rowId += 2;
                }

                // print grand totals
                ws.Cells[rowId, 1].Value = "Grand Totals";
                ws.Cells[rowId, 3].Value = totCodeLines;
                ws.Cells[rowId, 4].Value = totCommentLines;
                ws.Cells[rowId, 5].Value = totBlankLines;
                ws.Cells[rowId, 6].Value = totDesignerLines;
                ws.Cells[rowId, 7].Value = totTotalAllLines;

                // now, continue playing with numbers
                // designer generated lines are not free b/c the develope
                // used the designer to build the form, etc. and the generated
                // lines were a result of that work
                var secsPerTotalLines = totSlnSeconds/totTotalAllLines;
                rowId += 2;
                WriteBoldCell(ws, rowId, 1, "Seconds Per Line:");
                WriteBoldCell(ws, rowId, 2, secsPerTotalLines.ToString());

                Excel.SaveAs(ExcelFile);
                Excel.Dispose();
            }
            catch (Exception ex)
            {
                Util.LogError(ex.Message);
            }
        }

        private void WriteBoldCell(ExcelWorksheet ws, int rowId, int cell, string value)
        {
            ws.Cells[rowId, cell].Value = value;
            ws.Cells[rowId, cell].Style.Numberformat.Format = "@";
            ws.Cells[rowId, cell].Style.Font.Bold = true;
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
            ws.Cells[rowId, 3].Value  = tmpCodeLines;
            ws.Cells[rowId, 4].Value  = tmpCommLines;
            ws.Cells[rowId, 5].Value  = tmpBlankLines;
            ws.Cells[rowId, 6].Value  = tmpDesgLines;
            ws.Cells[rowId, 7].Value = tmpCodeTotal;
            ws.Cells[rowId, 8].Value  = dr["UpdateCount"].GetNotDBNullInt();
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
            ", UserName, UserDisplayName " +
            "from " +
            "(select DevProjectName, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            ", UserName, UserDisplayName " +
            "from DevTrkr..WindowEvents w " +
            "where 1=1 " +
            GetListSQL(projects) +
            GetListSQL(developers, "UserName") +
            GetDateSQL(startTime, endTime, true) +
            ") x " +
            "Group by UserName, UserDisplayName, DevProjectName " +
            "Order by DevProjectName, UserName "; // +
            //"select 'Total' as [Total], " +
            //"sum(TotalSeconds) / 3600 as Hours, " +
            //"(sum(TotalSeconds) % 3600) / 60 as Minutes " +
            //", sum(TotalSeconds) % 60 as Seconds " +
            //"from " +
            //"( " +
            //"select DevProjectName, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            //"from DevTrkr..WindowEvents w " +
            //"where 1=1 " +
            //GetListSQL(projects) +
            //GetListSQL(developers, "UserName") +
            //GetDateSQL(startTime, endTime, true) +
            //") x ";

            sql += "select * from DevTrkr..ProjectFiles with (nolock) " +
                "where 1=1 " +
                GetListSQL(projects) +
                GetListSQL(developers, "UserName") +
                GetDateSQL(startTime, endTime, true) +
                "Order by DevProjectName, RelativeFileName ";
            return sql;
        }

    }
}
