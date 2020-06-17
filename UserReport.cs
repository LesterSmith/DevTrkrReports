using System;
using System.Collections.Generic;
using BusinessObjects;
using OfficeOpenXml;
using System.IO;
using DataHelpers;
using System.Windows.Forms;
using System.Data;
using AppWrapper;
namespace DevTrkrReports
{
    /// <summary>
    /// This report is meant to show the time spent by a developer
    /// during a Day, Week, or Month
    /// It will map together the time spent charged to projects, scheduled to
    /// be spent in meetings, and time actually spent in Outlook.
    /// Basically that would cover all time except for time spent on the phone
    /// or at the water fountain...
    /// So, it should have a parameter to tell us how to summarize the data
    /// or group by, e.g. group by Startdate
    /// because of the volume of data, we will have to 
    /// </summary>
    public class UserReport : Reporter
    {
        public bool Process(List<ReportProjects> projects, List<DeveloperNames> developers)
        {
            try
            {
                Excel = new ExcelPackage();
                Excel.Workbook.Worksheets.Add("Developer Detail");
                ExcelFile = new FileInfo(FileName);
                WorkSheet = GetNewWorkBook();
                PopulateSheet(projects, developers); //, startTime, endTime);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ProgramError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void PopulateSheet(List<ReportProjects> projects, List<DeveloperNames> developers) //, DateTime? startTime, DateTime? endTime)
        {
            try
            {
                var lastUser = string.Empty;
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
                var dt = ds.Tables[0];
                
                rowId = 3;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var dr = dt.Rows[i];
                    if (i == 0)
                    {
                        lastUser = dt.Rows[i]["UserDisplayName"].GetNotDBNull();
                        ws.Cells[rowId, 1].Value = $"User: {lastUser}";
                        ws.Cells[rowId, 1].Style.Font.Bold = true;
                        ws.Cells[rowId, 1].Style.Font.UnderLine = true;
                        rowId++;
                    }

                    var thisRowSecs = dr["Seconds"].GetNotDBNullInt();
                    var thisRowMins = dr["Minutes"].GetNotDBNullInt(); 
                    var thisRowHours = dr["Hours"].GetNotDBNullInt();

                    if (lastUser != dr["UserDisplayName"].GetNotDBNull())
                    {
                        // compute ctrs for sub total and put in the sheet
                        tmpSecs = 0;
                        tmpMins = subMins + Math.DivRem(subSecs, 60, out tmpSecs);
                        tmpHrs = subHrs + Math.DivRem(subMins, 60, out tmpMins);

                        ws.Cells[rowId, 1].Value = "Sub Total";
                        ws.Cells[rowId, 2].Value = tmpHrs;
                        ws.Cells[rowId, 3].Value = tmpMins;
                        ws.Cells[rowId, 4].Value = tmpSecs;
                        rowId += 2;
                        lastUser = dr["UserDisplayName"].GetNotDBNull();
                        rowId++;
                        ws.Cells[rowId, 1].Value = $"User: {lastUser}";
                        ws.Cells[rowId, 1].Style.Font.Bold = true;
                        ws.Cells[rowId, 1].Style.Font.UnderLine = true;
                        rowId++;
                        subHrs = subMins = subSecs = 0; // reset ctrs for next user
                    }

                    // mow deal with current row data
                    ws.Cells[rowId, 1].Value = dr["DevProjectName"].GetNotDBNull();
                    ws.Cells[rowId, 2].Value = thisRowHours;
                    ws.Cells[rowId, 3].Value = thisRowMins;
                    ws.Cells[rowId, 4].Value = thisRowSecs;
                    ws.Cells[rowId, 5].Value = dr["UserDisplayName"].GetNotDBNull();
                    ws.Cells[rowId, 6].Value = dr["Activity"].GetNotDBNull();
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
                Util.LogError(ex,true);
            }
        }

        private string GetMeetingData(DataRow dr)
        {
            if (!dr["Activity"].GetNotDBNull().StartsWith("Meeting"))
                return string.Empty;
            var recur = dr["Recurring"].GetNotDBNullBool() ? "Yes" : "No";
            return $"Meeting Organized By: {dr["Organizer"].GetNotDBNull()} Recurring: {recur}";
        }
        private string CreateSQL(List<ReportProjects> projects, List<DeveloperNames> developers, DateTime? startTime, DateTime? endTime)
        {
            var hlpr = new DHWindowEvents(string.Empty);
            string userName = developers[0].UserName;
            string userDisplayName = hlpr.GetUserDisplayName(userName, Environment.MachineName);
            //DevProjectName, Hours, Minutes, UserDisplayName, Activity
            string sql = string.Empty;
            // **** Project related code ****
            sql += "SELECT DevProjectName, sum(TotalSeconds) / 3600 as Hours, " +
            "(sum(TotalSeconds) % 3600) / 60 as Minutes, sum(TotalSeconds) % 60 as Seconds, UserDisplayName , 'Development' as Activity " +
            "from " +
            "(select DevProjectName, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            //", Username, UserDisplayName " +
            ", UserDisplayName " +
            "from DevTrkr..WindowEvents w " +
            "where 1 = 1 " +
            GetListSQL(developers, "username") +
            GetDateSQL(StartTime, EndTime, true) +
            GetListSQL(projects, "DevProjectName") +
            "and DevProjectName in (Select DevProjectName from DevTrkr..DevProjects with (nolock) where isnull(DatabaseProject,0)  != 1)" +
            ") as x " +
            "Group by UserDisplayName, DevProjectName  " +

            "Union " +
            // *** Meetings ****
            "Select Subject as DevProjectName, sum(TotalSeconds) / 3600 as Hours, sum(TotalSeconds) % 3600 / 60 as Minutes, sum(TotalSeconds) % 60 as Seconds, " +
            "UserDisplayName, 'Meeting' as Activity " +
            //"Case Recurring = 1 then 'Meetings (Recurring)' Else 'Meetings (Scheduled)' End as Activity " +
            "from " +
            "(select [Subject] + ' (Organizer: ' + Organizer + '  Recurs: ' + Case when Recurring = 1 then 'Yes' else 'No' end + ')' as Subject, DATEDIFF(second, StartTime, EndTime) as TotalSeconds, " +
            "UserDisplayName " +
            "from DevTrkr..Meetings m with(nolock) " +
            "where 1=1 " +
            GetDateSQL(StartTime, EndTime, true) +
            GetListSQL(developers, "username") +
            //"and username = '" + userName + "'" + 
            " ) as x " +
            "Group by UserDisplayName, Subject  " +

            "Union " +

            // ****  Time spent in Email/Outlook Only ***
            "Select DevProjectName, sum(TotalSeconds)/ 3600 as Hours, sum(TotalSeconds) % 3600 / 60 as Minutes, sum(TotalSeconds) % 60 as Seconds, " +
            //"'" + userDisplayName + "' as UserDisplayName , 'EMail' as Activity " +
            "UserDisplayName , 'EMail' as Activity " +
            "from " +
            "(select DevProjectName, DATEDIFF(second, StartTime, EndTime) as TotalSeconds, " +
            "'" + userDisplayName + "' as UserDisplayName " +
            "from DevTrkr..WindowEvents w with(nolock) Where 1=1 " +
            GetListSQL(developers, "username") +
            GetDateSQL(StartTime, EndTime, true) +
            //"and username = '" + userName + "' " +
            "and AppName = 'outlook' " +
            ") as x " +
            "Group by UserDisplayName, DevProjectName  " +

            "Union " +
            // time spent in database work
            "Select DevProjectName, sum(TotalSeconds)/ 3600 as Hours, sum(TotalSeconds) % 3600 / 60 as Minutes, sum(TotalSeconds) % 60 as Seconds, " +
            //"'" + userDisplayName + "' as UserDisplayName, 'Database' as Activity " +
            "UserDisplayName, 'Database' as Activity " +
            "from " +
            "(select DevProjectName, DATEDIFF(second, StartTime, EndTime) as TotalSeconds, " +
             //"'" + userDisplayName + "' as UserDisplayName " +
             "UserDisplayName " +
            "from DevTrkr..WindowEvents w with(nolock) " +
            "where 1=1 " +
            GetListSQL(developers, "username") +
            GetDateSQL(StartTime, EndTime, true) +
            "and DevProjectName in (Select DevProjectName from DevTrkr..DevProjects with (nolock) where DatabaseProject  = 1)" +
            ") as x " +
            "Group by UserDisplayName, DevProjectName  " +

            "Union " +
            //TODO **** Documentation **** need generic code rather than hard coded appnames below
            "Select DevProjectName, sum(TotalSeconds)/ 3600 as Hours, sum(TotalSeconds) % 3600 / 60 as Minutes, sum(TotalSeconds) % 60 as Seconds, " +
            //"'" + userDisplayName + "' as UserDisplayName, 'Documentation' as Activity " +
            "UserDisplayName, 'Documentation' as Activity " +
            "from " +
            "(select DevProjectName, DATEDIFF(second, StartTime, EndTime) as TotalSeconds, " +
            //"'" + userDisplayName + "' as UserDisplayName " +
            "UserDisplayName " +
            "from DevTrkr..WindowEvents w with(nolock) " +
            "where 1=1 " +
            GetListSQL(developers, "username") +
            //"and username = '" + userName + "' " +
            "and AppName in ('winword', 'notepad++', 'notepad', 'wordpad') " +
            ") as x " +
            "Group by UserDisplayName, DevProjectName  " +

            "Union " +
            // **** Computer Locked ******
            "SELECT DevProjectName, sum(TotalSeconds) / 3600 as Hours, " +
            "(sum(TotalSeconds) % 3600) / 60 as Minutes, sum(TotalSeconds) % 60 as Seconds, UserDisplayName , 'Development' as Activity " +
            "from " +
            "(select DevProjectName, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            ", Username, UserDisplayName " +
            "from DevTrkr..WindowEvents w " +
            "where 1 = 1 " +
            "and DevProjectName = 'ComputerLocked' " +
            GetListSQL(developers, "username") +
            GetDateSQL(StartTime, EndTime, true) +
            ") as x " +
            "Group by UserDisplayName, DevProjectName  " +
            "Order by UserDisplayName, Activity, DevProjectName ";
            return sql;
        }
    }
}
