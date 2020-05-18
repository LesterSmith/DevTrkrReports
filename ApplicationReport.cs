using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessObjects;
using OfficeOpenXml;
using System.IO;
using DataHelpers;
using System.Windows.Forms;
using System.Data;

namespace DevTrkrReports
{
    class ApplicationReport : Reporter
    {
        public bool Process(List<NotableApplication> apps, List<DeveloperNames> developers)
        {
            try
            {
                Excel = new ExcelPackage();
                Excel.Workbook.Worksheets.Add("Application Usage Report");
                ExcelFile = new FileInfo(FileName);
                WorkSheet = GetNewWorkBook();
                PopulateSheet(apps, developers); //, startTime, endTime);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ProgramError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void PopulateSheet(List<NotableApplication> apps, List<DeveloperNames> developers) //, DateTime? startTime, DateTime? endTime)
        {
            var hlpr = new DHReports(string.Empty);
            var ws = Excel.Workbook.Worksheets[1];
            var sql = CreateSQL(apps, developers, StartTime, EndTime);
            DataSet ds = hlpr.GetDataSetFromSQL(sql);

            var rowId = 0;
            //Table 0 is detail
            var dt = ds.Tables[0];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowId = i + 3;
                ws.Cells[rowId, 1].Value = dt.Rows[i]["Application Name"].GetNotDBNull();
                ws.Cells[rowId, 2].Value = dt.Rows[i]["Hours"].GetNotDBNullInt();
                ws.Cells[rowId, 3].Value = dt.Rows[i]["Minutes"].GetNotDBNullInt();
                ws.Cells[rowId, 4].Value = dt.Rows[i]["Seconds"].GetNotDBNullInt();
                ws.Cells[rowId, 5].Value = dt.Rows[i]["UserDisplayName"];
                //ws.Cells[rowId, 5].Value = GetPrjPath((string)ws.Cells[rowId, 1].Value, apps);
            }

            if (ds.Tables[1].Rows.Count > 0)
            {
                ws.Cells[rowId + 2, 1].Value = ds.Tables[1].Rows[0]["Total"].GetNotDBNull();
                ws.Cells[rowId + 2, 2].Value = ds.Tables[1].Rows[0]["Hours"].GetNotDBNullInt();
                ws.Cells[rowId + 2, 3].Value = ds.Tables[1].Rows[0]["Minutes"].GetNotDBNullInt();
                ws.Cells[rowId + 2, 4].Value = ds.Tables[1].Rows[0]["Seconds"].GetNotDBNullInt();
            }
            Excel.SaveAs(ExcelFile);
            Excel.Dispose();
        }

        private string CreateSQL(List<NotableApplication> apps, List<DeveloperNames> developers, DateTime? startTime, DateTime? endTime)
        {
            string sql = "SET NOCOUNT ON; " +
            "select AppFriendlyName as [Application Name], " +
            "sum(TotalSeconds) / 3600 as Hours, " +
            "(sum(TotalSeconds) % 3600) / 60 as Minutes " +
            ", sum(TotalSeconds) % 60 as Seconds " +
            ", UserName, UserDisplayName " +
            "from " +
            "(select AppFriendlyName, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            ", UserName, UserDisplayName " +
            "from DevTrkr..WindowEvents w " +
            "inner join DevTrkr..NotableApplications na on w.AppName = na.AppName " +
            "where 1=1 " +
            GetListSQL(apps, "appname") +
            GetListSQL(developers, "UserName") +
            GetDateSQL(startTime, endTime) +
            ") x " +
            "Group by AppFriendlyName, UserName, UserDisplayName " +
            "Order by AppFriendlyName, UserName " +
            "select 'Total' as [Total], " +
            "sum(TotalSeconds) / 3600 as Hours, " +
            "(sum(TotalSeconds) % 3600) / 60 as Minutes " +
            ", sum(TotalSeconds) % 60 as Seconds " +
            "from " +
            "( " +
            "select appname, DateDiff(second, StartTime, EndTime) as TotalSeconds " +
            "from DevTrkr..WindowEvents w " +
            "where 1=1 " +
            GetListSQL(apps, "appname") +
            GetListSQL(developers, "UserName") +
            GetDateSQL(startTime, endTime) +
            ") x ";
            return sql;
        }

    }
}
