using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using BusinessObjects;
namespace DevTrkrReports
{
    public abstract class Reporter : IDisposable
    {
        #region ..ctors

        #endregion // ..ctor


        #region public and private members
        public const string ProgramError = "Program Error";
        const string yes = "Yes";
        const string no = "No";
        private int NbrColumns { get; set; }
        public List<ColHdr> ColHdrs { get; set; }
        public string FileName { get; set; }
        public string Title { get; set; }
        public string TitleCell { get; set; }
        public string HdrRange { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public ExcelWorksheet WorkSheet { get; set; }
        public ExcelPackage Excel { get; set; }
        public FileInfo ExcelFile { get; set; }
        //private int RowId { get; set; }
        //private int ValidCell { get; set; }
        public string[] Projects { get; set; }
        public string[] Developers { get; set; }
        #endregion // public and private members

        #region public methods
        protected internal string GetPrjPath(string value, List<DevProjPath> projects)
        {
            var prj = projects.Find(x => x.DevProjectName == value);
            return prj != null ? prj.DevProjectPath : "Unknown";
        }

        protected internal string GetDateSQL(DateTime? startTime, DateTime? endTime, bool dateOnly=false)
        {
            if (startTime == null)
                return string.Empty;

            return "and (starttime >= '" + (!dateOnly ? startTime.Value.Date.ToString("MM/dd/yyyy HH:mm:ss") : startTime.Value.Date.ToString("MM/dd/yyyy")) + "' and endtime <= '" + (!dateOnly ? endTime.Value.Date.ToString("MM/dd/yyyy HH:mm:ss") : endTime.Value.Date.ToString("MM/dd/yyyy")) + "') ";
        }

        protected internal string GetListSQL(List<DevProjPath> projects, string field)
        {
            const string comma = ",";
            bool first = true;
            const string singQte = "'";
            //if (projects.Count.Equals(1) && projects[0].Equals("All"))
            //    return string.Empty; // all projects, no where clause for projects
            var allProjects = projects.Count.Equals(1) && projects[0].Equals("All");
            var sql = $"and {field} in (";
            for (int i = 0; i < projects.Count; i++)
            {
                if (projects[i].Selected || allProjects)
                {
                    sql += first ? singQte + projects[i].DevProjectName + singQte : comma + singQte + projects[i].DevProjectName + singQte;
                    first = false;
                }
            }
            sql += ") ";
            return sql;
        }
        //TODO this method does not line up with the frmReporter nothing is selected in the incoming list
        /// <summary>
        /// 
        /// </summary>
        /// <param name="developers"></param>
        /// <param name="field"></param>
        /// <returns></returns>
        protected internal string GetListSQL(List<DeveloperNames> developers, string field)
        {
            const string comma = ",";
            bool first = true;
            const string singQte = "'";
            if (developers.Count.Equals(1) && developers[0].UserName.Equals("All"))
                return string.Empty; // all projects, no where clause for developers

            var sql = $"and {field} in (";
            for (int i = 0; i < developers.Count; i++) 
            {
                if (developers[i].Selected)
                {
                    sql += first ? singQte + developers[i].UserName + singQte : comma + singQte + developers[i].UserName + singQte;
                    first = false;
                }
            }
            sql += ") ";
            return sql;
        }

        protected internal string GetListSQL(List<NotableApplication> apps, string field)
        {
            const string comma = ",";
            const string singQte = "'";
            if (apps.Count.Equals(0))
                return string.Empty; // all applications, no where clause for applications

            var sql = $"and w.{field} in (";
            for (int i = 0; i < apps.Count; i++)
            {
                sql += i == 0 ? singQte + apps[i].AppName + singQte : comma + singQte + apps[i].AppName + singQte;
            }
            sql += ") ";
            return sql;
        }


        #endregion // public methods

        #region private methods
        public ExcelWorksheet GetNewWorkBook()
        {
            try
            {
                var ws = Excel.Workbook.Worksheets[1];
                CreateHdrBoldStyle(ws, HdrRange);
                ws.Cells[TitleCell].Value = Title;
                for (var i = 0; i < ColHdrs.Count; i++)
                {
                    var hdr = ColHdrs[i];
                    ws.Cells[2, i + 1].Value = hdr.Header;
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

        public void CreateHdrBoldStyle(ExcelWorksheet ws, string hdrRange, int titleHeight=15)
        {
            try
            {
                ws.Cells[hdrRange].Style.Numberformat.Format = "@";
                ws.Cells[hdrRange].Style.Font.Bold = true;
                ws.Cells[hdrRange].Style.Font.Size = 14;
                ws.Row(1).Height = titleHeight;
                ws.Cells[hdrRange].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
               // ws.Cells[hdrRange].Style.WrapText = true;
                ws.Cells[hdrRange].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[hdrRange].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
                ws.Cells[hdrRange].Style.Font.Color.SetColor(Color.White);
                ws.View.FreezePanes(3, 1);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public void Dispose()
        {
            //this.Dispose();
        }


        #endregion // private methods

    }
}
