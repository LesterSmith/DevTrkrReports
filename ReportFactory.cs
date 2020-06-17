using BusinessObjects;
namespace DevTrkrReports
{
    public static class ReportFactory 
    {
        public static Reporter Factory(ReporterParms parms)
        {
            switch (parms.Type)
            {
                case ReportType.ProjectSummaryByUser:
                    var pr = (Reporter) new ProjectReportByUser 
                    { 
                        FileName=parms.FileName,
                        Title=parms.Header.Title,
                        TitleCell=parms.Header.TitleCell,
                        ColHdrs=parms.Header.Hdrs,
                        HdrRange=parms.Header.HdrRange,
                        StartTime=parms.StartTime,
                        EndTime=parms.EndTime
                    };
                    return pr;
                case ReportType.DeveloperDetail:
                    var ur = (Reporter)new UserReport
                    {
                        FileName = parms.FileName,
                        Title = parms.Header.Title,
                        TitleCell = parms.Header.TitleCell,
                        ColHdrs = parms.Header.Hdrs,
                        HdrRange = parms.Header.HdrRange,
                        StartTime=parms.StartTime,
                        EndTime=parms.EndTime
                    };
                    return ur;
                case ReportType.ApplicationUsage:
                    var ar = (Reporter)new ApplicationReport
                    {
                        FileName = parms.FileName,
                        Title = parms.Header.Title,
                        TitleCell = parms.Header.TitleCell,
                        ColHdrs = parms.Header.Hdrs,
                        HdrRange = parms.Header.HdrRange,
                        StartTime = parms.StartTime,
                        EndTime = parms.EndTime
                    };
                    return ar;
                case ReportType.ProjectSummaryByProject:
                    var prp = (Reporter)new ProjectReportByProject
                    {
                        FileName = parms.FileName,
                        Title = parms.Header.Title,
                        TitleCell = parms.Header.TitleCell,
                        ColHdrs = parms.Header.Hdrs,
                        HdrRange = parms.Header.HdrRange,
                        StartTime = parms.StartTime,
                        EndTime = parms.EndTime
                    };
                    return prp;
                case ReportType.ProjectDetail:
                    var pd = (Reporter)new ProjectDetail
                    {
                        FileName = parms.FileName,
                        Title = parms.Header.Title,
                        TitleCell = parms.Header.TitleCell,
                        ColHdrs = parms.Header.Hdrs,
                        HdrRange = parms.Header.HdrRange,
                        StartTime = parms.StartTime,
                        EndTime = parms.EndTime
                    };
                    return pd;
                default:
                    return null;
            }
        }
    }
}
