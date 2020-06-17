using System;

namespace DevTrkrReports
{
    public class DevTrkrReports : IDisposable
    {
        #region class ..ctor
        public DevTrkrReports()
        {
        }

        public void Dispose()
        {
            
        }
        #endregion

        #region public methods
        public void RunForm()
        {
            var f = new frmReporter();
            f.ShowDialog();
            this.Dispose();
        }
        #endregion
    }
}
