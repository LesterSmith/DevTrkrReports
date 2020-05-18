using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
