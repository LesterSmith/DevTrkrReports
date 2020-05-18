using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevTrkrReports
{
    class GenericReportHelper
    {
        //private void WriteGenericSheet(string custId, string template = "")
        //{
        //    try
        //    {
        //        if (ptr == 11)
        //            Debug.WriteLine(ptr);

        //        var desc = Descriptors[ptr];

        //        var ws = wb.Worksheets[desc.SSNbr];
        //        ws.Name = desc.SSName;

        //        var hlpr = new DataHelpers.DHCifAcctSummary(ApplicationWrapper.OldGisConnString);
        //        DataSet ds;
        //        if (!desc.ResultTypeDataTable)
        //            ds = hlpr.GetDataSetForGenericSheetWriter(custId, $"ImplementAdmin..{desc.SProc}", desc.TemplateNeeded ? template : string.Empty);
        //        else
        //        {
        //            // here the sproc name is used as as string pointer into the GenericMethodSQL class
        //            // to use dyn sql until the sproc is in production
        //            var hlpr2 = new DataHelpers.DHMisc(ApplicationWrapper.OldGisConnString);
        //            var spg = new GenericMethodSQL();

        //            ds = hlpr2.GetDataSetFromSQL(custId, (string)ApplicationWrapper.GetPropValue(spg, desc.SProc), false);
        //        }

        //        var dt = ds.Tables[desc.DSTableNbr];

        //        // if we are just starting this sheet put the column headers in it
        //        if (desc.RowId.Equals(1))
        //        {
        //            string[] columnNames = (from dc in dt.Columns.Cast<DataColumn>()
        //                                    select dc.ColumnName).ToArray();

        //            for (var k = 0; k < columnNames.Length; k++)
        //            {
        //                ws.Cells[0, k].PutValue(columnNames[k]);
        //            }
        //        }

        //        if (ds.Tables[desc.DSTableNbr].Rows.Count < 1)
        //        {
        //            wb.Save(SSFileName);
        //            return;
        //        }

        //        // set format for the header column
        //        CreateHdrBoldStyle(ref wb, ref ws, desc.ColumnFormats[0].StartColumn, desc.ColumnFormats[0].EndColumn, desc.ColumnFormats[0].Format);

        //        for (var i = 1; i <= desc.ColumnFormats.Count - 1; i++)
        //            CreateNumberedStyle(ref wb, ref ws, string.Format(desc.ColumnFormats[i].StartColumn, desc.RowId), string.Format(desc.ColumnFormats[i].EndColumn, desc.RowId + dt.Rows.Count), desc.ColumnFormats[i].Format);

        //        // populate the rows and columns for this set of datarows
        //        for (var j = 0; j < dt.Rows.Count; j++)
        //        {
        //            var row = dt.Rows[j];

        //            // loop thru the columns of the datarow
        //            // since the column headers are already in the sheet, put the data into row 
        //            for (var i = 0; i < dt.Columns.Count; i++)
        //            {
        //                ws.Cells[desc.RowId, i].PutValue(row[i] == DBNull.Value ? string.Empty : row[i]);
        //            }
        //            desc.RowId++;
        //        }
        //        //Debug.Print("End   GenMethod - Custid: " + custId + " Sheet: " + desc.SSName + " Rowid: " + desc.RowId);
        //        //wb.Save(tempFile);
        //        //Descriptors[ptr] = desc;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw;
        //    }
        //}
    }
}
