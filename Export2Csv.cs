using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using IgalDAL;
using System.Data;
using System.IO;

namespace Export2ExlEngine
{
    public class Export2Csv : IDisposable
    {
        bool bDisposed = false;

        public Export2Csv()
        {        }

        public void Dispose()
        {            Dispose(true);        }

        protected virtual void Dispose(bool disposing)
        {
            if (bDisposed)
                return;

            if (disposing)
            {
                GC.Collect();
            }

            bDisposed = true;
        }

        public bool SuppressFileIfEmpty { get; set; }

        public string FileName { get; set; }
        public int ExportToCSV(System.Data.DataTable dt)
        {
            int m_exported = 0;

            try
            {                
                m_exported = ExportTableToCSV(ref dt);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return m_exported;
        }

        public int ExportToCSV(DataSet ds)
        {
            int m_exported = 0;

            System.Data.DataTable dt;
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                dt = ds.Tables[i];
                m_exported += ExportTableToCSV(ref dt, (ds.Tables.Count>1)?(i+1):0); //if there's multiple select - add a counter to a csv name
            }

            return m_exported;
        }

        private int ExportTableToCSV(ref System.Data.DataTable dt, int iTableNo = 0)
        {
            int m_exported = 0;

            StringBuilder sb = new StringBuilder();
            IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName);
            sb.AppendLine(string.Join(",", columnNames));
            foreach (DataRow row in dt.Rows)
            {                
                IEnumerable<string> fields = row.ItemArray.Select(field =>
                string.Concat("\"", field.ToString().Replace("\"", "\"\""), "\""));
                sb.AppendLine(string.Join(",", fields));
                m_exported++;
            }

            string sNameAddition = (iTableNo == 0)?"":"_" + iTableNo.ToString();
            FileInfo fi = new FileInfo(FileName);
            FileName = Path.GetFileNameWithoutExtension(fi.Name) + sNameAddition +  fi.Extension;
            File.WriteAllText(FileName, sb.ToString(), Encoding.UTF8);

            return m_exported;
        }
    }
}