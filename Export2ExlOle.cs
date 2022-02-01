using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace Export2ExlEngine
{    
    public class Export2ExlOle: IDisposable
    {
        OleDbConnection conn;
        public string FilePath { get; }
        public bool HasHeaders { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pathOfFileToCreate"></param>
        /// <param name="WithHeaders">specifying true will still leave 1st line blank</param>
        public Export2ExlOle(string pathOfFileToCreate, bool WithHeaders = true)
        {            
            HasHeaders = WithHeaders;
            FilePath = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES"";", pathOfFileToCreate);
            conn = new OleDbConnection();
        }

        public bool ExportToExcel(System.Data.DataTable dt, string sTableName = "sheet1")
        {
            if (dt == null || dt.Rows.Count == 0)
                return false;

            try
            {
                conn.ConnectionString = FilePath;
                conn.Open();
                var cmd = conn.CreateCommand();
                cmd.CommandText = Tools.TempTableFromDataTableCmdText(dt, sTableName); // Create Sheet With Name Sheet1
                cmd.ExecuteNonQuery();
                for (int i = 0; i < dt.Rows.Count; i++) // Sample Data Insert 
                {
                    cmd.CommandText = String.Format("INSERT INTO {0} VALUES({1},'{2}')", sTableName, i, "Name" + i.ToString());
                    cmd.ExecuteNonQuery(); // Execute insert query against excel file.
                }
                if (!HasHeaders)
                {
                    //delete header row but leave blank line
                    cmd.CommandText = "UPDATE [sheet1$] SET F1 = \"\", F2 = \"\"";
                    cmd.ExecuteNonQuery();
                }
            }
            finally
            {
                if (conn.State == System.Data.ConnectionState.Open)
                    conn.Close();
            }
            return true;
        }

        public void Dispose()
        {
            if (conn.State == System.Data.ConnectionState.Open)
                conn.Close();
            conn = null;
            GC.Collect();
        }
    }
}
