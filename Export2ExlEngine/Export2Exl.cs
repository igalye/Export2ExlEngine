using System;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using IgalDAL;
using System.Data;
using System.IO;

namespace Export2ExlEngine
{
    public class Export2Exl : clsBaseConnection, IDisposable
    {
        bool bDisposed = false;

        #region Excel
        Microsoft.Office.Interop.Excel.Application excelApp = null;
        Workbook workbook = null;
        Worksheet worksheet = null;
        System.Globalization.CultureInfo oldCI = null;        

        public bool SuppressFileIfEmpty { get; set; }

        public bool AppendToFile { get; set; }

        public bool SilentOpen { get; set; }

        public string XlFileName { get; set; }

        public Export2Exl(string sConnection):base(sConnection)
        {
            SilentOpen = false;
            SuppressFileIfEmpty = false;
        }

        private bool OpenWorkBook(int SheetCount)
        {
            bool bSuccess = false;

            excelApp = new Application();
            excelApp.DisplayAlerts = !SilentOpen;
            excelApp.Visible = !SilentOpen;            

            Worksheet lastSheet;
            //igal 19-6-19
            if (File.Exists(XlFileName) && AppendToFile)
            {
                workbook = excelApp.Workbooks.Open(XlFileName);
            }
            else
            {
                workbook = excelApp.Workbooks.Add();
            }

            oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            for (int i = 1; i <= SheetCount; i++)
            {
                if (i > workbook.Worksheets.Count)
                {
                    lastSheet = workbook.Sheets[workbook.Sheets.Count] as Worksheet;
                    worksheet = workbook.Worksheets.Add(After: lastSheet) as Worksheet;
                }
                else
                {
                    worksheet = workbook.Worksheets[i] as Worksheet;
                }
            }

            bSuccess = true;


            return bSuccess;
        }

        private int ExportTableToExcel (ref System.Data.DataTable dt, int SheetNo)
        {
            string sSheetName = dt.TableName;
            int m_exported= 0;
            worksheet = workbook.Sheets[SheetNo];

            ////there's sometimes a problem quering from the sheet 
            //if (sSheetName.Contains("-"))
            //{
            //    sSheetName.Replace('-', '_');
            //    Console.WriteLine("There're *hypens* in the sheet name - replacing with *underscore*");
            //}            
            m_exported = ExportToExcel(ref worksheet, dt);
            worksheet.Name = sSheetName;

            return m_exported;
        }

        public int ExportToExcel(System.Data.DataTable dt, bool bAutoFit = true)
        {
            int m_exported = 0;

            if (!IfProceedFileCreate(dt.Rows.Count))
                return 0;

            OpenWorkBook(1);
            m_exported = ExportTableToExcel(ref dt, 1);            

            return m_exported;
        }

        public int ExportToExcel(DataSet ds, bool bAutoFit = true)
        {
            int m_exported = 0;

            //igal 22/7/19
            int iTotalRowsInDs = 0;
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                iTotalRowsInDs += ds.Tables[i].Rows.Count;
            }

            if (!IfProceedFileCreate(iTotalRowsInDs))
                return 0;

            OpenWorkBook(ds.Tables.Count);

            System.Data.DataTable dt;
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                dt = ds.Tables[i];
                m_exported += ExportTableToExcel(ref dt, i + 1);
            }

            if (bAutoFit)
            {
                AutoFitSheets();
            }            
                
            return m_exported;
        }

        private bool IfProceedFileCreate(int iTotalDataRows)
        {
            bool bProceed = ((SuppressFileIfEmpty & iTotalDataRows > 0) | !SuppressFileIfEmpty);
            return bProceed;
        }

        public bool SaveFile(string FileNameToSave = "")
        {

            if (FileNameToSave.Trim().Length != 0)
                XlFileName = FileNameToSave;


            if (workbook == null)
            {
                throw new Exception("Workbook is closed or doesn't exist");
            }

            CheckFileAndFolderPermissions();

            bool bAlert = excelApp.DisplayAlerts;
            excelApp.DisplayAlerts = !SilentOpen;
            try
            {
                workbook.SaveCopyAs(XlFileName);
            }
            catch (Exception ex)
            {
                throw new Exception (string.Format("Error saving file [{0}]\nError:\n{1}", XlFileName, ex.Message),ex.InnerException);
            }
            finally
            { excelApp.DisplayAlerts = bAlert; }

            return true;
        }

        public bool CheckFileAndFolderPermissions(bool RemoveOldFile = false)
        {
            bool bOk = true;

            int index = Path.GetFileName(XlFileName).IndexOfAny(Path.GetInvalidFileNameChars());
            if (Path.GetFileName(XlFileName).Length > 0 && index != -1)
            {
                throw new Exception("CheckFileAndFolderPermissions: Illeagal character(s) in file name " + XlFileName + Environment.NewLine + Path.GetFileName(XlFileName) + Environment.NewLine + "index wrong=" + index.ToString());
            }

            if (XlFileName == null || XlFileName.Trim().Length == 0)
            {
                throw new Exception("Please specify file name");
            }

            DirectoryInfo di = new System.IO.DirectoryInfo(Path.GetDirectoryName(XlFileName));
            if (!di.Exists)
                throw new Exception("Directory " + Path.GetFullPath(XlFileName) + " doesn't exits");

            if (!di.IsWriteable())
                throw new Exception("No writing permissions for directory " + di.Name);

            try
            {
                if (!AppendToFile && File.Exists(XlFileName))
                    File.Delete(XlFileName);
            }
            catch (Exception)
            {
                throw new Exception(string.Format("Error deleting file {0}.\nProbably is opened",XlFileName));
            }

            return bOk;
        }

        private void AutoFitSheets()
        {
            for (int i = 1; i <= workbook.Sheets.Count; i++)
			{
                workbook.Sheets[i].Columns.EntireColumn.AutoFit();
			}
        }

        Export2Exl()
        {
            Dispose(false);
        }

        private static string CreateTempTableFromDataTable(ref SqlConnection con, System.Data.DataTable dt)
        {
            Random rnd = new Random();
            string tbl = "##tbl_" + rnd.Next(10000).ToString();

            StringBuilder sbTempTable = new StringBuilder("CREATE TABLE " + tbl + "(");
            string sColDef = "";
            foreach (DataColumn col in dt.Columns)
            {
                switch (col.DataType.ToString())
                {
                    case "System.Int64":
                        sColDef = "[" + col.ColumnName + "] bigint ";
                        sColDef += (col.AutoIncrement) ? " Identity (" + col.AutoIncrementSeed.ToString() + "," + col.AutoIncrementStep.ToString() + ")," : ",";
                        sbTempTable.AppendLine(sColDef);
                        break;
                    case "System.Int32":
                        sColDef = "[" + col.ColumnName + "] int ";
                        sColDef += (col.AutoIncrement) ? " Identity (" + col.AutoIncrementSeed.ToString() + "," + col.AutoIncrementStep.ToString() + ")," : ",";
                        sbTempTable.AppendLine(sColDef);
                        break;
                    case "System.DateTime":
                        sbTempTable.AppendLine("[" + col.ColumnName + "] datetime2, ");
                        break;
                    case "System.String":
                        sColDef = "[" + col.ColumnName + "] varchar( ";
                        sColDef += (col.MaxLength == -1) ? "max" : col.MaxLength.ToString();
                        sColDef += "), ";
                        sbTempTable.AppendLine(sColDef);
                        break;
                    case "System.Single":
                        sbTempTable.AppendLine("[" + col.ColumnName + "] float , ");
                        break;
                    case "System.Double":
                        sbTempTable.AppendLine("[" + col.ColumnName + "] float , ");
                        break;
                    case "System.Int16":
                        sbTempTable.AppendLine("[" + col.ColumnName + "] smallint , ");
                        break;
                    case "System.Boolean":
                        sbTempTable.AppendLine("[" + col.ColumnName + "] bit , ");
                        break;
                    case "System.Decimal":
                        sbTempTable.AppendLine("[" + col.ColumnName + "] decimal(19,4) , ");
                        break;
                    case "System.Byte":
                        sbTempTable.AppendLine("[" + col.ColumnName + "] tinyint, ");
                        break;
                    case "System.Guid":
                        sbTempTable.AppendLine("[" + col.ColumnName + "] uniqueidentifier, ");
                        break;
                    default:
                        throw new Exception("No definition for column type " + col.DataType.ToString());
                }
            }

            string pks="";
            if (dt.PrimaryKey.Length > 0)
            {
                pks = "CONSTRAINT PK_" + tbl + " PRIMARY KEY (";
                for (int i = 0; i < dt.PrimaryKey.Length; i++)
                {
                    pks += dt.PrimaryKey[i].ColumnName + ",";
                }
                pks = pks.Substring(0, pks.Length - 1) + ")";
                
            }
            if (pks != "")
                sbTempTable.AppendLine(pks);
            else
                sbTempTable.Remove(sbTempTable.Length-1, 1);
            sbTempTable.Append(")");

            int iTemp = SqlDAC.ExecuteNonQuery(con, CommandType.Text, sbTempTable.ToString(), null);           

            return tbl;
        }

        private int ExportToExcel(ref Worksheet worksheet, System.Data.DataTable dt)
        {
            string sCon = this.ConnectionString;
            SqlConnection con = new SqlConnection(sCon);

            string tbl = "";

            try
            {
                con.Open();

                tbl = CreateTempTableFromDataTable(ref con, dt);

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sCon))
                {
                    bulkCopy.DestinationTableName = tbl;
                    //if (dt.Columns.Contains(destinationTable.Columns[i].ToString()))//contain method is not case sensitive
                    //{
                    //    //Once column matched get its index
                    //    int sourceColumnIndex = dt.Columns.IndexOf(destinationTable.Columns[i].ToString());
                    //    //give coluns name of source table rather then destination table so that it would avoid case sensitivity
                    //    bulkCopy.ColumnMappings.Add(dt.Columns[sourceColumnIndex].ToString(), dt.Columns[sourceColumnIndex].ToString());
                    //}
                    bulkCopy.WriteToServer(dt);
                    bulkCopy.Close();
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error creating temporary table\nex.Message", ex.InnerException);
            }

            Microsoft.Office.Interop.Excel.Range rng = null;
            Microsoft.Office.Interop.Excel.QueryTable qry = null;
            SqlConnectionStringBuilder cb = new SqlConnectionStringBuilder(this.ConnectionString);

            StringBuilder sConn = new StringBuilder();
            sConn.Append("OLEDB;Provider=SQLOLEDB.1;");
            sConn.Append("Data Source=" + con.DataSource + ";");
            sConn.Append("Initial Catalog=" + con.Database + ";");
            if (cb.UserID.Length > 0)
            {
                sConn.Append("User ID=" + cb.UserID + ";");
                sConn.Append("Password=" + cb.Password + ";");
            }                
            else
            {
                sConn.Append("Persist Security Info = False;Integrated Security=SSPI;");
            }

            try
            {
                if (AppendToFile)
                {
                    Range rngLast = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
                    rng = worksheet.Cells[rngLast.Row, 1];
                        }
                else
                    rng = worksheet.Range["$A$1"];

                qry = worksheet.ListObjects.AddEx(SourceType: Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcExternal, Source: sConn.ToString(), Destination: rng).QueryTable;
                qry.CommandText = "Select * from " + tbl;
                qry.FillAdjacentFormulas = false;
                qry.PreserveFormatting = true;
                qry.RefreshOnFileOpen = false;
                qry.BackgroundQuery = true;
                qry.SavePassword = false;
                qry.SaveData = true;
                qry.AdjustColumnWidth = true;
                qry.RefreshPeriod = 0;
                qry.PreserveColumnInfo = true;
                qry.ListObject.Name = "qry" + worksheet.ListObjects.Count.ToString();
                qry.Refresh(BackgroundQuery: false);
            }
            catch (Exception ex)
            {
                new Exception(string.Format("Error in retrieving data from excel\n{0}",ex.Message), ex.InnerException);
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
                if (con != null)
                    con.Dispose();
            }

            return qry.ResultRange.Rows.Count - 1; //ללא שורת כותרת
        }

        #endregion

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (bDisposed)
                return;

            if(disposing)
            {
                //excelApp.ActiveWindow.FreezePanes = true;                 
                if (oldCI != null)
                    System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;

                if (worksheet != null)
                    worksheet = null;
                if (workbook != null)
                    workbook = null;
                if (excelApp != null)
                {
                    if (SilentOpen)
                        excelApp.Quit();
                    excelApp = null;
                }

                GC.Collect();            
            }

            bDisposed = true;
        }
    }
}
