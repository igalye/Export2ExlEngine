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
        ExcelSheetHelper helper;

        #region Excel
        Microsoft.Office.Interop.Excel.Application excelApp = null;
        Workbook workbook = null;
        Worksheet worksheet = null;
        System.Globalization.CultureInfo oldCI = null;
        //igal 26-8-20
        int iMaxColWidth = 80;

        //igal 27/1/22
        public int MaxColumnWidth
        {
            get { return iMaxColWidth; }
            set { iMaxColWidth = value; }
        }

        public bool SuppressFileIfEmpty { get; set; }

        public bool AppendToFile { get; set; }

        public bool SilentOpen { get; set; }

        public string XlFileName { get; set; }

        public Export2Exl(string sConnection) : base(sConnection)
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

            //igal 23/9/20
            helper = new ExcelSheetHelper(excelApp, workbook);

            bSuccess = true;


            return bSuccess;
        }

        private int ExportTableToExcel(ref System.Data.DataTable dt, int SheetNo)
        {
            if (dt == null) return 0;

            string sSheetName = dt.TableName;
            int m_exported = 0;
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
            if (dt == null) return 0;

            if (!IfProceedFileCreate(dt.Rows.Count))
                return 0;

            OpenWorkBook(1);
            m_exported = ExportTableToExcel(ref dt, 1);

            return m_exported;
        }

        public int ExportToExcel(DataSet ds, bool bAutoFit = true)
        {
            if (ds == null) return 0;

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
                workbook.Sheets[1].Select();
                workbook.SaveCopyAs(XlFileName);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Error saving file [{0}]\nError:\n{1}", XlFileName, ex.Message), ex.InnerException);
            }
            finally
            { excelApp.DisplayAlerts = bAlert; }

            return true;
        }

        public bool CheckFileAndFolderPermissions(bool RemoveOldFile = false)
        {
            if (XlFileName == null || XlFileName.Trim().Length == 0)
            {
                throw new Exception("Please specify file name");
            }

            bool bOk = true;

            int index = Path.GetFileName(XlFileName).IndexOfAny(Path.GetInvalidFileNameChars());
            if (Path.GetFileName(XlFileName).Length > 0 && index != -1)
            {
                throw new Exception("CheckFileAndFolderPermissions: Illeagal character(s) in file name " + XlFileName + Environment.NewLine + Path.GetFileName(XlFileName) + Environment.NewLine + "index wrong=" + index.ToString());
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
                throw new Exception(string.Format("Error deleting file {0}.\nProbably is opened", XlFileName));
            }

            return bOk;
        }

        private void AutoFitSheets()
        {            
            for (int i = 1; i <= helper.SheetsCount; i++)
            {
                workbook.Sheets[i].Columns.EntireColumn.AutoFit();
                int iColCount = helper.GetSheetLastRowOrColumn(i, CellFillType.LastFilledColumn);
                for (int iCol = 1; iCol < iColCount; iCol++)
                {
                    Range col = workbook.Sheets[i].Columns[iCol];
                    double dWidth = col.ColumnWidth;

                    if (dWidth > iMaxColWidth)                    
                        col.ColumnWidth = (double)iMaxColWidth;                    
                }
            }
        }

        Export2Exl()
        {
            Dispose(false);
        }

        //igal 13-7-20 - in case there's a date based fields - there's a possibility to define them as date or datetime
        private IgalDAL.SqlFields _formatFields = new SqlFields();
        public IgalDAL.SqlFields formatFields { get { return _formatFields; } }

        private static string CreateTempTableFromDataTable(ref SqlConnection con, System.Data.DataTable dt, IgalDAL.SqlFields formatFields = null)
        {
            try
            {
                Random rnd = new Random();
                string tbl = "##tbl_" + rnd.Next(10000).ToString();

                int iTemp = SqlDAC.ExecuteNonQuery(con, CommandType.Text, Tools.TempTableFromDataTableCmdText(dt, tbl, formatFields), null);

                return tbl;
            }
            catch (Exception ex)
            {
                throw new Exception("Error creating temporary table", ex.InnerException);
            }

        }

        private int ExportToExcel(ref Worksheet worksheet, System.Data.DataTable dt)
        {
            //igal 2021-3-2
            bool bIsXl2003 = (Path.GetExtension(XlFileName) == "xls");
            int iMaxRows = ((bIsXl2003) ? UInt16.MaxValue: 1048575);
            int iRowsInData = dt.Rows.Count;
            if (iRowsInData > iMaxRows)
                throw new Exception($"Rows count in table {dt.TableName} is {iRowsInData} - more than possible in excel sheet - {iMaxRows}");

            string sCon = this.ConnectionString;
            if(string.IsNullOrWhiteSpace(sCon))
                throw new ArgumentNullException(sCon);

            SqlConnection con = new SqlConnection(sCon);

            string tbl = "";

            try
            {
                con.Open();

                CheckDateColumn(dt);

                tbl = CreateTempTableFromDataTable(ref con, dt, _formatFields);

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sCon))
                {
                    bulkCopy.BulkCopyTimeout = 0; //19-10-22 - once a datatable is here - don't limit timeout
                    bulkCopy.BatchSize = 25000;
                    bulkCopy.DestinationTableName = tbl;
                    DateTime start = DateTime.Now;
                    bulkCopy.WriteToServer(dt);
                    DateTime end = DateTime.Now;
                    bulkCopy.Close();
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"Error copy data to temp table\n{ex.Message}", ex);
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
                qry.CommandText = $"Select {Tools.GetFieldsSelect(dt, formatFields)} from {tbl}";
                qry.FillAdjacentFormulas = false;
                qry.PreserveFormatting = true;
                qry.RefreshOnFileOpen = false;
                qry.BackgroundQuery = true;
                qry.SavePassword = false;
                qry.SaveData = true;
                qry.AdjustColumnWidth = true;
                qry.RefreshPeriod = 0;
                qry.PreserveColumnInfo = true;
                qry.ListObject.Name = $"qry{worksheet.ListObjects.Count.ToString()}";
                qry.Refresh(BackgroundQuery: false);
            }
            catch (Exception ex)
            {
                new Exception($"Error in retrieving data from excel\n{ex.Message}", ex.InnerException);
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

        private void CheckDateColumn(System.Data.DataTable dt)
        {
            EnumerableRowCollection<DataRow> dr;
            SqlFieldType sqlField = new SqlFieldType();
            sqlField.sqlFieldType = "";

            try
            {
                foreach (DataColumn col in dt.Columns)
                {
                    if (col.DataType.Name == "DateTime")
                    {
                        dr = dt.AsEnumerable().Where(item => item[col.Ordinal].ToString() != "" && DateTime.Parse(item[col.Ordinal].ToString()).ToLongTimeString() == "12:00:00 AM");
                        if (dr.AsDataView().Count == dt.Rows.Count)
                            _formatFields.Add(col.ColumnName, new SqlFieldType { sqlFieldType = "date" }, "");
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
                if (SilentOpen)
                {
                    if (oldCI != null)
                        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;

                    if (worksheet != null)
                        worksheet = null;
                    if (workbook != null)
                        workbook = null;
                    if (excelApp != null)
                        excelApp.Quit();
                    excelApp = null;

                    GC.Collect();
                }
            }

            if (helper != null)
                helper.Dispose();

            bDisposed = true;
        }
    }
}
