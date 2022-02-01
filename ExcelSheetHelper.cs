using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Export2ExlEngine
{
    public enum CellFillType
    {
        AbsoluteRowInSheet,
        AbsoluteColumnInSheet,        
        LastFilledRow,
        LastFilledColumn,
        LastFilledRowInColumn,        
    }
    public enum SaveWorkBookType
    {
        SaveThis,
        SaveNew,
        SaveAndReopen,
        SaveNewAndReopen
    }

    public class ExcelSheetHelper:IDisposable
    {
        static Microsoft.Office.Interop.Excel.Application xl = null;
        static int iInstancesOpened = 0;
        private Workbook thisWorkBook;
        private string sWorkBookFile = "";
        bool _bReadOnly = false, bIsDebug = false, bDontCloseApp = false;

        #region Technical
        public ExcelSheetHelper(bool IsDebug = false)
        {
            bIsDebug = IsDebug;
            OpenXlApp();
            iInstancesOpened++;
        }

        public ExcelSheetHelper(Microsoft.Office.Interop.Excel.Application xlApp, Workbook wb)
        {
            if (xlApp == null || wb == null)
                throw new Exception("");

            xl = xlApp;
            thisWorkBook = wb;
            bDontCloseApp = true;
        }

        public ExcelSheetHelper(string sFileName, bool bReadOnly = false):this()
        {
            sWorkBookFile = sFileName;

            thisWorkBook = OpenWorkBook(sFileName, bReadOnly);
        }

        public Workbook OpenWorkBook(string sFileName = "", bool bReadOnly = false)
        {
            if (CheckIntegrigy())
            {
                thisWorkBook = xl.Workbooks.Open(Filename: sFileName, ReadOnly: bReadOnly);
                _bReadOnly = bReadOnly;
                return thisWorkBook;
            }
            else
                return null;
        }

        private bool CheckIntegrigy()
        {
            bool result = true;

            if (sWorkBookFile.Trim().Length == 0)
                throw new Exception("File path and name are empty!");
            if (sWorkBookFile.StartsWith("http://"))
            {
                WebRequest webRequest = WebRequest.Create(sWorkBookFile);
                webRequest.Timeout = 1200; // miliseconds
                webRequest.Method = "HEAD";
                try
                {
                    webRequest.GetResponse();
                }
                catch
                {
                    result = false;
                }
            }
            else
                result = File.Exists(sWorkBookFile);
            
            if(!result) throw new IOException($"File {sWorkBookFile} doesn't exist at the specified location");

            return true;
        }

        ExcelSheetHelper()
        {   Dispose();  }

        public void Dispose()
        {
            if(iInstancesOpened > 0 & CloseWorkBook())
                iInstancesOpened--;            

            if (!bDontCloseApp && iInstancesOpened == 0 && xl != null)
            {
                xl.Quit();
                xl = null;
                GC.Collect();
            }
        }

        private void OpenXlApp()
        {
            //System.Windows.Forms.Application.DoEvents();

            if (xl == null)
            {
                xl = new Microsoft.Office.Interop.Excel.Application();
                xl.DisplayAlerts = false;
                xl.Visible = bIsDebug;
            }
        }

        private bool CloseWorkBook(bool bSaveChanges = false)
        {
            bool bClosing = false;

            if (thisWorkBook != null)
            {
                try
                {
                    xl.DisplayAlerts = false;
                    thisWorkBook.Close(SaveChanges: bSaveChanges);
                    bClosing = true;
                }
                catch (Exception ex)
                {                    
                }
                finally
                {
                    thisWorkBook = null;                    
                }                
            }
            return bClosing;
        }
        #endregion

        public string WorkBookName { get; set; }

        public Worksheet Sheet (string SheetName)
        {            
            try
            {
                foreach (Worksheet item in thisWorkBook.Sheets)
                {
                    if (item.Name.ToLower() == SheetName.ToLower())
                        return item;                        
                }                
            }
            catch (Exception ex)
            {                return null;            }
            return null;
        }

        public Worksheet Sheet(int SheetNo)
        {
            Worksheet worksheet = null;
            try
            {
                worksheet = thisWorkBook.Sheets[SheetNo];
            }
            catch (Exception ex)
            { return null; }
            return worksheet;
        }

        public int SheetsCount
        {
            get
            {
                if (thisWorkBook != null)
                    return thisWorkBook.Sheets.Count;
                else
                    return 0;
            }
        }

        public XlFileFormat FileFormat { get { return thisWorkBook.FileFormat; } }
        public string Path
        {            get { return thisWorkBook.Path; }        }
        public string Name
        { get { return thisWorkBook.Name; } }
        public void Save()
        { thisWorkBook.Save(); }

        public int GetSheetLastRowOrColumn(string SheetName, CellFillType fillType, int ColumnNo = 0)
        {
            int iLastRow = 1;
            _GetSheetLastRowOrColumn(Sheet(SheetName), fillType);
            return iLastRow;
        }

        public int GetSheetLastRowOrColumn(int SheetNo, CellFillType fillType, int ColumnNo = 0)
        {
            int iValue = 1;
            iValue = _GetSheetLastRowOrColumn(Sheet(SheetNo), fillType, ColumnNo);
            return iValue;
        }

        private int _GetSheetLastRowOrColumn(Worksheet worksheet, CellFillType fillType, int ColumnNo = 0)
        {
            if (worksheet == null)
                throw new Exception("Object worksheet is emtpy");

            switch (fillType)
            {
                case CellFillType.AbsoluteRowInSheet:
                    return worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                case CellFillType.AbsoluteColumnInSheet:
                    return worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;                    
                case CellFillType.LastFilledRow:
                    int iLastRow = 0;
                    for (int iColumn = 1; iColumn <= worksheet.Columns.Count; iColumn++)
                    {
                        iLastRow = Math.Max(iLastRow, GetLastFilledRowInColumn(worksheet, iColumn));
                    }
                    return iLastRow;
                case CellFillType.LastFilledColumn:                    
                    return GetLastFilledColumn(worksheet);
                case CellFillType.LastFilledRowInColumn:
                    return GetLastFilledRowInColumn(worksheet, ColumnNo);
                default:
                    break;                    
            }
            return 0;
        }

        private int GetLastFilledColumn(Worksheet worksheet)
        {
            Range rng = null;
            int iCol = 0, iLastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row, iLastCol = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;            

            for (int ColumnNo = 1; ColumnNo <= iLastCol; ColumnNo++)
            {
                //if there's a value in the 1st cell - the column is not empty
                if(worksheet.Cells[1, ColumnNo].Value != null && worksheet.Cells[1,ColumnNo].Value.ToString().Trim() != "")
                    iCol = ColumnNo;
                else
                {
                    //otherwise - jump loop until last row or until found non-empty trimmed cell
                    bool bStop = false;
                    rng = worksheet.Cells[1, ColumnNo];                    
                    do
                    {
                        rng = rng.End[XlDirection.xlDown];
                        if (rng == null || (rng != null && (rng.Row == worksheet.Rows.Count || rng.Value == null || rng.Value.ToString().Trim() != "")))
                            bStop = true;

                    } while (!bStop);

                }                    
            }
            return iCol;
        }

        //slow 
        private int GetLastFilledColumnOld(Worksheet worksheet)
        {            
            Range rng = null;
            int iCol = 0, iLastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row, iLastCol = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

            for (int ColumnNo = 1; ColumnNo <= iLastCol; ColumnNo++)
            {
                rng = (worksheet.Range[worksheet.Cells[1, ColumnNo], worksheet.Cells[iLastRow, ColumnNo]] as Range);                
                string[] inputData = rng.Cells.Cast<Range>().Select(item=>item.ToString().Trim()).ToArray<string>();                
                string[] separator = new string[1];
                string[] tmpResult;
                separator[0] = Environment.NewLine;                
                tmpResult = inputData.AsEnumerable().Where(item => item.ToString().Trim() != "").ToArray();
                if (tmpResult.Length > 0)
                    iCol = ColumnNo;
            }
            return iCol;
        }       

        private int GetLastFilledRowInColumn(Worksheet worksheet, int ColumnNo)
        {
            Range rng;
            int iLastRow = 0;
            iLastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            rng = worksheet.Cells[iLastRow, ColumnNo]; //get to the last cell in specified row and then jump up to the next not empty
            if (rng != null && rng.Value != null && rng.Value.ToString().Trim().Length > 0)
                return rng.Row;
            else
            {
                rng = rng.End[XlDirection.xlUp];
                if (rng != null)
                    return rng.Row;
                else
                    return 0;
            }
        }

        public string GetColumnName(int SheetNo, int Column)
        {
            Worksheet sht = thisWorkBook.Sheets[SheetNo];
            return sht.Cells[1, Column].Value.ToString();
        }

        public int GetColumnIndexByName(int SheetNo, string ColumnName)
        {            
            Worksheet sht = thisWorkBook.Sheets[SheetNo];
            Range rg1stRow = (sht.Rows[1] as Range);
            Range col = rg1stRow.Find(ColumnName);
            if (col != null)
                return col.Column;
            else
                return 0;
        }

        /// <summary>
        /// Save current workbook with options
        /// </summary>
        /// <param name="saveType"></param>
        /// <param name="NewFileName">only for SaveNew and SaveNewAndReopen</param>
        /// <returns></returns>
        public Workbook SaveWorkBook (SaveWorkBookType saveType = SaveWorkBookType.SaveThis, string NewFileName = "")
        {
            Workbook wb = null;

            switch (saveType)
            {
                case SaveWorkBookType.SaveThis:
                    thisWorkBook.Save();
                    break;
                case SaveWorkBookType.SaveNew:
                    thisWorkBook.SaveCopyAs(NewFileName);
                    break;
                case SaveWorkBookType.SaveAndReopen:
                    thisWorkBook.Save();
                    thisWorkBook.Close();
                    wb = OpenWorkBook(bReadOnly: _bReadOnly);
                    break;
                case SaveWorkBookType.SaveNewAndReopen:
                    thisWorkBook.SaveCopyAs(NewFileName);
                    thisWorkBook.Close();
                    wb = OpenWorkBook(NewFileName, _bReadOnly);
                    break;
                default:
                    break;
            }
            return wb;
        }
    }
}
