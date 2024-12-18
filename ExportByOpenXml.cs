using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Export2ExlEngine
{
    public static class ExportByOpenXml
    {
        public static void ExportDataSet(DataSet ds, string destination)
        {
            SpreadsheetDocument workbook;
            if (File.Exists(destination))
                workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            else
                workbook = SpreadsheetDocument.Open(destination, true);
                                
            WorkbookPart workbookPart = workbook.AddWorkbookPart();
            workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();
            foreach (System.Data.DataTable table in ds.Tables)
            {
                WorksheetPart sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);
                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }
                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                    { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                sheets.Append(sheet);
                DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);
                foreach (System.Data.DataRow dsrow in table.Rows)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (String col in columns)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                        newRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(newRow);
                }
            }
            workbook.Dispose();
        }

        public static DataSet ReadExcelFile(string filePath)
        {
            // Create a DataSet to hold all the DataTables
            DataSet dataSet = new DataSet();

            // Use FileStream to open the Excel file in read-only mode
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Create an ExcelDataReader to read the Excel file
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Convert the ExcelDataReader to a DataSet
                    dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true // Treat the first row as the header
                        }
                    });
                }
            }

            // Set DataTable names to corresponding sheet names
            foreach (DataTable table in dataSet.Tables)
            {
                table.TableName = table.TableName; // Optional: Set DataTable name to sheet name
            }

            return dataSet;
        }
    }
}