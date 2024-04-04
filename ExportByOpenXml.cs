using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
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

        public static DataSet ReadExcelToDataTable(string filePath)

        {
            DataTable dt;
            DataSet ds = new DataSet();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                if (workbookPart.Workbook.Descendants<Sheet>().Count() == 0)
                {
                    throw new Exception("No sheet found in the Excel file.");
                }

                foreach (Sheet sheet in workbookPart.Workbook.Descendants<Sheet>())
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    if (sheetData == null)
                    {
                        throw new Exception("No sheet data found in the Excel file.");
                    }

                    dt = new DataTable();

                    Row headerRow = sheetData.Elements<Row>().FirstOrDefault();
                    if (headerRow == null)
                    {
                        throw new Exception("No header row found in the Excel file.");
                    }

                    foreach (Cell cell in headerRow.Elements<Cell>())
                    {
                        dt.Columns.Add(GetCellValue(spreadsheetDocument, cell));
                    }

                    dt.TableName = sheet.Name;

                    foreach (Row row in sheetData.Elements<Row>().Skip(1)) // Skip header row
                    {
                        DataRow newRow = dt.Rows.Add();
                        int columnIndex = 0;

                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            while (columnIndex < dt.Columns.Count)
                            {
                                newRow[columnIndex++] = GetCellValue(spreadsheetDocument, cell);
                                break;
                            }
                        }
                    }
                    ds.Tables.Add(dt);
                }
            }

            return ds;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart sharedStringPart = document.WorkbookPart.SharedStringTablePart;
            string cellValue = "";

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (int.TryParse(cell.InnerText, out int index))
                {
                    cellValue = sharedStringPart.SharedStringTable.ElementAt(index).InnerText;
                }
            }
            else
            {
                cellValue = cell.InnerText;
            }

            return cellValue;
        }
    }
}