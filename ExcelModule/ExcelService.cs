using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
 

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console9_Excel
{
    internal class ExcelService
    {

        public byte[] Write( List<List<List<string>>> sheets )
        {
           
            using (var memoryStream = new MemoryStream())
            {
                var spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);

                // Add a WorkbookPart to the document.
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any()
                    ? spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                    : spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();

                foreach (var sheet in sheets)
                {
                    // Insert a new worksheet.
                    var worksheetPart = ExcelExportUtils.InsertWorksheet(spreadsheetDocument.WorkbookPart, "");

                    uint rowIndex = 1;
                    uint columnIndex = 1;

                    foreach (var row in sheet)
                    {
                        columnIndex = 1;
                        foreach (object item in row)
                        {
                            int index = ExcelExportUtils.InsertSharedStringItem(item.ToString(), shareStringPart);

                             
                            Cell cell = ExcelExportUtils.InsertCellInWorksheet(columnIndex, rowIndex, worksheetPart);

                            cell.CellValue = new CellValue(index.ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                            columnIndex++;
                        }
                        rowIndex++;
                    }
                }
                workbookPart.Workbook.Save();
                spreadsheetDocument.Save();
                spreadsheetDocument.Close();
                return memoryStream.ToArray();   
            }                       
        }


        public List<List<string>> Read( byte[] data )
        {
            var result = new List<List<string>>();
            using var stream = new MemoryStream(data);
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                Worksheet sheet = worksheetPart?.Worksheet;

                var rows = sheet?.Descendants<Row>();
                if (rows is not null)
                {
                    foreach (Row row in rows)
                    {
                        var inlineResult = new List<string>();
                        foreach (var cell in row.Elements<Cell>())
                            inlineResult.Add(cell.InnerText.Trim());
                        result.Add(inlineResult);
                    }
                }
            }
            return result;
        }
    }

    public class ExcelServiceUtils
    {
        /// <summary>
        /// Given a WorkbookPart, inserts a new worksheet.
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static WorksheetPart InsertWorksheet(WorkbookPart workbookPart, string sheetName)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets is null)
            {
                sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            }

            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        public static string GetExcelColumnName(uint columnNumber)
        {
            int dividend = (int)columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        ///  Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        /// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="shareStringPart"></param>
        /// <returns></returns>
        public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        /// <summary>
        /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        /// If the cell already exists, returns it. 
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="worksheetPart"></param>
        /// <returns></returns>
        public static Cell InsertCellInWorksheet(uint columnIndex, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            var columnName = GetExcelColumnName(columnIndex);
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Any(r => r.RowIndex == rowIndex))
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Any())
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).FirstOrDefault();
            }
            else
            {
                Cell newCell = new Cell() { CellReference = cellReference };
                row.Append(newCell);

                worksheet.Save();
                return newCell;
            }
        }
    }
}
