using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Console9_Excel
{
    internal class ExcelProgram
    {

        static void Main(string[] args)
        {
            var document = System.IO.File.ReadAllBytes(@"D:\System-Config\DataStore\books.xlsx");
            var excel = new ExcelService();
            Console.WriteLine(JsonConvert.SerializeObject(excel.Read(document),Formatting.Indented));

            byte[] data = excel.Write(new List<List<List<string>>>() { excel.Read(document) });
            System.IO.File.WriteAllBytes("book.xlsx", data);
            /**
             * 
             *  SampleFormula.Run();
             var bytes = System.IO.File.ReadAllBytes(@"C:\ftp\NetProjects\Mbd_Feature_Excel\Console9_Excel\KeyVallu.xls");
             var excel = new Program();
             excel.Read(bytes); */
            /*
            var data = excel.Write(new List<object>()
            {
                new { k=1, v=1 },
                new { k=2, v=2 },
                new { k=3, v=3 },                
            });
            Console.WriteLine(data.Length+" байт");*/
        }


        private void Read(byte[] data)
        {
            var datasource = new List<List<object>>();
            using( var reader = new MemoryStream(data))
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(reader, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;

                int rowCount = sheetData.Elements<Row>().Count();

                foreach (Row r in sheetData.Elements<Row>())
                {
                    var row = new List<object>();
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c.CellValue.Text;
                        Console.WriteLine(c.CellFormula);
                        row.Add(c);
                        Console.Write(text + " ");
                    }
                    datasource.Add(row);
                }
                Console.WriteLine("");

            }
           


        }






        public byte[] Write(  List<object> resultset)
        {
            using (var stream = new MemoryStream())
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create("1.xls", SpreadsheetDocumentType.Workbook))
                {
                    WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet/*, "Commercial"*/);
                    if (worksheetPart != null)
                    {
                        Worksheet worksheet = new Worksheet();
                        worksheetPart.Worksheet = worksheet;
                        SheetData sheetData = new SheetData();
                        UInt32 rowIndex = 1;
                        foreach (var item in resultset)
                        {
                            Row row = new Row() { RowIndex = rowIndex };
                            item.GetType().GetProperties().ToList().ForEach((property) =>
                            {
                                object value = property.GetValue(item);

                                Cell cell = new Cell()
                                {
                                    CellReference = property.Name + rowIndex,
                                    DataType = CellValues.Number,
                                    CellValue = new CellValue(value.ToString())
                                };

                                row.Append(cell);
                            });
                            rowIndex = rowIndex + 1;
                            sheetData.Append(row);
                        }
                        worksheet.Append(sheetData);
                        worksheetPart.Worksheet.Save();
                    }
                    spreadSheet.WorkbookPart.Workbook.Save();
                    Console.WriteLine(stream.Length);

                }
                return stream.ToArray();
            }
          
        }

        private WorksheetPart GetWorksheetPartByName(SpreadsheetDocument spreadSheet )
        {
            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadSheet.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            return worksheetPart;
        }
    }
}
