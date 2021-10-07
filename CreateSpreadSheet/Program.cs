using System;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;


namespace CreateSpreadSheet
{
    public static class Program
    {
        private static void Main()
        {
            //create spreadsheetdocument object
            const string path = @"C:\Users\paulf\RiderProjects\CreateSpreadSheet\CreateSpreadSheet\test.xlsx";
            var testSpreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            //WorkbookPART
           //
            var wbPart = testSpreadsheet.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            wbPart.Workbook.AppendChild(new Sheets());
           
            //worksheets for workbookpart
            var worksheetPart = wbPart.AddNewPart<WorksheetPart>();
            //Provie a new workshet for worksheetpart.
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheet = new Sheet
            {
                Id = wbPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "yeppers"
            };
            var sheet2 = new Sheet
            {
                Id = wbPart.GetIdOfPart(worksheetPart),
                SheetId = 2,
                Name = "TEST2"
            };
            
            //Add data to sheet
            // Create row object
            Row row = new Row();
            Row row2 = new Row();
            Column column = new Column();

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            sheetData.Append(row);
            sheetData.Append(row2);
            
            
            
            Cell cell = new Cell
            {
                //CellReference = "A1",
                CellValue = new CellValue("Hello World"),
                DataType = CellValues.String
            };
            
            //If you dont preference a cell when appending to a row it will just add to it
            
            
            Cell cell2 = new Cell
            {
                //CellReference = "B1",
                CellValue = new CellValue("CHECK"),
                DataType = CellValues.String
            };
            row.Append(cell);
            row2.Append(cell2);

            testSpreadsheet.WorkbookPart.Workbook.Sheets.Append(sheet, sheet2);
           
            wbPart.Workbook.Save();
            testSpreadsheet.Close();

            var a = new CreateSpreadSheet();
        }
        private static void CheckFile(string path)
        {
            Console.Write(File.Exists(path) ? "File exist" : "File does not exist");
        }
    }
}