using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CreateSpreadSheet
{
    public class CreateSpreadSheet
    {
        private static SpreadsheetDocument _spreadsheetDocument;
        private static WorkbookPart _workbookPart;
        private static WorksheetPart _worksheetPart;
        private static Sheets _sheets;
        private static Sheet _sheet;
        

        private static readonly string[] Header =
        {
            "Recall Title", "Store", "Department", "Item Code", "UPC", "Description", "Quantity", "Total Cost"
        };

        public CreateSpreadSheet()
        {
            const string path =
                @"C:\Users\pford\Desktop\Developer\CreateSpreadSheet\CreateSpreadSheet\ExcelTest\constructorTest.xlsx";
            CreateSpreadsheetDocument(path);
            AddWorkBookToDocument();
            AddWorkSheetPartToWorkBookPart();
            CreateSheets();
            CreateSpreadsheetWorkbook();
        }
        private static void CreateSpreadsheetDocument(string filepath)
        {
            _spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
        }

        private static void AddWorkBookToDocument()
        {
            _workbookPart = _spreadsheetDocument.AddWorkbookPart();
            _workbookPart.Workbook = new Workbook();
        }

        private static void AddWorkSheetPartToWorkBookPart()
        {
            _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            _worksheetPart.Worksheet = new Worksheet(new SheetData());
        }

        private static void CreateSheets()
        {
            _sheets = _spreadsheetDocument.WorkbookPart.Workbook.AppendChild((new Sheets()));
            _sheet = new Sheet()
            {
                Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(_worksheetPart),
                SheetId = 1,
                Name = "Constructor Sheet"
            };
            _sheets.Append(_sheet);
        }

        private static void CreateSpreadsheetWorkbook()
        {
            var row = new Row();
            var row2 = new Row();

            var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData != null)
            {
                sheetData.Append(row);
                sheetData.Append(row2);
            }

            CreateHeader(Header, row);

            _workbookPart.Workbook.Save();
            _spreadsheetDocument.Close();
        }

        private static void CreateHeader(IEnumerable<string> header, Row row)
        {
            foreach (var headerTitle in header)
            {
                var cell = new Cell
                {
                    CellValue = new CellValue(headerTitle),
                    DataType = CellValues.String
                };
                row.Append(cell);
            }
        }
    }
}