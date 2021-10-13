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

        private static readonly string[] Header =
        {
            "Recall Title", "Store", "Department", "Item Code", "UPC", "Description", "Quantity", "Total Cost"
        };

        public CreateSpreadSheet()
        {
            const string path =
                @"C:\Users\paulf\RiderProjects\CreateSpreadSheet\CreateSpreadSheet\constructorTest.xlsx";
            CreateSreedsheetDocument(path);
            AddWorkBookAndSheetsToDocument();
            AddWorkSheetPartToWorkBookPart();
            AddColumnsToWorkSheet();
            AddRows();
            CreateSpreadsheetWorkbook();
        }


        private void CreateSreedsheetDocument(string filepath)
        {
            _spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
        }

        private void AddWorkBookAndSheetsToDocument()
        {
            _workbookPart = _spreadsheetDocument.AddWorkbookPart();
            _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            _workbookPart.Workbook = new Workbook();

            var sheets = new Sheets();
            var workSheetPartID = _workbookPart.GetIdOfPart(_worksheetPart);
            var sheet = new Sheet()
                { Id = workSheetPartID, SheetId = 1, Name = "mySheet" };
            var sheet2 = new Sheet()
                { Id = workSheetPartID, SheetId = 2, Name = "multi-sheet" };
            sheets.Append(sheet, sheet2);

            _workbookPart.Workbook.Append(sheets);
        }

        private void AddWorkSheetPartToWorkBookPart()
        {
            _worksheetPart.Worksheet = new Worksheet();
        }

        private void AddColumnsToWorkSheet()
        {
            var columns = new Columns();
            var column = new Column()
            {
                Min = 1,
                Max = 1,
                Width = 50,
                CustomWidth = true
            };
            var column2 = new Column()
            {
                Min = 2,
                Max = 2,
                Width = 25,
                CustomWidth = true
            };

            columns.Append(column, column2);
            _worksheetPart.Worksheet.Append(columns);
        }

        private void AddRows()
        {
            var row = new Row();
            var row2 = new Row();
            CreateHeader(Header, row);
            var sheetData = new SheetData();
            sheetData.Append(row, row2);
            _worksheetPart.Worksheet.Append(sheetData);
        }
        
        private static void CreateHeader(string[] header, Row row)
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

        private static void CreateSpreadsheetWorkbook()
        {
            _workbookPart.Workbook.Save();
            _spreadsheetDocument.Close();
        }
    }
}