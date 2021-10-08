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
                @"C:\Users\pford\Desktop\Developer\CreateSpreadSheet\CreateSpreadSheet\constructorTest.xlsx";
            CreateSreedsheetDocument(path);
            AddWorkBookToDocument();
            AddWorkSheetPartToWorkBookPart();
            CreateSpreadsheetWorkbook();
        }


        private void CreateSreedsheetDocument(string filepath)
        {
            _spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
        }

        private void AddWorkBookToDocument()
        {
            _workbookPart = _spreadsheetDocument.AddWorkbookPart();
            _workbookPart.Workbook = new Workbook();
        }

        private void AddWorkSheetPartToWorkBookPart()
        {
            _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            _worksheetPart.Worksheet = new Worksheet(new SheetData());
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
            var sheets = _spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            
            var sheet = new Sheet()
                {Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(_worksheetPart), SheetId = 1, Name = "mySheet"};
            sheets.Append(sheet);

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
    }
}