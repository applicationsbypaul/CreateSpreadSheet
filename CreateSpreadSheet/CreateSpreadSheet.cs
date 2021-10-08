using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ColorType = DocumentFormat.OpenXml.Office2010.Excel.ColorType;

namespace CreateSpreadSheet
{
    public class CreateSpreadSheet
    {
        private static SpreadsheetDocument _spreadsheetDocument;
        private static WorkbookPart _workbookPart;
        private static WorksheetPart _worksheetPart;


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

        private static void CreateSpreadsheetWorkbook()
        {
            // Add Sheets to the Workbook.
            Sheets sheets = _spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
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

            Cell cell = new Cell
            {
                //CellReference = "A1",
                CellValue = new CellValue("Hello World"),
                DataType = CellValues.String
            };
            
            Cell cell2 = new Cell
            {
                //CellReference = "B1",
                CellValue = new CellValue("CHECK"),
                DataType = CellValues.String
                
            };
            row.Append(cell);
            row.Append(cell2);
            
            _workbookPart.Workbook.Save();

            // Close the document.
            _spreadsheetDocument.Close();
        }
    }
}