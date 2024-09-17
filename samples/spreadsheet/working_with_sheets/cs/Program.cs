using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

CreateSpreadsheetWorkbook(args[0]);

static void CreateSpreadsheetWorkbook(string filepath)
{
    // Create a spreadsheet document by supplying the filepath.
    // By default, AutoSave = true, Editable = true, and Type = xlsx.
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
    {

        // Add a WorkbookPart to the document.
        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        // Add a WorksheetPart to the WorkbookPart.
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        sheets.Append(sheet);

        // Get the sheetData cell table.
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>() ?? worksheetPart.Worksheet.AppendChild(new SheetData());

        // Add a row to the cell table.
        Row row;
        row = new Row() { RowIndex = 1 };
        sheetData.Append(row);

        // In the new row, find the column location to insert a cell in A1.  
        Cell? refCell = null;

        foreach (Cell cell in row.Elements<Cell>())
        {
            if (string.Compare(cell.CellReference?.Value, "A1", true) > 0)
            {
                refCell = cell;
                break;
            }
        }

        // Add the cell to the cell table at A1.
        Cell newCell = new Cell() { CellReference = "A1" };
        row.InsertBefore(newCell, refCell);

        // Set the cell value to be a numeric value of 100.
        newCell.CellValue = new CellValue("100");
        newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
    }
}
