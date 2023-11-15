
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

CreateSpreadsheetWorkbook(args[0]);

static void CreateSpreadsheetWorkbook(string filepath)
{
    // Create a spreadsheet document by supplying the filepath.
    // By default, AutoSave = true, Editable = true, and Type = xlsx.
    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

    // Add a WorkbookPart to the document.
    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    // Add a WorksheetPart to the WorkbookPart.
    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart.Worksheet = new Worksheet(new SheetData());

    // Add Sheets to the Workbook.
    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

    // Append a new worksheet and associate it with the workbook.
    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
    sheets.Append(sheet);

    workbookPart.Workbook.Save();

    // Dispose the document.
    spreadsheetDocument.Dispose();
}
