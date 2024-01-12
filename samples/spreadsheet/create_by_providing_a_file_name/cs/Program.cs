
// <Snippet0>
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

static void CreateSpreadsheetWorkbook(string filepath)
{
    // <Snippet1>
    // Create a spreadsheet document by supplying the filepath.
    // By default, AutoSave = true, Editable = true, and Type = xlsx.
    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
    // </Snippet1>

    // <Snippet2>
    // Add a WorkbookPart to the document.
    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();
    // </Snippet2>

    // <Snippet3>
    // Add a WorksheetPart to the WorkbookPart.
    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart.Worksheet = new Worksheet(new SheetData());

    // Add Sheets to the Workbook.
    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

    // Append a new worksheet and associate it with the workbook.
    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
    sheets.Append(sheet);
    // </Snippet3>

    workbookPart.Workbook.Save();

    // Dispose the document.
    spreadsheetDocument.Dispose();
}
// </Snippet0>

CreateSpreadsheetWorkbook(args[0]);
