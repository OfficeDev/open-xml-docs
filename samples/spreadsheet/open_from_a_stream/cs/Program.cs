using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;

FileStream fileStream = new(args[0], FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
OpenAndAddToSpreadsheetStream(fileStream);

static void OpenAndAddToSpreadsheetStream(Stream stream)
{
    // Open a SpreadsheetDocument based on a stream.
    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, true);

    if (spreadsheetDocument is not null)
    {
        // Get or create the WorkbookPart
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();

        // Add a new worksheet.
        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        newWorksheetPart.Worksheet.Save();

        Workbook workbook = workbookPart.Workbook ?? new Workbook();

        if (workbookPart.Workbook is null)
        {
            workbookPart.Workbook = workbook;
        }

        Sheets sheets = workbook.GetFirstChild<Sheets>() ?? workbook.AppendChild(new Sheets());
        string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        // Get a unique ID for the new worksheet.
        uint sheetId = 1;

        if (sheets.Elements<Sheet>().Count() > 0)
        {
            sheetId = (sheets.Elements<Sheet>().Select(s => s.SheetId?.Value).Max() + 1) ?? (uint)sheets.Elements<Sheet>().Count() + 1;
        }

        // Give the new worksheet a name.
        string sheetName = "Sheet" + sheetId;

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);
        workbookPart.Workbook.Save();

        // Dispose the document handle.
        spreadsheetDocument.Dispose();
    }
}