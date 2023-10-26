#nullable disable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;

static void OpenAndAddToSpreadsheetStream(Stream stream)
{
    // Open a SpreadsheetDocument based on a stream.
    SpreadsheetDocument spreadsheetDocument =
        SpreadsheetDocument.Open(stream, true);

    // Add a new worksheet.
    WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
    newWorksheetPart.Worksheet = new Worksheet(new SheetData());
    newWorksheetPart.Worksheet.Save();

    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
    string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

    // Get a unique ID for the new worksheet.
    uint sheetId = 1;
    if (sheets.Elements<Sheet>().Count() > 0)
    {
        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
    }

    // Give the new worksheet a name.
    string sheetName = "Sheet" + sheetId;

    // Append the new worksheet and associate it with the workbook.
    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
    sheets.Append(sheet);
    spreadsheetDocument.WorkbookPart.Workbook.Save();

    // Dispose the document handle.
    spreadsheetDocument.Dispose();
}