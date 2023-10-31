#nullable disable
using DocumentFormat.OpenXml.Packaging;

static void OpenSpreadsheetDocumentReadonly(string filepath)
{
    // Open a SpreadsheetDocument based on a filepath.
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
    {
        // Attempt to add a new WorksheetPart.
        // The call to AddNewPart generates an exception because the file is read-only.
        WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

        // The rest of the code will not be called.
    }
}