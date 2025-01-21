using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.IO.Packaging;

// <Snippet0>
static void OpenSpreadsheetDocumentReadonly(string filePath)
{
    // <Snippet1>
    // Open a SpreadsheetDocument based on a file path.
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
    // </Snippet1>
    {
        if (spreadsheetDocument.WorkbookPart is not null)
        {
            // Attempt to add a new WorksheetPart.
            // The call to AddNewPart generates an exception because the file is read-only.
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            // The rest of the code will not be called.
        }
    }

    // <Snippet2>
    // Open a SpreadsheetDocument based on a stream.
    Stream stream = File.Open(filePath, FileMode.Open);

    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
    // </Snippet2>
    {
        if (spreadsheetDocument.WorkbookPart is not null)
        {
            // Attempt to add a new WorksheetPart.
            // The call to AddNewPart generates an exception because the file is read-only.
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            // The rest of the code will not be called.
        }
    }

    // <Snippet3>
    // Open System.IO.Packaging.Package.
    Package spreadsheetPackage = Package.Open(filePath, FileMode.Open, FileAccess.Read);

    // Open a SpreadsheetDocument based on a package.
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage))
    // </Snippet3>
    {
        if (spreadsheetDocument.WorkbookPart is not null)
        {
            // Attempt to add a new WorksheetPart.
            // The call to AddNewPart generates an exception because the file is read-only.
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            // The rest of the code will not be called.
        }
    }
}
// </Snippet0>