// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
// <Snippet2>

// <Snippet1>
Sheets? sheets = GetAllWorksheets(args[0]);
// </Snippet1>

if (sheets is not null)
{
    foreach (Sheet sheet in sheets)
    {
        Console.WriteLine(sheet.Name);
    }
}
// </Snippet2>

// Retrieve a List of all the sheets in a workbook.
// The Sheets class contains a collection of 
// OpenXmlElement objects, each representing one of 
// the sheets.
static Sheets? GetAllWorksheets(string fileName)
{
    // <Snippet3>
    Sheets? theSheets = null;
    // </Snippet3>
    // <Snippet4>
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // <Snippet5>
        theSheets = document?.WorkbookPart?.Workbook.Sheets;
        // </Snippet5>
        // </Snippet4>
    }

    return theSheets;
}
// </Snippet0>
