using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

GetAllWorksheets(args[0]);

// Retrieve a List of all the sheets in a workbook.
// The Sheets class contains a collection of 
// OpenXmlElement objects, each representing one of 
// the sheets.
static Sheets? GetAllWorksheets(string fileName)
{
    Sheets? theSheets = null;

    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        theSheets = document?.WorkbookPart?.Workbook.Sheets;
    }

    return theSheets;
}