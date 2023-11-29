// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

static List<Sheet> GetHiddenSheets(string fileName)
{
    List<Sheet> returnVal = new List<Sheet>();

    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // <Snippet1>
        WorkbookPart? wbPart = document.WorkbookPart;

        if (wbPart is not null)
        {
            var sheets = wbPart.Workbook.Descendants<Sheet>();
            // </Snippet1>

            // Look for sheets where there is a State attribute defined, 
            // where the State has a value,
            // and where the value is either Hidden or VeryHidden.

            // <Snippet2>
            var hiddenSheets = sheets.Where((item) => item.State is not null &&
                item.State.HasValue &&
                (item.State.Value == SheetStateValues.Hidden ||
                item.State.Value == SheetStateValues.VeryHidden));
            // </Snippet2>

            returnVal = hiddenSheets.ToList();
        }
    }

    return returnVal;
}
// </Snippet0>

var sheets = GetHiddenSheets(args[0]);

foreach (var sheet in sheets)
{
    Console.WriteLine(sheet.Name);
}