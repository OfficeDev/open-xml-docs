using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

GetHiddenSheets(args[0]);

static List<Sheet> GetHiddenSheets(string fileName)
{
    List<Sheet> returnVal = new List<Sheet>();

    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart? wbPart = document.WorkbookPart;

        if (wbPart is not null)
        {
            var sheets = wbPart.Workbook.Descendants<Sheet>();

            // Look for sheets where there is a State attribute defined, 
            // where the State has a value,
            // and where the value is either Hidden or VeryHidden.
            var hiddenSheets = sheets.Where((item) => item.State is not null &&
                item.State.HasValue &&
                (item.State.Value == SheetStateValues.Hidden ||
                item.State.Value == SheetStateValues.VeryHidden));

            returnVal = hiddenSheets.ToList();
        }
    }

    return returnVal;
}