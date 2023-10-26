#nullable disable
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

static List<Sheet> GetHiddenSheets(string fileName)
{
    List<Sheet> returnVal = new List<Sheet>();

    using (SpreadsheetDocument document =
        SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart wbPart = document.WorkbookPart;
        var sheets = wbPart.Workbook.Descendants<Sheet>();

        // Look for sheets where there is a State attribute defined, 
        // where the State has a value,
        // and where the value is either Hidden or VeryHidden.
        var hiddenSheets = sheets.Where((item) => item.State != null &&
            item.State.HasValue &&
            (item.State.Value == SheetStateValues.Hidden ||
            item.State.Value == SheetStateValues.VeryHidden));

        returnVal = hiddenSheets.ToList();
    }
    return returnVal;
}