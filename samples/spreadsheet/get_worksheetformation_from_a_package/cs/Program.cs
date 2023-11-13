
using DocumentFormat.OpenXml.Packaging;
using System;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;

GetSheetInfo(args[0]);

static void GetSheetInfo(string fileName)
{
    // Open file as read-only.
    using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
    {
        S? sheets = mySpreadsheet.WorkbookPart?.Workbook?.Sheets;

        if (sheets is not null)
        {
            // For each sheet, display the sheet information.
            foreach (E sheet in sheets)
            {
                foreach (A attr in sheet.GetAttributes())
                {
                    Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
                }
            }
        }
    }
}
