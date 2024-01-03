// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using System;
using OpenXmlAttribute = DocumentFormat.OpenXml.OpenXmlAttribute;
using OpenXmlElement = DocumentFormat.OpenXml.OpenXmlElement;
using Sheets = DocumentFormat.OpenXml.Spreadsheet.Sheets;


static void GetSheetInfo(string fileName)
{
    // Open file as read-only.
    using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
    {
        // <Snippet1>
        Sheets? sheets = mySpreadsheet.WorkbookPart?.Workbook?.Sheets;
        // </Snippet1>

        if (sheets is not null)
        {
            // For each sheet, display the sheet information.
            // <Snippet2>
            foreach (OpenXmlElement sheet in sheets)
            {
                foreach (OpenXmlAttribute attr in sheet.GetAttributes())
                {
                    Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
                }
            }
            // </Snippet2>
        }
    }
}
// </Snippet0>

GetSheetInfo(args[0]);
