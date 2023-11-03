
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
    using E = DocumentFormat.OpenXml.OpenXmlElement;
    using A = DocumentFormat.OpenXml.OpenXmlAttribute;


    public static void GetSheetInfo(string fileName)
    {
        // Open file as read-only.
        using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
        {
            S sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;

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
