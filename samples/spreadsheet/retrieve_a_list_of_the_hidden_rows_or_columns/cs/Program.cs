#nullable disable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

static List<uint> GetHiddenRowsOrCols(
  string fileName, string sheetName, bool detectRows)
{
    // Given a workbook and a worksheet name, return 
    // either a list of hidden row numbers, or a list 
    // of hidden column numbers. If detectRows is true, return
    // hidden rows. If detectRows is false, return hidden columns. 
    // Rows and columns are numbered starting with 1.

    List<uint> itemList = new List<uint>();

    using (SpreadsheetDocument document =
        SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart wbPart = document.WorkbookPart;

        Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
            Where((s) => s.Name == sheetName).FirstOrDefault();
        if (theSheet == null)
        {
            throw new ArgumentException("sheetName");
        }
        else
        {
            // The sheet does exist.
            WorksheetPart wsPart =
                (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            Worksheet ws = wsPart.Worksheet;

            if (detectRows)
            {
                // Retrieve hidden rows.
                itemList = ws.Descendants<Row>().
                    Where((r) => r.Hidden != null && r.Hidden.Value).
                    Select(r => r.RowIndex.Value).ToList<uint>();
            }
            else
            {
                // Retrieve hidden columns.
                var cols = ws.Descendants<Column>().
                    Where((c) => c.Hidden != null && c.Hidden.Value);
                foreach (Column item in cols)
                {
                    for (uint i = item.Min.Value; i <= item.Max.Value; i++)
                    {
                        itemList.Add(i);
                    }
                }
            }
        }
    }
    return itemList;
}