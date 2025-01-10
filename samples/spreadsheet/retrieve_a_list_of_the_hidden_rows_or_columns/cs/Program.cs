using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

List<uint>? items = null;

if (args is [{ } fileName, { } sheetName, { } detectRows])
{
    items = GetHiddenRowsOrCols(fileName, sheetName, detectRows);
}
else if (args is [{ } fileName2, { } sheetName2])
{
    items = GetHiddenRowsOrCols(fileName2, sheetName2);
}

if (items is null)
{
    throw new ArgumentException("Invalid arguments.");
}

foreach (uint item in items)
{
    Console.WriteLine(item);
}

// <Snippet0>
static List<uint> GetHiddenRowsOrCols(string fileName, string sheetName, string detectRows = "false")
{
    // Given a workbook and a worksheet name, return 
    // either a list of hidden row numbers, or a list 
    // of hidden column numbers. If detectRows is true, return
    // hidden rows. If detectRows is false, return hidden columns. 
    // Rows and columns are numbered starting with 1.
    List<uint> itemList = new List<uint>();

    // <Snippet1>
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        if (document is not null)
        {
            WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();
            // </Snippet1>

            // <Snippet2>
            Sheet? theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault((s) => s.Name == sheetName);
            // </Snippet2>

            if (theSheet is null || theSheet.Id is null)
            {
                throw new ArgumentException("sheetName");
            }
            else
            {

                // <Snippet3>
                // The sheet does exist.
                WorksheetPart? wsPart = wbPart.GetPartById(theSheet.Id!) as WorksheetPart;
                Worksheet? ws = wsPart?.Worksheet;
                // </Snippet3>

                if (ws is not null)
                {
                    if (detectRows.ToLower() == "true")
                    {
                        // <Snippet4>
                        // Retrieve hidden rows.
                        itemList = ws.Descendants<Row>()
                            .Where((r) => r?.Hidden is not null && r.Hidden.Value)
                            .Select(r => r.RowIndex?.Value)
                            .Cast<uint>()
                            .ToList();
                        // </Snippet4>
                    }
                    else
                    {
                        // Retrieve hidden columns.
                        // <Snippet5>
                        var cols = ws.Descendants<Column>().Where((c) => c?.Hidden is not null && c.Hidden.Value);

                        foreach (Column item in cols)
                        {
                            if (item.Min is not null && item.Max is not null)
                            {
                                for (uint i = item.Min.Value; i <= item.Max.Value; i++)
                                {
                                    itemList.Add(i);
                                }
                            }
                        }
                        // </Snippet5>
                    }
                }
            }
        }
    }

    return itemList;
}
// </Snippet0>
