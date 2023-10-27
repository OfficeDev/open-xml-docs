using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

if (args.Length >= 3)
{
    GetHiddenRowsOrCols(args[0], args[1], args[2]);
}
else
{
    GetHiddenRowsOrCols(args[0], args[1]);
}

static List<uint> GetHiddenRowsOrCols(string fileName, string sheetName, string detectRows = "false")
{
    // Given a workbook and a worksheet name, return 
    // either a list of hidden row numbers, or a list 
    // of hidden column numbers. If detectRows is true, return
    // hidden rows. If detectRows is false, return hidden columns. 
    // Rows and columns are numbered starting with 1.

    List<uint> itemList = new List<uint>();

    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        if (document is not null)
        {
            WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();

            Sheet? theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault((s) => s.Name == sheetName);

            if (theSheet is null)
            {
                throw new ArgumentException("sheetName");
            }
            else
            {
                string id = theSheet.Id?.ToString() ?? string.Empty;
                // The sheet does exist.
                WorksheetPart? wsPart = wbPart.GetPartById(id) as WorksheetPart;
                Worksheet? ws = wsPart?.Worksheet;

                if (ws is not null)
                {
                    if (detectRows.ToLower() == "true")
                    {
                        // Retrieve hidden rows.
                        itemList = ws.Descendants<Row>()
                            .Where((r) => r?.Hidden is not null && r.Hidden.Value)
                            .Select(r => r.RowIndex?.Value)
                            .Cast<uint>()
                            .ToList();
                    }
                    else
                    {
                        // Retrieve hidden columns.
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
                    }
                }
            }
        }
    }

    return itemList;
}