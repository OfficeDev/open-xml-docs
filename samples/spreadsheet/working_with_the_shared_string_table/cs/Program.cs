using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

using (var spreadsheetDoc = SpreadsheetDocument.Open(args[0], true))
{
    if (spreadsheetDoc is null) throw new ArgumentException("SpreadsheetDocument does not exist");

    WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart ?? spreadsheetDoc.AddWorkbookPart();
    SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart ?? 
        workbookPart.AddNewPart<SharedStringTablePart>();

    InsertSharedStringItem("test text", sharedStringTablePart);
}

// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
// <Snippet0>
static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
{
    // If the part does not contain a SharedStringTable, create one.
    if (shareStringPart.SharedStringTable is null)
    {
        shareStringPart.SharedStringTable = new SharedStringTable();
    }

    int i = 0;

    // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
    {
        if (item.InnerText == text)
        {
            return i;
        }

        i++;
    }

    // The text does not exist in the part. Create the SharedStringItem and return its index.
    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));

    return i;
}
// </Snippet0>