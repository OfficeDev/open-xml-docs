
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

DeleteTextFromCell(args[0], args[1], args[2], uint.Parse(args[3]));

// Given a document, a worksheet name, a column name, and a one-based row index,
// deletes the text from the cell at the specified column and row on the specified worksheet.
static void DeleteTextFromCell(string docName, string sheetName, string colName, uint rowIndex)
{
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
    {
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()?.Elements<Sheet>()?.Where(s => s.Name is not null && s.Name == sheetName);
        if (sheets is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return;
        }
        string? relationshipId = sheets.First()?.Id?.Value;

        if (relationshipId is null)
        {
            // The worksheet does not have a relationship ID.
            return;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(relationshipId);

        // Get the cell at the specified column and row.
        Cell? cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
        if (cell is null)
        {
            // The specified cell does not exist.
            return;
        }

        cell.Remove();
        worksheetPart.Worksheet.Save();
    }
}

// Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
static Cell? GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
{
    IEnumerable<Row>? rows = worksheet.GetFirstChild<SheetData>()?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex);
    if (rows is null || rows.Count() == 0)
    {
        // A cell does not exist at the specified row.
        return null;
    }

    IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference?.Value, columnName + rowIndex, true) == 0);

    if (cells.Count() == 0)
    {
        // A cell does not exist at the specified column, in the specified row.
        return null;
    }

    return cells.FirstOrDefault();
}

// Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
// reference the specified SharedStringItem and removes the item.
static void RemoveSharedStringItem(int shareStringId, SpreadsheetDocument document)
{
    bool remove = true;

    if (document.WorkbookPart is null)
    {
        return;
    }

    foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
    {
        var cells = part.Worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>();

        if (cells is null)
        {
            continue;
        }

        foreach (var cell in cells)
        {
            // Verify if other cells in the document reference the item.
            if (cell.DataType is not null &&
                cell.DataType.Value == CellValues.SharedString &&
                cell.CellValue?.Text == shareStringId.ToString())
            {
                // Other cells in the document still reference the item. Do not remove the item.
                remove = false;
                break;
            }
        }

        if (!remove)
        {
            break;
        }
    }

    // Other cells in the document do not reference the item. Remove the item.
    if (remove)
    {
        SharedStringTablePart? shareStringTablePart = document.WorkbookPart.SharedStringTablePart;

        if (shareStringTablePart is null)
        {
            return;
        }

        SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
        if (item is not null)
        {
            item.Remove();

            // Refresh all the shared string references.
            foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
            {
                var cells = part.Worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>();

                if (cells is null)
                {
                    continue;
                }

                foreach (var cell in cells)
                {
                    if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int itemIndex))
                    {
                        if (itemIndex > shareStringId)
                        {
                            cell.CellValue.Text = (itemIndex - 1).ToString();
                        }
                    }
                }
                part.Worksheet.Save();
            }

            document.WorkbookPart.SharedStringTablePart?.SharedStringTable.Save();
        }
    }
}
