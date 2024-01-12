// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;


// Given a document name, a worksheet name, and a cell name, gets the column of the cell and returns
// the content of the first cell in that column.
static string? GetColumnHeading(string docName, string worksheetName, string cellName)
{
    // Open the document as read-only.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
    {
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);

        if (sheets is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return null;
        }

        string? id = sheets.First().Id;

        if (id is null)
        {
            // The worksheet does not have an ID.
            return null;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);

        // <Snippet3>
        // Get the column name for the specified cell.
        string columnName = GetColumnName(cellName);

        // Get the cells in the specified column and order them by row.
        IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0)
            .OrderBy(r => GetRowIndex(r.CellReference) ?? 0);
        // </Snippet3>

        // <Snippet4>
        if (cells.Count() == 0)
        {
            // The specified column does not exist.
            return null;
        }

        // Get the first cell in the column.
        Cell headCell = cells.First();
        // </Snippet4>

        // <Snippet5>
        // If the content of the first cell is stored as a shared string, get the text of the first cell
        // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
        if (headCell.DataType is not null && headCell.DataType.Value == CellValues.SharedString && int.TryParse(headCell.CellValue?.Text, out int index))
        {
            SharedStringTablePart shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringItem[] items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

            return items[index].InnerText;
        }
        else
        {
            return headCell.CellValue?.Text;
        }
        // </Snippet5>
    }
}
// Given a cell name, parses the specified cell to get the column name.
static string GetColumnName(string? cellName)
{
    if (cellName is null)
    {
        return string.Empty;
    }

    // <Snippet1>
    // Create a regular expression to match the column name portion of the cell name.
    Regex regex = new Regex("[A-Za-z]+");
    Match match = regex.Match(cellName);

    return match.Value;
    // </Snippet1>
}

// Given a cell name, parses the specified cell to get the row index.
static uint? GetRowIndex(string? cellName)
{
    if (cellName is null)
    {
        return null;
    }

    // <Snippet2>
    // Create a regular expression to match the row index portion the cell name.
    Regex regex = new Regex(@"\d+");
    Match match = regex.Match(cellName);

    return uint.Parse(match.Value);
    // </Snippet2>
}
// </Snippet0>

Console.WriteLine("Column heading: {0}", GetColumnHeading(args[0], args[1], args[2]));
