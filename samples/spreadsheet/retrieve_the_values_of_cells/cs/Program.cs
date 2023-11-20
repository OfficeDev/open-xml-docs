// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

Console.WriteLine(GetCellValue(args[0], args[1], args[2]));
// Retrieve the value of a cell, given a file name, sheet name, 
// and address name.
// <Snippet1>
static string GetCellValue(string fileName, string sheetName, string addressName)
// </Snippet1>
{
    // <Snippet2>
    string? value = null;
    // </Snippet2>
    // <Snippet3>
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        WorkbookPart? wbPart = document.WorkbookPart;
        // </Snippet3>
        // <Snippet4>
        // Find the sheet with the supplied name, and then use that 
        // Sheet object to retrieve a reference to the first worksheet.
        Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

        // Throw an exception if there is no sheet.
        if (theSheet is null || theSheet.Id is null)
        {
            throw new ArgumentException("sheetName");
        }
        // </Snippet4>
        // <Snippet5>
        // Retrieve a reference to the worksheet part.
        WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(theSheet.Id!);
        // </Snippet5>
        // <Snippet6>
        // Use its Worksheet property to get a reference to the cell 
        // whose address matches the address you supplied.
        Cell? theCell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == addressName).FirstOrDefault();
        // </Snippet6>
        // <Snippet7>
        // If the cell does not exist, return an empty string.
        if (theCell is null || theCell.InnerText.Length < 0)
        {
            return string.Empty;
        }
        value = theCell.InnerText;
        // </Snippet7>
        // <Snippet8>
        // If the cell represents an integer number, you are done. 
        // For dates, this code returns the serialized value that 
        // represents the date. The code handles strings and 
        // Booleans individually. For shared strings, the code 
        // looks up the corresponding value in the shared string 
        // table. For Booleans, the code converts the value into 
        // the words TRUE or FALSE.
        if (theCell.DataType is not null)
        {
            if (theCell.DataType.Value == CellValues.SharedString)
            {
                // </Snippet8>
                // <Snippet9>
                // For shared strings, look up the value in the
                // shared strings table.
                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                // </Snippet9>
                // <Snippet10>
                // If the shared string table is missing, something 
                // is wrong. Return the index that is in
                // the cell. Otherwise, look up the correct text in 
                // the table.
                if (stringTable is not null)
                {
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
                // </Snippet10>
            }
            else if (theCell.DataType.Value == CellValues.Boolean)
            {
                // <Snippet11>
                switch (value)
                {
                    case "0":
                        value = "FALSE";
                        break;
                    default:
                        value = "TRUE";
                        // </Snippet11>
                        break;
                }
            }
        }
    }

    return value;
}
// </Snippet0>
