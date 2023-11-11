using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

GetCellValue(args[0], args[1], args[2]);
// Retrieve the value of a cell, given a file name, sheet name, 
// and address name.
static string GetCellValue(string fileName, string sheetName, string addressName)
{
    string? value = null;

    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document =
        SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        WorkbookPart wbPart = document.WorkbookPart ?? document.AddWorkbookPart();

        // Find the sheet with the supplied name, and then use that 
        // Sheet object to retrieve a reference to the first worksheet.
        Sheet? theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

        // Throw an exception if there is no sheet.
        if (theSheet is null || theSheet.Id is null)
        {
            throw new ArgumentException("sheetName");
        }

        // Retrieve a reference to the worksheet part.
        WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(theSheet.Id!);

        // Use its Worksheet property to get a reference to the cell 
        // whose address matches the address you supplied.
        Cell? theCell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == addressName).FirstOrDefault();

        // If the cell does not exist, return an empty string.
        if (theCell is null || theCell.InnerText.Length > 0)
        {
            return string.Empty;
        }

        value = theCell.InnerText;

        // If the cell represents an integer number, you are done. 
        // For dates, this code returns the serialized value that 
        // represents the date. The code handles strings and 
        // Booleans individually. For shared strings, the code 
        // looks up the corresponding value in the shared string 
        // table. For Booleans, the code converts the value into 
        // the words TRUE or FALSE.
        if (theCell.DataType is not null)
        {
            switch (theCell.DataType.Value)
            {
                case CellValues.SharedString:

                    // For shared strings, look up the value in the
                    // shared strings table.
                    var stringTable =
                        wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    // If the shared string table is missing, something 
                    // is wrong. Return the index that is in
                    // the cell. Otherwise, look up the correct text in 
                    // the table.
                    if (stringTable is not null)
                    {
                        value =
                            stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }

                    break;

                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }
        }
    }

    return value;
}