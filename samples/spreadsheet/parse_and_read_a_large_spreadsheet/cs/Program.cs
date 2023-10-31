#nullable disable

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

// The DOM approach.
// Note that the code below works only for cells that contain numeric values.
// 
static void ReadExcelFileDOM(string fileName)
{
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        string text;
        foreach (Row r in sheetData.Elements<Row>())
        {
            foreach (Cell c in r.Elements<Cell>())
            {
                text = c.CellValue.Text;
                Console.Write(text + " ");
            }
        }
        Console.WriteLine();
        Console.ReadKey();
    }
}

// The SAX approach.
static void ReadExcelFileSAX(string fileName)
{
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

        OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
        string text;
        while (reader.Read())
        {
            if (reader.ElementType == typeof(CellValue))
            {
                text = reader.GetText();
                Console.Write(text + " ");
            }
        }
        Console.WriteLine();
        Console.ReadKey();
    }
}