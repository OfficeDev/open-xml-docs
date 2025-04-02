

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;

CopySheetDOM(args[0]);
CopySheetSAX(args[1]);

// <Snippet0>
void CopySheetDOM(string path)
{
    Console.WriteLine("Starting DOM method");

    Stopwatch sw = new();
    sw.Start();
    // <Snippet1>
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, true))
    {
        // Get the first sheet
        WorksheetPart? worksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault();

        if (worksheetPart is not null)
        // </Snippet1>
        {
            // <Snippet2>
            // Add a new WorksheetPart
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart!.AddNewPart<WorksheetPart>();

            // Make a copy of the original worksheet
            Worksheet newWorksheet = (Worksheet)worksheetPart.Worksheet.Clone();

            // Add the new worksheet to the new worksheet part
            newWorksheetPart.Worksheet = newWorksheet;
            // </Snippet2>

            Sheets? sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            if (sheets is null)
            {
                spreadsheetDocument.WorkbookPart.Workbook.AddChild(new Sheets());
            }

            // <Snippet3>
            // Find the new WorksheetPart's Id and create a new sheet id
            string id = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);
            uint newSheetId = (uint)(sheets!.ChildElements.Count + 1);

            // Append a new Sheet with the WorksheetPart's Id and sheet id to the Sheets element
            sheets.AppendChild(new Sheet() { Name = "My New Sheet", SheetId = newSheetId, Id = id });
            // </Snippet3>
        }
    }

    sw.Stop();

    Console.WriteLine($"DOM method took {sw.Elapsed.TotalSeconds} seconds");
}
// </Snippet0>

// <Snippet99>
void CopySheetSAX(string path)
{
    Console.WriteLine("Starting SAX method");

    Stopwatch sw = new();
    sw.Start();
    // <Snippet4>
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, true))
    {
        // Get the first sheet
        WorksheetPart? worksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault();

        if (worksheetPart is not null)
        // </Snippet4>
        {
            // <Snippet5>
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart!.AddNewPart<WorksheetPart>();
            // </Snippet5>

            // <Snippet6>
            using (OpenXmlReader reader = OpenXmlPartReader.Create(worksheetPart))
            using (OpenXmlWriter writer = OpenXmlPartWriter.Create(newWorksheetPart))
            // </Snippet6>
            {
                // <Snippet7>
                // Write the XML declaration with the version "1.0".
                writer.WriteStartDocument();

                // Read the elements from the original worksheet part
                while (reader.Read())
                {
                    // If the ElementType is CellValue it's necessary to explicitly add the inner text of the element
                    // or the CellValue element will be empty
                    if (reader.ElementType == typeof(CellValue))
                    {
                        if (reader.IsStartElement)
                        {
                            writer.WriteStartElement(reader);
                            writer.WriteString(reader.GetText());
                        }
                        else if (reader.IsEndElement)
                        {
                            writer.WriteEndElement();
                        }
                    }
                    // For other elements write the start and end elements
                    else
                    {
                        if (reader.IsStartElement)
                        {
                            writer.WriteStartElement(reader);
                        }
                        else if (reader.IsEndElement)
                        {
                            writer.WriteEndElement();
                        }
                    }
                }
                // </Snippet7>
            }

            // <Snippet8>
            Sheets? sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            if (sheets is null)
            {
                spreadsheetDocument.WorkbookPart.Workbook.AddChild(new Sheets());
            }

            string id = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);
            uint newSheetId = (uint)(sheets!.ChildElements.Count + 1);

            sheets.AppendChild(new Sheet() { Name = "My New Sheet", SheetId = newSheetId, Id = id });
            // </Snippet8>

            sw.Stop();

            Console.WriteLine($"SAX method took {sw.Elapsed.TotalSeconds} seconds");
        }
    }
}
// </Snippet99>
