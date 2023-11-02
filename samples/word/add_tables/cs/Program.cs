using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

AddTable(args[0], args[1]);

// Take the data from a two-dimensional array and build a table at the 
// end of the supplied document.
static void AddTable(string fileName, string json)
{
    // read the data from the json file
    var data = System.Text.Json.JsonSerializer.Deserialize<string[][]>(json);

    if (data is not null)
    {
        using (var document = WordprocessingDocument.Open(fileName, true))
        {
            if (document.MainDocumentPart is null || document.MainDocumentPart.Document.Body is null)
            {
                throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
            }

            var doc = document.MainDocumentPart.Document;

            Table table = new();

            TableProperties props = new(
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new BottomBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new LeftBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new RightBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new InsideHorizontalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new InsideVerticalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                }));

            table.AppendChild<TableProperties>(props);

            for (var i = 0; i < data.Length; i++)
            {
                var tr = new TableRow();
                for (var j = 0; j < data[i].Length; j++)
                {
                    var tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(data[i][j]))));

                    // Assume you want columns that are automatically sized.
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                    tr.Append(tc);
                }
                table.Append(tr);
            }
            doc.Body.Append(table);
            doc.Save();
        }
    }
}