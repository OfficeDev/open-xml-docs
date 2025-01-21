using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

// <Snippet2>
string fileName = args[0];

AddTable(fileName, new string[,] {
    { "Hawaii", "HI" },
    { "California", "CA" },
    { "New York", "NY" },
    { "Massachusetts", "MA" }
});
// </Snippet2>

// Take the data from a two-dimensional array and build a table at the 
// end of the supplied document.
// <Snippet0>
// <Snippet1>
static void AddTable(string fileName, string[,] data)
// </Snippet1>
{
    if (data is not null)
    {
        // <Snippet3>
        using (var document = WordprocessingDocument.Open(fileName, true))
        {
            if (document.MainDocumentPart is null || document.MainDocumentPart.Document.Body is null)
            {
                throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
            }

            var doc = document.MainDocumentPart.Document;
            // </Snippet3>
            // <Snippet4>
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
            // </Snippet4>
            // <Snippet5>
            for (var i = 0; i < data.GetUpperBound(0); i++)
            {
                var tr = new TableRow();
                for (var j = 0; j < data.GetUpperBound(1); j++)
                {
                    var tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                    // Assume you want columns that are automatically sized.
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                    tr.Append(tc);
                }
                table.Append(tr);
            }
            // </Snippet5>
            // <Snippet6>
            doc.Body.Append(table);
            // </Snippet6>
        }
    }
}
// </Snippet0>