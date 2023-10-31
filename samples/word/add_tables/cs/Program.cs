#nullable disable

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// Take the data from a two-dimensional array and build a table at the 
// end of the supplied document.
static void AddTable(string fileName, string[,] data)
{
    using (var document = WordprocessingDocument.Open(fileName, true))
    {

        var doc = document.MainDocumentPart.Document;

        Table table = new Table();

        TableProperties props = new TableProperties(
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

        for (var i = 0; i <= data.GetUpperBound(0); i++)
        {
            var tr = new TableRow();
            for (var j = 0; j <= data.GetUpperBound(1); j++)
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
        doc.Body.Append(table);
        doc.Save();
    }
}