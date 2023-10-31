#nullable disable

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// Insert a table into a word processing document.
static void CreateTable(string fileName)
{
    // Use the file name and path passed in as an argument 
    // to open an existing Word 2007 document.

    using (WordprocessingDocument doc
        = WordprocessingDocument.Open(fileName, true))
    {
        // Create an empty table.
        Table table = new Table();

        // Create a TableProperties object and specify its border information.
        TableProperties tblProp = new TableProperties(
            new TableBorders(
                new TopBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new BottomBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new LeftBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new RightBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new InsideHorizontalBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                },
                new InsideVerticalBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(BorderValues.Dashed),
                    Size = 24
                }
            )
        );

        // Append the TableProperties object to the empty table.
        table.AppendChild<TableProperties>(tblProp);

        // Create a row.
        TableRow tr = new TableRow();

        // Create a cell.
        TableCell tc1 = new TableCell();

        // Specify the width property of the table cell.
        tc1.Append(new TableCellProperties(
            new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

        // Specify the table cell content.
        tc1.Append(new Paragraph(new Run(new Text("some text"))));

        // Append the table cell to the table row.
        tr.Append(tc1);

        // Create a second table cell by copying the OuterXml value of the first table cell.
        TableCell tc2 = new TableCell(tc1.OuterXml);

        // Append the table cell to the table row.
        tr.Append(tc2);

        // Append the table row to the table.
        table.Append(tr);

        // Append the table to the document.
        doc.MainDocumentPart.Document.Body.Append(table);
    }
}