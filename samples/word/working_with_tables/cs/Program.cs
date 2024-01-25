using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

// <Snippet0>
static string InsertTableInDoc(string filepath)
{
    // Open a WordprocessingDocument for editing using the filepath.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {
        // Assign a reference to the existing document body or add one if necessary.

        if (wordprocessingDocument.MainDocumentPart is null)
        {
            wordprocessingDocument.AddMainDocumentPart();
        }

        if (wordprocessingDocument.MainDocumentPart!.Document is null)
        {
            wordprocessingDocument.MainDocumentPart.Document = new Document();
        }

        if (wordprocessingDocument.MainDocumentPart.Document.Body is null)
        {
            wordprocessingDocument.MainDocumentPart.Document.Body = new Body();
        }

        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

        // Create a table.
        Table tbl = new Table();

        // Set the style and width for the table.
        TableProperties tableProp = new TableProperties();
        TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };

        // Make the table width 100% of the page width.
        TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

        // Apply
        tableProp.Append(tableStyle, tableWidth);
        tbl.AppendChild(tableProp);

        // Add 3 columns to the table.
        TableGrid tg = new TableGrid(new GridColumn(), new GridColumn(), new GridColumn());
        tbl.AppendChild(tg);

        // Create 1 row to the table.
        TableRow tr1 = new TableRow();

        // Add a cell to each column in the row.
        TableCell tc1 = new TableCell(new Paragraph(new Run(new Text("1"))));
        TableCell tc2 = new TableCell(new Paragraph(new Run(new Text("2"))));
        TableCell tc3 = new TableCell(new Paragraph(new Run(new Text("3"))));
        tr1.Append(tc1, tc2, tc3);

        // Add row to the table.
        tbl.AppendChild(tr1);

        // Add the table to the document
        body.AppendChild(tbl);

        return tbl.LocalName;
    }
}
// </Snippet0>

Console.WriteLine($"Inserted {InsertTableInDoc(args[0])}");