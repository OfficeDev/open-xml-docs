using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

ChangeTextInCell(args[0], args[1]);

// Change the text in a table in a word processing document.
static void ChangeTextInCell(string filePath, string txt)
{
    // Use the file name and path passed in as an argument to 
    // open an existing document.            
    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
    {
        if (doc.MainDocumentPart is null || doc.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        // Find the first table in the document.
        Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

        // Find the second row in the table.
        TableRow row = table.Elements<TableRow>().ElementAt(1);

        // Find the third cell in the row.
        TableCell cell = row.Elements<TableCell>().ElementAt(2);

        // Find the first paragraph in the table cell.
        Paragraph p = cell.Elements<Paragraph>().First();

        // Find the first run in the paragraph.
        Run r = p.Elements<Run>().First();

        // Set the text for the run.
        Text t = r.Elements<Text>().First();
        t.Text = txt;
    }
}