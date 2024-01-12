// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

static void WriteToWordDoc(string filepath, string txt)
{
    // Open a WordprocessingDocument for editing using the filepath.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {
        // Assign a reference to the existing document body.
        MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();
        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());

        // Apply bold formatting to the run.
        RunProperties runProperties = run.AppendChild(new RunProperties(new Bold()));
        run.AppendChild(new Text(txt));
    }
}
// </Snippet0>
WriteToWordDoc(args[0], args[1]);