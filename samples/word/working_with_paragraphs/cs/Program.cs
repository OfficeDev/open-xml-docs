// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;


static void WriteToWordDoc(string filepath, string txt)
{
    // Open a WordprocessingDocument for editing using the filepath.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {
        if (wordprocessingDocument is null)
        {
            throw new ArgumentNullException(nameof(wordprocessingDocument));
        }
        // Assign a reference to the existing document body.
        MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();
        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

        // Add a paragraph with some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
    }
}
// </Snippet0>

WriteToWordDoc(args[0], args[1]);