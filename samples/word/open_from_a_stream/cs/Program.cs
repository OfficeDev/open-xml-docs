
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;


static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
{
    // Open a WordProcessingDocument based on a stream.
    WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true);

    // Assign a reference to the existing document body.
    MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
    mainDocumentPart.Document ??= new Document();
    Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

    // Add new text.
    Paragraph para = body.AppendChild(new Paragraph());
    Run run = para.AppendChild(new Run());
    run.AppendChild(new Text(txt));

    // Dispose the document handle.
    wordprocessingDocument.Dispose();

    // Caller must close the stream.
}
