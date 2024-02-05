// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;


static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
{
    // <Snippet1>
    // Open a WordProcessingDocument based on a stream.
    WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true);
    // </Snippet1>

    // <Snippet2>
    // Assign a reference to the document body.
    MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
    mainDocumentPart.Document ??= new Document();
    Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());
    // </Snippet2>

    // <Snippet3>
    // Add new text.
    Paragraph para = body.AppendChild(new Paragraph());
    Run run = para.AppendChild(new Run());
    run.AppendChild(new Text(txt));
    // </Snippet3>

    // Dispose the document handle.
    wordprocessingDocument.Dispose();

    // Caller must close the stream.
}
// </Snippet0>

// <Snippet4>
string filePath = args[0];
string txt = args[1];

using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
{
    OpenAndAddToWordprocessingStream(fileStream, txt);
}
// </Snippet4>
