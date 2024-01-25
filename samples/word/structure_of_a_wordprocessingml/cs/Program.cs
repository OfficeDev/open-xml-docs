// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

static void CreateWordDoc(string filepath, string msg)
{
    using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
    {
        // Add a main document part. 
        MainDocumentPart mainPart = doc.AddMainDocumentPart();

        // Create the document structure and add some text.
        mainPart.Document = new Document();
        Body body = mainPart.Document.AppendChild(new Body());
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());

        // String msg contains the text from the msg parameter"
        run.AppendChild(new Text(msg));
    }
}
// </Snippet0>

CreateWordDoc(args[0], args[1]);
