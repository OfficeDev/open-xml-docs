
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// <Snippet2>
CreateWordprocessingDocument(args[0]);
// </Snippet2>

// <Snippet0>
static void CreateWordprocessingDocument(string filepath)
{
    // Create a document by supplying the filepath. 
    // <Snippet1>
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
    {
        // </Snippet1>
        // Add a main document part. 
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

        // Create the document structure and add some text.
        mainPart.Document = new Document();
        Body body = mainPart.Document.AppendChild(new Body());
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
    }
    // </Snippet0>
}
