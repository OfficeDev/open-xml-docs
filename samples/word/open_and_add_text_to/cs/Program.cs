// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

static void OpenAndAddTextToWordDocument(string filepath, string txt)
{
    // <Snippet1>
    // Open a WordprocessingDocument for editing using the filepath.
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {

        if (wordprocessingDocument is null)
        {
            throw new ArgumentNullException(nameof(wordprocessingDocument));
        }
        // </Snippet1>

        // <Snippet2>
        // Assign a reference to the existing document body.
        MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();
        mainDocumentPart.Document.Body ??= mainDocumentPart.Document.AppendChild(new Body());
        Body body = wordprocessingDocument.MainDocumentPart!.Document!.Body!;
        // </Snippet2>

        // <Snippet3>
        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
        // </Snippet3>
    }
}
// </Snippet0>

// <Snippet4>
string file = args[0];
string txt = args[1];

OpenAndAddTextToWordDocument(args[0], args[1]);
// </Snippet4>
