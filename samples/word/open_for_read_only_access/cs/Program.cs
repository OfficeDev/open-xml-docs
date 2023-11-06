
using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

OpenWordprocessingDocumentReadonly(args[0]);

static void OpenWordprocessingDocumentReadonly(string filepath)
{
    // Open a WordprocessingDocument based on a filepath.
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, false))
    {
        if (wordDocument is null)
        {
            throw new ArgumentNullException(nameof(wordDocument));
        }
        // Assign a reference to the existing document body.
        MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart ?? wordDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();

        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

        // Attempt to add some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

        // Call Save to generate an exception and show that access is read-only.
        //mainDocumentPart.Document.Save();
    }
}

