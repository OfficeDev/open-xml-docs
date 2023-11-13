
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

OpenAndAddTextToWordDocument(args[0], args[1]);

static void OpenAndAddTextToWordDocument(string filepath, string txt)
{
    // Open a WordprocessingDocument for editing using the filepath.
    WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true);

    if (wordprocessingDocument is null)
    {
        throw new ArgumentNullException(nameof(wordprocessingDocument));
    }

    // Assign a reference to the existing document body.
    MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
    mainDocumentPart.Document ??= new Document();
    mainDocumentPart.Document.Body ??= mainDocumentPart.Document.AppendChild(new Body());
    Body body = wordprocessingDocument.MainDocumentPart!.Document!.Body!;

    // Add new text.
    Paragraph para = body.AppendChild(new Paragraph());
    Run run = para.AppendChild(new Run());
    run.AppendChild(new Text(txt));

    // Dispose the handle explicitly.
    wordprocessingDocument.Dispose();
}
