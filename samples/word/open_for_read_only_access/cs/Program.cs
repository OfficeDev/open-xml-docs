// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO.Packaging;
using System.IO;


static void OpenWordprocessingDocumentReadonly(string filepath)
{
    // <Snippet1>
    // Open a WordprocessingDocument based on a filepath.
    using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filepath, false))
    // </Snippet1>
    {
        if (wordProcessingDocument is null)
        {
            throw new ArgumentNullException(nameof(wordProcessingDocument));
        }

        // <Snippet4>
        // Assign a reference to the existing document body or create a new one if it is null.
        MainDocumentPart mainDocumentPart = wordProcessingDocument.MainDocumentPart ?? wordProcessingDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();

        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());
        // </Snippet4>

        // <Snippet5>
        // Attempt to add some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

        // Call Save to generate an exception and show that access is read-only.
        // mainDocumentPart.Document.Save();
        // </Snippet5>
    }
}

static void OpenWordprocessingPackageReadonly(string filepath)
{
    // <Snippet2>
    // Open System.IO.Packaging.Package.
    Package wordPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);

    // Open a WordprocessingDocument based on a package.
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(wordPackage))
    // </Snippet2>
    {
        // Assign a reference to the existing document body or create a new one if it is null.
        MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart ?? wordDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();

        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

        // Attempt to add some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingPackageReadonly"));

        // Call Save to generate an exception and show that access is read-only.
        // mainDocumentPart.Document.Save();
    }

    // Close the package.
    wordPackage.Close();
}


static void OpenWordprocessingStreamReadonly(string filepath)
{
    // <Snippet3>
    // Get a stream of the wordprocessing document
    using (FileStream fileStream = new FileStream(filepath, FileMode.Open))

    // Open a WordprocessingDocument for read-only access based on a stream.
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(fileStream, false))
    // </Snippet3>
    {


        // Assign a reference to the existing document body or create a new one if it is null.
        MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart ?? wordDocument.AddMainDocumentPart();
        mainDocumentPart.Document ??= new Document();

        Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

        // Attempt to add some text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingStreamReadonly"));

        // Call Save to generate an exception and show that access is read-only.
        // mainDocumentPart.Document.Save();
    }
}
// </Snippet0>

// <Snippet6>
OpenWordprocessingDocumentReadonly(args[0]);
// </Snippet6>

// <Snippet7>
OpenWordprocessingPackageReadonly(args[0]);
// </Snippet7>

// <Snippet8>
OpenWordprocessingStreamReadonly(args[0]);
// </Snippet8>

