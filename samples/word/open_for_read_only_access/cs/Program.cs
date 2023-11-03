
    using System.IO;
    using System.IO.Packaging;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;


    public static void OpenWordprocessingDocumentReadonly(string filepath)
    {
        // Open a WordprocessingDocument based on a filepath.
        using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Open(filepath, false))
        {
            // Assign a reference to the existing document body.  
            Body body = wordDocument.MainDocumentPart.Document.Body;

            // Attempt to add some text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

            // Call Save to generate an exception and show that access is read-only.
            // wordDocument.MainDocumentPart.Document.Save();
        }
    }

    public static void OpenWordprocessingPackageReadonly(string filepath)
    {
        // Open System.IO.Packaging.Package.
        Package wordPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);

        // Open a WordprocessingDocument based on a package.
        using (WordprocessingDocument wordDocument = 
            WordprocessingDocument.Open(wordPackage))
        {
            // Assign a reference to the existing document body. 
            Body body = wordDocument.MainDocumentPart.Document.Body;

            // Attempt to add some text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingPackageReadonly"));

            // Call Save to generate an exception and show that access is read-only.
            // wordDocument.MainDocumentPart.Document.Save();
        }

        // Close the package.
        wordPackage.Close();
    }
