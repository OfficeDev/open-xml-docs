
    public static void WriteToWordDoc(string filepath, string txt)
    {
        // Open a WordprocessingDocument for editing using the filepath.
        using (WordprocessingDocument wordprocessingDocument =
             WordprocessingDocument.Open(filepath, true))
        {
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add a paragraph with some text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));
        }
    }

    public static void WriteToWordDoc(string filepath, string txt)
    {
        // Open a WordprocessingDocument for editing using the filepath.
        using (WordprocessingDocument wordprocessingDocument =
             WordprocessingDocument.Open(filepath, true))
        {
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add a paragraph with some text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));
        }
    }