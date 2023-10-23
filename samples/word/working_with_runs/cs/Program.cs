
    public static void WriteToWordDoc(string filepath, string txt)
    {
        // Open a WordprocessingDocument for editing using the filepath.
        using (WordprocessingDocument wordprocessingDocument =
             WordprocessingDocument.Open(filepath, true))
        {
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());

            // Apply bold formatting to the run.
            RunProperties runProperties = run.AppendChild(new RunProperties(new Bold()));   
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

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());

            // Apply bold formatting to the run.
            RunProperties runProperties = run.AppendChild(new RunProperties(new Bold()));   
            run.AppendChild(new Text(txt));                
        }
    }