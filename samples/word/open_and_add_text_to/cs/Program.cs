
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;


    public static void OpenAndAddTextToWordDocument(string filepath, string txt)
    {   
        // Open a WordprocessingDocument for editing using the filepath.
        WordprocessingDocument wordprocessingDocument = 
            WordprocessingDocument.Open(filepath, true);

        // Assign a reference to the existing document body.
        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
        
        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
        
        // Close the handle explicitly.
        wordprocessingDocument.Close();
    }
