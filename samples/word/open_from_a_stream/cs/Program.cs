
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;


    public static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
    {
        // Open a WordProcessingDocument based on a stream.
        WordprocessingDocument wordprocessingDocument = 
            WordprocessingDocument.Open(stream, true);
        
        // Assign a reference to the existing document body.
        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));

        // Close the document handle.
        wordprocessingDocument.Close();
        
        // Caller must close the stream.
    }
