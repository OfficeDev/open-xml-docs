
    using System.Text;
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;


    // To create a new package as a Word document.
    public static void CreateNewWordDocument(string document)
    {
       using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
       {
          // Set the content of the document so that Word can open it.
          MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

          SetMainDocumentContent(mainPart);
       }
    }

    // Set the content of MainDocumentPart.
    public static void SetMainDocumentContent(MainDocumentPart part)
    {
       const string docXml =
        @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> 
        <w:document xmlns:w=""https://schemas.openxmlformats.org/wordprocessingml/2006/main"">
            <w:body>
                <w:p>
                    <w:r>
                        <w:t>Hello world!</w:t>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>";

        using (Stream stream = part.GetStream())
        {
            byte[] buf = (new UTF8Encoding()).GetBytes(docXml);
            stream.Write(buf, 0, buf.Length);
        }
    }
