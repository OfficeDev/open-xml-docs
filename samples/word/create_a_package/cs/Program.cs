
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Text;

CreateNewWordDocument(args[0]);
// <Snippet0>
// To create a new package as a Word document.
static void CreateNewWordDocument(string document)
{
    // <Snippet1>
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
        // </Snippet1>
    {
        // Set the content of the document so that Word can open it.
        MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

        SetMainDocumentContent(mainPart);
    }
}

// Set the content of MainDocumentPart.
static void SetMainDocumentContent(MainDocumentPart part)
{
    const string docXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
                            <w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                              <w:body>
                                <w:p>
                                  <w:r>
                                    <w:t>Hello World</w:t>
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
// </Snippet0>
