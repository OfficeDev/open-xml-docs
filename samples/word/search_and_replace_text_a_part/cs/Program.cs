using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Text.RegularExpressions;

// <Snippet2>
SearchAndReplace(args[0]);
// </Snippet2>

// To search and replace content in a document part.
// <Snippet0>
static void SearchAndReplace(string document)
{
    // <Snippet1>
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
        // </Snippet1>
    {
        string? docText = null;


        if (wordDoc.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
        {
            docText = sr.ReadToEnd();
        }

        Regex regexText = new Regex("Hello World!");
        docText = regexText.Replace(docText, "Hi Everyone!");

        using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
        {
            sw.Write(docText);
        }
    }
}
// </Snippet0>
