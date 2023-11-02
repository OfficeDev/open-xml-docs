using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Text.RegularExpressions;

SearchAndReplace(args[0]);

// To search and replace content in a document part.
static void SearchAndReplace(string document)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
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

        Regex regexText = new Regex("Hello world!");
        docText = regexText.Replace(docText, "Hi Everyone!");

        using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
        {
            sw.Write(docText);
        }
    }
}