
    using System;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;


    // To get the contents of a document part.
    public static string GetCommentsFromDocument(string document)
    {
        string comments = null;

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            WordprocessingCommentsPart WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart;

            using (StreamReader streamReader = new StreamReader(WordprocessingCommentsPart.GetStream()))
            {
                comments = streamReader.ReadToEnd();
            }
        }
        return comments;
    }
