
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

GetCommentsFromDocument(args[0]);

// To get the contents of a document part.
static string GetCommentsFromDocument(string document)
{
    string? comments = null;

    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
    {
        if (wordDoc is null)
        {
            throw new ArgumentNullException(nameof(wordDoc));
        }

        MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
        WordprocessingCommentsPart WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();

        using (StreamReader streamReader = new StreamReader(WordprocessingCommentsPart.GetStream()))
        {
            comments = streamReader.ReadToEnd();
        }
    }

    return comments;
}
