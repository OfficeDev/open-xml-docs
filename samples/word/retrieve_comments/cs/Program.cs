using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

static void GetCommentsFromDocument(string fileName)
{
    using (WordprocessingDocument wordDoc =
        WordprocessingDocument.Open(fileName, false))
    {
        if (wordDoc.MainDocumentPart is null || wordDoc.MainDocumentPart.WordprocessingCommentsPart is null)
        {
            throw new System.NullReferenceException("MainDocumentPart and/or WordprocessingCommentsPart is null.");
        }

        WordprocessingCommentsPart commentsPart =
            wordDoc.MainDocumentPart.WordprocessingCommentsPart;

        if (commentsPart != null && commentsPart.Comments != null)
        {
            foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
            {
                Console.WriteLine(comment.InnerText);
            }
        }
    }
}