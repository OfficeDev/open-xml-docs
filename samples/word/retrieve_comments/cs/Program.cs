using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

GetCommentsFromDocument(args[0]);

static void GetCommentsFromDocument(string fileName)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fileName, false))
    {
        if (wordDoc.MainDocumentPart is null || wordDoc.MainDocumentPart.WordprocessingCommentsPart is null)
        {
            throw new System.InvalidOperationException("MainDocumentPart and/or WordprocessingCommentsPart is null.");
        }

        WordprocessingCommentsPart commentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart;

        if (commentsPart is not null && commentsPart.Comments is not null)
        {
            foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
            {
                Console.WriteLine(comment.InnerText);
            }
        }
    }
}