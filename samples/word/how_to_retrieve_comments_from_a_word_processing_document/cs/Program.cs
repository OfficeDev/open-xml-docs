
using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

static void GetCommentsFromDocument(string fileName)
{
    using (WordprocessingDocument wordDoc =
        WordprocessingDocument.Open(fileName, false))
    {
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