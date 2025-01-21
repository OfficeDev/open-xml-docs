// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

static void GetCommentsFromDocument(string fileName)
{
    // <Snippet1>
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fileName, false))
    {
        if (wordDoc.MainDocumentPart is null || wordDoc.MainDocumentPart.WordprocessingCommentsPart is null)
        {
            throw new System.ArgumentNullException("MainDocumentPart and/or WordprocessingCommentsPart is null.");
        }
        // </Snippet1>

        // <Snippet2>
        WordprocessingCommentsPart commentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart;

        if (commentsPart is not null && commentsPart.Comments is not null)
        {
            foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
            {
                Console.WriteLine(comment.InnerText);
            }
        }
        // </Snippet2>
    }
}
// </Snippet0>

GetCommentsFromDocument(args[0]);
