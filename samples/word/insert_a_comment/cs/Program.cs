#nullable disable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

// Insert a comment on the first paragraph.
static void AddCommentOnFirstParagraph(string fileName,
    string author, string initials, string comment)
{
    // Use the file name and path passed in as an 
    // argument to open an existing Wordprocessing document. 
    using (WordprocessingDocument document =
        WordprocessingDocument.Open(fileName, true))
    {
        // Locate the first paragraph in the document.
        Paragraph firstParagraph =
            document.MainDocumentPart.Document.Descendants<Paragraph>().First();
        Comments comments = null;
        string id = "0";

        // Verify that the document contains a 
        // WordProcessingCommentsPart part; if not, add a new one.
        if (document.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Count() > 0)
        {
            comments = document.MainDocumentPart.WordprocessingCommentsPart.Comments;
            if (comments.HasChildren)
            {
                // Obtain an unused ID.
                id = (comments.Descendants<Comment>().Select(e => int.Parse(e.Id.Value)).Max() + 1).ToString();
            }
        }
        else
        {
            // No WordprocessingCommentsPart part exists, so add one to the package.
            WordprocessingCommentsPart commentPart = document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            commentPart.Comments = new Comments();
            comments = commentPart.Comments;
        }

        // Compose a new Comment and add it to the Comments part.
        Paragraph p = new Paragraph(new Run(new Text(comment)));
        Comment cmt =
            new Comment()
            {
                Id = id,
                Author = author,
                Initials = initials,
                Date = DateTime.Now
            };
        cmt.AppendChild(p);
        comments.AppendChild(cmt);
        comments.Save();

        // Specify the text range for the Comment. 
        // Insert the new CommentRangeStart before the first run of paragraph.
        firstParagraph.InsertBefore(new CommentRangeStart()
        { Id = id }, firstParagraph.GetFirstChild<Run>());

        // Insert the new CommentRangeEnd after last run of paragraph.
        var cmtEnd = firstParagraph.InsertAfter(new CommentRangeEnd()
        { Id = id }, firstParagraph.Elements<Run>().Last());

        // Compose a run with CommentReference and insert it.
        firstParagraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
    }
}