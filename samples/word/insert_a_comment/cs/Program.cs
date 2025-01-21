using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;


// <Snippet>
// Insert a comment on the first paragraph.
static void AddCommentOnFirstParagraph(string fileName, string author, string initials, string comment)
{
    // Use the file name and path passed in as an 
    // argument to open an existing Wordprocessing document. 
    // <Snippet1>
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
    // </Snippet1>
    {
        if (document.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        WordprocessingCommentsPart wordprocessingCommentsPart = document.MainDocumentPart.WordprocessingCommentsPart ?? document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();

        // Locate the first paragraph in the document.
        // <Snippet2>
        Paragraph firstParagraph = document.MainDocumentPart.Document.Descendants<Paragraph>().First();
        wordprocessingCommentsPart.Comments ??= new Comments();
        string id = "0";
        // </Snippet2>

        // Verify that the document contains a 
        // WordProcessingCommentsPart part; if not, add a new one.
        // <Snippet3>
        if (document.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Count() > 0)
        {
            if (wordprocessingCommentsPart.Comments.HasChildren)
            {
                // Obtain an unused ID.
                id = (wordprocessingCommentsPart.Comments.Descendants<Comment>().Select(e =>
                {
                    if (e.Id is not null && e.Id.Value is not null)
                    {
                        return int.Parse(e.Id.Value);
                    }
                    else
                    {
                        throw new ArgumentNullException("Comment id and/or value are null.");
                    }
                })
                    .Max() + 1).ToString();
            }
        }
        // </Snippet3>

        // Compose a new Comment and add it to the Comments part.
        // <Snippet4>
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
        wordprocessingCommentsPart.Comments.AppendChild(cmt);
        // </Snippet4>

        // Specify the text range for the Comment. 
        // Insert the new CommentRangeStart before the first run of paragraph.
        // <Snippet5>
        firstParagraph.InsertBefore(new CommentRangeStart()
        { Id = id }, firstParagraph.GetFirstChild<Run>());

        // Insert the new CommentRangeEnd after last run of paragraph.
        var cmtEnd = firstParagraph.InsertAfter(new CommentRangeEnd()
        { Id = id }, firstParagraph.Elements<Run>().Last());

        // Compose a run with CommentReference and insert it.
        firstParagraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
        // </Snippet5>
    }
}
// </Snippet>

// <Snippet6>
string fileName = args[0];
string author = args[1];
string initials = args[2];
string comment = args[3];

AddCommentOnFirstParagraph(fileName, author, initials, comment);
// </Snippet6>
