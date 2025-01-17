// <Snippet>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;


// <Snippet1>
// Delete comments by a specific author. Pass an empty string for the 
// author to delete all comments, by all authors.
static void DeleteComments(string fileName, string author = "")
// </Snippet1>
{
    // <Snippet3>
    // Get an existing Wordprocessing document.
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
    {

        if (document.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart is null.");
        }

        // Set commentPart to the document WordprocessingCommentsPart, 
        // if it exists.
        WordprocessingCommentsPart? commentPart = document.MainDocumentPart.WordprocessingCommentsPart;

        // If no WordprocessingCommentsPart exists, there can be no 
        // comments. Stop execution and return from the method.
        if (commentPart is null)
        {
            return;
        }
        // </Snippet3>

        // Create a list of comments by the specified author, or
        // if the author name is empty, all authors.
        // <Snippet4>
        List<Comment> commentsToDelete = commentPart.Comments.Elements<Comment>().ToList();
        // </Snippet4>

        // <Snippet5>
        if (!String.IsNullOrEmpty(author))
        {
            commentsToDelete = commentsToDelete.Where(c => c.Author == author).ToList();
        }
        // </Snippet5>

        // <Snippet6>
        IEnumerable<string?> commentIds = commentsToDelete.Where(r => r.Id is not null && r.Id.HasValue).Select(r => r.Id?.Value);
        // </Snippet6>

        // <Snippet7>
        // Delete each comment in commentToDelete from the 
        // Comments collection.
        foreach (Comment c in commentsToDelete)
        {
            if (c is not null)
            {
                c.Remove();
            }
        }
        // </Snippet7>

        // <Snippet8>
        Document doc = document.MainDocumentPart.Document;
        // </Snippet8>

        // <Snippet9>
        // Delete CommentRangeStart for each
        // deleted comment in the main document.
        List<CommentRangeStart> commentRangeStartToDelete = doc.Descendants<CommentRangeStart>()
                                                            .Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                                                            .ToList();

        foreach (CommentRangeStart c in commentRangeStartToDelete)
        {
            c.Remove();
        }

        // Delete CommentRangeEnd for each deleted comment in the main document.
        List<CommentRangeEnd> commentRangeEndToDelete = doc.Descendants<CommentRangeEnd>()
                                                        .Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                                                        .ToList();

        foreach (CommentRangeEnd c in commentRangeEndToDelete)
        {
            c.Remove();
        }

        // Delete CommentReference for each deleted comment in the main document.
        List<CommentReference> commentRangeReferenceToDelete = doc.Descendants<CommentReference>()
                                                                .Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                                                                .ToList();

        foreach (CommentReference c in commentRangeReferenceToDelete)
        {
            c.Remove();
        }
        // </Snippet9>
    }
}
// </Snippet>

// <Snippet2>
if (args is [{ } fileName, { } author])
{
    DeleteComments(fileName, author);
}
else if (args is [{ } fileName2])
{
    DeleteComments(fileName2);
}
// </Snippet2>
