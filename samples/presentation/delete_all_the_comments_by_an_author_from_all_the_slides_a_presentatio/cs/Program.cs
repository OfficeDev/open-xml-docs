using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System.Linq;

DeleteCommentsByAuthorInPresentation(args[0], args[1]);

// Remove all the comments in the slides by a certain author.
static void DeleteCommentsByAuthorInPresentation(string fileName, string author)
{
    using (PresentationDocument doc = PresentationDocument.Open(fileName, true))
    {
        // Get the specified comment author.
        IEnumerable<CommentAuthor>? commentAuthors = doc.PresentationPart?.CommentAuthorsPart?.CommentAuthorList?.Elements<CommentAuthor>()
            .Where(e => e.Name is not null && e.Name.Value is not null && e.Name.Value.Equals(author));

        if (commentAuthors is null)
        {
            return;
        }

        // Iterate through all the matching authors.
        foreach (CommentAuthor commentAuthor in commentAuthors)
        {
            UInt32Value? authorId = commentAuthor.Id;
            IEnumerable<SlidePart>? slideParts = doc.PresentationPart?.SlideParts;

            // If there's no author ID or slide parts, return.
            if (authorId is null || slideParts is null)
            {
                return;
            }

            // Iterate through all the slides and get the slide parts.
            foreach (SlidePart slide in slideParts)
            {
                SlideCommentsPart? slideCommentsPart = slide.SlideCommentsPart;

                // Get the list of comments.
                if (slideCommentsPart is not null && slide.SlideCommentsPart?.CommentList is not null)
                {
                    IEnumerable<Comment> commentList = slideCommentsPart.CommentList.Elements<Comment>().Where(e => e.AuthorId is not null && e.AuthorId == authorId.Value);
                    List<Comment> comments = new List<Comment>();
                    comments = commentList.ToList<Comment>();

                    foreach (Comment comm in comments)
                    {
                        // Delete all the comments by the specified author.

                        slideCommentsPart.CommentList.RemoveChild<Comment>(comm);
                    }

                    // If the commentPart has no existing comment.
                    if (slideCommentsPart.CommentList.ChildElements.Count == 0)
                        // Delete this part.
                        slide.DeletePart(slideCommentsPart);
                }
            }
            // Delete the comment author from the comment authors part.
            doc.PresentationPart?.CommentAuthorsPart?.CommentAuthorList.RemoveChild(commentAuthor);
        }
    }
}
