using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using Comment = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment;
using CommentList = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList;

DeleteCommentsByAuthorInPresentation(args[0], args[1]);
// <Snippet0>
// Remove all the comments in the slides by a certain x.
static void DeleteCommentsByAuthorInPresentation(string fileName, string author)
{
    // <Snippet1>
    using (PresentationDocument doc = PresentationDocument.Open(fileName, true))
    // </Snippet1>
    {
        // <Snippet2>
        // Get the modern comments.
        IEnumerable<Author>? commentAuthors = doc.PresentationPart?.authorsPart?.AuthorList.Elements<Author>()
            .Where(x => x.Name is not null && x.Name.HasValue && x.Name.Value!.Equals(author));
        // </Snippet2>

        if (commentAuthors is null)
        {
            return;
        }

        // <Snippet3>
        // Iterate through all the matching authors.
        foreach (Author commentAuthor in commentAuthors)
        {
            string? authorId = commentAuthor.Id;
            IEnumerable<SlidePart>? slideParts = doc.PresentationPart?.SlideParts;

            // If there's no author ID or slide parts or slide parts, return.
            if (authorId is null || slideParts is null)
            {
                return;
            }

            // Iterate through all the slides and get the slide parts.
            foreach (SlidePart slide in slideParts)
            {
                IEnumerable<PowerPointCommentPart>? slideCommentsParts = slide.commentParts;

                // Get the list of comments.
                if (slideCommentsParts is not null)
                {
                    IEnumerable<Tuple<PowerPointCommentPart, Comment>> commentsTup = slideCommentsParts
                        .SelectMany(scp => scp.CommentList.Elements<Comment>()
                        .Where(comment => comment.AuthorId is not null && comment.AuthorId == authorId)
                        .Select(c => new Tuple<PowerPointCommentPart, Comment>(scp, c)));

                    foreach (Tuple<PowerPointCommentPart, Comment> comment in commentsTup)
                    {
                        // Delete all the comments by the specified author.
                        comment.Item1.CommentList.RemoveChild(comment.Item2);

                        // If the commentPart has no existing comment.
                        if (comment.Item1.CommentList.ChildElements.Count == 0)
                        {
                            // Delete this part.
                            slide.DeletePart(comment.Item1);
                        }
                    }

                }
            }

            // Delete the comment author from the authors part.
            doc.PresentationPart?.authorsPart?.AuthorList.RemoveChild(commentAuthor);
        }
        // </Snippet3>
    }
}
// </Snippet0>
