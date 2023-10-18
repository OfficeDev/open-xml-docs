using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Linq;

AddCommentToPresentation(args[0], args[1], args[2], string.Join(' ', args[3..]));

static void AddCommentToPresentation(string file, string initials, string name, string text)
{
    using (PresentationDocument doc = PresentationDocument.Open(file, true))
    {

        // Declare a CommentAuthorsPart object.
        CommentAuthorsPart authorsPart;

        // Verify that there is an existing comment authors part. 
        if (doc.PresentationPart.CommentAuthorsPart == null)
        {
            // If not, add a new one.
            authorsPart = doc.PresentationPart.AddNewPart<CommentAuthorsPart>();
        }
        else
        {
            authorsPart = doc.PresentationPart.CommentAuthorsPart;
        }

        // Verify that there is a comment author list in the comment authors part.
        if (authorsPart.CommentAuthorList == null)
        {
            // If not, add a new one.
            authorsPart.CommentAuthorList = new CommentAuthorList();
        }

        // Declare a new author ID.
        uint authorId = 0;
        CommentAuthor author = null;

        // If there are existing child elements in the comment authors list...
        if (authorsPart.CommentAuthorList.HasChildren)
        {
            // Verify that the author passed in is on the list.
            var authors = authorsPart.CommentAuthorList.Elements<CommentAuthor>().Where(a => a.Name == name && a.Initials == initials);

            // If so...
            if (authors.Any())
            {
                // Assign the new comment author the existing author ID.
                author = authors.First();
                authorId = author.Id;
            }

            // If not...
            if (author == null)
            {
                // Assign the author passed in a new ID                        
                authorId = authorsPart.CommentAuthorList.Elements<CommentAuthor>().Select(a => a.Id.Value).Max();
            }
        }

        // If there are no existing child elements in the comment authors list.
        if (author == null)
        {

            authorId++;

            // Add a new child element(comment author) to the comment author list.
            author = authorsPart.CommentAuthorList.AppendChild<CommentAuthor>
                (new CommentAuthor()
                {
                    Id = authorId,
                    Name = name,
                    Initials = initials,
                    ColorIndex = 0
                });
        }

        // Get the first slide, using the GetFirstSlide method.
        SlidePart slidePart1 = GetFirstSlide(doc);

        // Declare a comments part.
        SlideCommentsPart commentsPart;

        // Verify that there is a comments part in the first slide part.
        if (slidePart1.GetPartsOfType<SlideCommentsPart>().Count() == 0)
        {
            // If not, add a new comments part.
            commentsPart = slidePart1.AddNewPart<SlideCommentsPart>();
        }
        else
        {
            // Else, use the first comments part in the slide part.
            commentsPart = slidePart1.GetPartsOfType<SlideCommentsPart>().First();
        }

        // If the comment list does not exist.
        if (commentsPart.CommentList == null)
        {
            // Add a new comments list.
            commentsPart.CommentList = new CommentList();
        }

        // Get the new comment ID.
        uint commentIdx = author.LastIndex == null ? 1 : author.LastIndex + 1;
        author.LastIndex = commentIdx;

        // Add a new comment.
        Comment comment = commentsPart.CommentList.AppendChild<Comment>(
            new Comment()
            {
                AuthorId = authorId,
                Index = commentIdx,
                DateTime = DateTime.Now
            });

        // Add the position child node to the comment element.
        comment.Append(
            new Position() { X = 100, Y = 200 },
            new Text() { Text = text });

        // Save the comment authors part.
        authorsPart.CommentAuthorList.Save();

        // Save the comments part.
        commentsPart.CommentList.Save();
    }
}

// Get the slide part of the first slide in the presentation document.
static SlidePart GetFirstSlide(PresentationDocument presentationDocument)
{
    // Get relationship ID of the first slide
    PresentationPart part = presentationDocument.PresentationPart;
    SlideId slideId = part.Presentation.SlideIdList.GetFirstChild<SlideId>();
    string relId = slideId.RelationshipId;

    // Get the slide part by the relationship ID.
    SlidePart slidePart = (SlidePart)part.GetPartById(relId);

    return slidePart;
}
