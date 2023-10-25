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

        // If the presentation does not contain a comment authors part, add a new one.
        PresentationPart presentationPart = doc.PresentationPart ?? doc.AddPresentationPart();

        // Verify that there is an existing comment authors part and add a new one if not.
        authorsPart = presentationPart.CommentAuthorsPart ?? presentationPart.AddNewPart<CommentAuthorsPart>();

        // Verify that there is a comment author list in the comment authors part and add one if not.
        CommentAuthorList authorList = authorsPart.CommentAuthorList ?? new CommentAuthorList();
        authorsPart.CommentAuthorList = authorList;

        // Declare a new author ID as either the max existing ID + 1 or 1 if there are no existing IDs.
        uint authorId = authorList.Elements<CommentAuthor>().Select(a => a.Id?.Value).Max() ?? 0;
        authorId++;
        // If there is an existing author with matching name and initials, use that author otherwise create a new CommentAuthor.
        var foo = authorList.Elements<CommentAuthor>().Where(a => a.Name == name && a.Initials == initials).FirstOrDefault();
        CommentAuthor author = foo ??
            authorList.AppendChild
                (new CommentAuthor()
                {
                    Id = authorId,
                    Name = name,
                    Initials = initials,
                    ColorIndex = 0
                });
        // get the author id
        authorId = author.Id ?? authorId;

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
        if (commentsPart.CommentList is null)
        {
            // Add a new comments list.
            commentsPart.CommentList = new CommentList();
        }

        // Get the new comment ID.
        uint commentIdx = author.LastIndex is null ? 1 : author.LastIndex + 1;
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
        authorList.Save();

        // Save the comments part.
        commentsPart.CommentList.Save();
    }
}

// Get the slide part of the first slide in the presentation document.
static SlidePart GetFirstSlide(PresentationDocument? presentationDocument)
{
    // Get relationship ID of the first slide
    PresentationPart? part = presentationDocument?.PresentationPart;
    SlideId? slideId = part?.Presentation?.SlideIdList?.GetFirstChild<SlideId>();
    string? relId = slideId?.RelationshipId;
    if (relId is null)
    {
        throw new NullReferenceException("The first slide does not contain a relationship ID.");
    }
    // Get the slide part by the relationship ID.
    SlidePart? slidePart = part?.GetPartById(relId) as SlidePart;

    if (slidePart is null)
    {
        throw new NullReferenceException("The slide part is null.");
    }

    return slidePart;
}
