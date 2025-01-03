using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2016.Presentation.Command;
using DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Linq;

AddCommentToPresentation(args[0], args[1], args[2], string.Join(' ', args[3..]));
// <Snippet0>
static void AddCommentToPresentation(string file, string initials, string name, string text)
{
    using (PresentationDocument presentationDocument = PresentationDocument.Open(file, true))
    {
        PresentationPart presentationPart = presentationDocument?.PresentationPart ?? throw new MissingFieldException("PresentationPart");

        // <Snippet1>
        // create missing PowerPointAuthorsPart if it is null
        if (presentationDocument.PresentationPart.authorsPart is null)
        {
            presentationDocument.PresentationPart.AddNewPart<PowerPointAuthorsPart>();
        }
        // </Snippet1>

        // <Snippet2>
        // Add missing AuthorList if it is null
        if (presentationDocument.PresentationPart.authorsPart!.AuthorList is null)
        {
            presentationDocument.PresentationPart.authorsPart.AuthorList = new AuthorList();
        }

        // Get the author or create a new one
        Author? author = presentationDocument.PresentationPart.authorsPart.AuthorList
            .ChildElements.OfType<Author>().Where(a => a.Name?.Value == name).FirstOrDefault();

        if (author is null)
        {
            string authorId = string.Concat("{", Guid.NewGuid(), "}");
            string userId = string.Concat(name.Split(" ").FirstOrDefault() ?? "user", "@example.com::", Guid.NewGuid());
            author = new Author() { Id = authorId, Name = name, Initials = initials, UserId = userId, ProviderId = string.Empty };

            presentationDocument.PresentationPart.authorsPart.AuthorList.AppendChild(author);
        }
        // </Snippet2>

        // <Snippet3>
        // Get the Id of the slide to add the comment to
        SlideId? slideId = presentationDocument.PresentationPart.Presentation.SlideIdList?.Elements<SlideId>()?.FirstOrDefault();
        
        // If slideId is null, there are no slides, so return
        if (slideId is null) return;
        // </Snippet3>
        Random ran = new Random();
        UInt32Value cid = Convert.ToUInt32(ran.Next(100000000, 999999999));

        // <Snippet4>
        // Get the relationship id of the slide if it exists
        string? relId = slideId.RelationshipId;

        // Use the relId to get the slide if it exists, otherwise take the first slide in the sequence
        SlidePart slidePart = relId is not null ? (SlidePart)presentationPart.GetPartById(relId) 
            : presentationDocument.PresentationPart.SlideParts.First();

        // If the slide part has comments parts take the first PowerPointCommentsPart
        // otherwise add a new one
        PowerPointCommentPart powerPointCommentPart = slidePart.commentParts.FirstOrDefault() ?? slidePart.AddNewPart<PowerPointCommentPart>();
        // </Snippet4>

        // <Snippet5>
        // Create the comment using the new modern comment class DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment
        DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment comment = new DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment(
                new SlideMonikerList(
                    new DocumentMoniker(),
                    new SlideMoniker()
                    {
                        CId = cid,
                        SldId = slideId.Id,
                    }),
                new TextBodyType(
                    new BodyProperties(),
                    new ListStyle(),
                    new Paragraph(new Run(new DocumentFormat.OpenXml.Drawing.Text(text)))))
        {
            Id = string.Concat("{", Guid.NewGuid(), "}"),
            AuthorId = author.Id,
            Created = DateTime.Now,
        };

        // If the comment list does not exist, add one.
        powerPointCommentPart.CommentList ??= new DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList();
        // Add the comment to the comment list
        powerPointCommentPart.CommentList.AppendChild(comment);
        // </Snippet5>
        
        // <Snippet6>
        // Get the presentation extension list if it exists
        SlideExtensionList? presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().FirstOrDefault();
        // Create a boolean that determines if this is the slide's first comment
        bool isFirstComment = false;

        // If the presentation extension list is null, add one and set this as the first comment for the slide
        if (presentationExtensionList is null)
        {
            isFirstComment = true;
            slidePart.Slide.AppendChild(new SlideExtensionList());
            presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().First();
        }

        // Get the slide extension if it exists
        SlideExtension? presentationExtension = presentationExtensionList.ChildElements.OfType<SlideExtension>().FirstOrDefault();

        // If the slide extension is null, add it and set this as a new comment
        if (presentationExtension is null)
        {
            isFirstComment = true;
            presentationExtensionList.AddChild(new SlideExtension() { Uri = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}" });
            presentationExtension = presentationExtensionList.ChildElements.OfType<SlideExtension>().First();
        }

        // If this is the first comment for the slide add the comment relationship
        if (isFirstComment)
        {
            presentationExtension.AddChild(new CommentRelationship()
            { Id = slidePart.GetIdOfPart(powerPointCommentPart) });
        }
        // </Snippet6>
    }
}
// </Snippet0>