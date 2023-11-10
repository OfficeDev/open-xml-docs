
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using System;

MoveParagraphToPresentation(args[0], args[1]);

// Moves a paragraph range in a TextBody shape in the source document
// to another TextBody shape in the target document.
static void MoveParagraphToPresentation(string sourceFile, string targetFile)
{
    // Open the source file as read/write.
    using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, true))
    // Open the target file as read/write.
    using (PresentationDocument targetDoc = PresentationDocument.Open(targetFile, true))
    {
        // Get the first slide in the source presentation.
        SlidePart slide1 = GetFirstSlide(sourceDoc);

        // Get the first TextBody shape in it.
        TextBody textBody1 = slide1.Slide.Descendants<TextBody>().First();

        // Get the first paragraph in the TextBody shape.
        // Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
        Drawing.Paragraph p1 = textBody1.Elements<Drawing.Paragraph>().First();

        // Get the first slide in the target presentation.
        SlidePart slide2 = GetFirstSlide(targetDoc);

        // Get the first TextBody shape in it.
        TextBody textBody2 = slide2.Slide.Descendants<TextBody>().First();

        // Clone the source paragraph and insert the cloned. paragraph into the target TextBody shape.
        // Passing "true" creates a deep clone, which creates a copy of the 
        // Paragraph object and everything directly or indirectly referenced by that object.
        textBody2.Append(p1.CloneNode(true));

        // Remove the source paragraph from the source file.
        textBody1.RemoveChild<Drawing.Paragraph>(p1);

        // Replace the removed paragraph with a placeholder.
        textBody1.AppendChild<Drawing.Paragraph>(new Drawing.Paragraph());

        // Save the slide in the source file.
        slide1.Slide.Save();

        // Save the slide in the target file.
        slide2.Slide.Save();
    }
}

// Get the slide part of the first slide in the presentation document.
static SlidePart GetFirstSlide(PresentationDocument presentationDocument)
{
    // Get relationship ID of the first slide
    PresentationPart part = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();
    SlideIdList slideIdList = part.Presentation.SlideIdList ?? part.Presentation.AppendChild(new SlideIdList());
    SlideId slideId = part.Presentation.SlideIdList?.GetFirstChild<SlideId>() ?? slideIdList.AppendChild<SlideId>(new SlideId());
    string? relId = slideId.RelationshipId;

    if (relId is null)
    {
        throw new ArgumentNullException(nameof(relId));
    }

    // Get the slide part by the relationship ID.
    SlidePart slidePart = (SlidePart)part.GetPartById(relId);

    return slidePart;
}
