
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

GetAllTextInSlide(args[0], int.Parse(args[1]));

// Get all the text in a slide.
static string[]? GetAllTextInSlide(string presentationFile, int slideIndex)
{
    // Open the presentation as read-only.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        // Pass the presentation and the slide index
        // to the next GetAllTextInSlide method, and
        // then return the array of strings it returns. 
        return GetAllTextInSlideFromPresentation(presentationDocument, slideIndex);
    }
}
static string[]? GetAllTextInSlideFromPresentation(PresentationDocument presentationDocument, int slideIndex)
{
    // Verify that the slide index is not out of range.
    if (slideIndex < 0)
    {
        throw new ArgumentOutOfRangeException("slideIndex");
    }

    // Get the presentation part of the presentation document.
    PresentationPart? presentationPart = presentationDocument.PresentationPart;

    // Verify that the presentation part and presentation exist.
    if (presentationPart is not null && presentationPart.Presentation is not null)
    {
        // Get the Presentation object from the presentation part.
        Presentation presentation = presentationPart.Presentation;

        // Verify that the slide ID list exists.
        if (presentation.SlideIdList is not null)
        {
            // Get the collection of slide IDs from the slide ID list.
            DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;

            // If the slide ID is in range...
            if (slideIndex < slideIds.Count)
            {
                // Get the relationship ID of the slide.
                string? slidePartRelationshipId = ((SlideId)slideIds[slideIndex]).RelationshipId;

                if (slidePartRelationshipId is null)
                {
                    return null;
                }

                // Get the specified slide part from the relationship ID.
                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                // Pass the slide part to the next method, and
                // then return the array of strings that method
                // returns to the previous method.
                return GetAllTextInSlideFromPart(slidePart);
            }
        }
    }

    // Else, return null.
    return null;
}
static string[] GetAllTextInSlideFromPart(SlidePart slidePart)
{
    // Verify that the slide part exists.
    if (slidePart is null)
    {
        throw new ArgumentNullException("slidePart");
    }

    // Create a new linked list of strings.
    LinkedList<string> texts = new LinkedList<string>();

    // If the slide exists...
    if (slidePart.Slide is not null)
    {
        // Iterate through all the paragraphs in the slide.
        foreach (DocumentFormat.OpenXml.Drawing.Paragraph paragraph in
            slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
        {
            // Create a new string builder.                    
            StringBuilder paragraphText = new StringBuilder();

            // Iterate through the lines of the paragraph.
            foreach (DocumentFormat.OpenXml.Drawing.Text text in
                paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
            {
                // Append each line to the previous lines.
                paragraphText.Append(text.Text);
            }

            if (paragraphText.Length > 0)
            {
                // Add each paragraph to the linked list.
                texts.AddLast(paragraphText.ToString());
            }
        }
    }

    // Return an array of strings.
    return texts.ToArray();
}
