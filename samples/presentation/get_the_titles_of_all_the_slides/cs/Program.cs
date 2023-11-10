
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

GetSlideTitles(args[0]);

// Get a list of the titles of all the slides in the presentation.
static IList<string>? GetSlideTitles(string presentationFile)
{
    // Open the presentation as read-only.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        return GetSlideTitlesFromPresentation(presentationDocument);
    }
}

// Get a list of the titles of all the slides in the presentation.
static IList<string>? GetSlideTitlesFromPresentation(PresentationDocument presentationDocument)
{
    // Get a PresentationPart object from the PresentationDocument object.
    PresentationPart? presentationPart = presentationDocument.PresentationPart;

    if (presentationPart is not null && presentationPart.Presentation is not null)
    {
        // Get a Presentation object from the PresentationPart object.
        Presentation presentation = presentationPart.Presentation;

        if (presentation.SlideIdList is not null)
        {
            List<string> titlesList = new List<string>();

            // Get the title of each slide in the slide order.
            foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
            {
                if (slideId.RelationshipId is null)
                {
                    continue;
                }

                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);

                // Get the slide title.
                string title = GetSlideTitle(slidePart);

                // An empty title can also be added.
                titlesList.Add(title);
            }

            return titlesList;
        }

    }

    return null;
}

// Get the title string of the slide.
static string GetSlideTitle(SlidePart slidePart)
{
    if (slidePart is null)
    {
        throw new ArgumentNullException("presentationDocument");
    }

    // Declare a paragraph separator.
    string? paragraphSeparator = null;

    if (slidePart.Slide is not null)
    {
        // Find all the title shapes.
        var shapes = from shape in slidePart.Slide.Descendants<Shape>()
                     where IsTitleShape(shape)
                     select shape;

        StringBuilder paragraphText = new StringBuilder();

        foreach (var shape in shapes)
        {
            var paragraphs = shape.TextBody?.Descendants<D.Paragraph>();
            if (paragraphs is null)
            {
                continue;
            }

            // Get the text in each paragraph in this shape.
            foreach (var paragraph in paragraphs)
            {
                // Add a line break.
                paragraphText.Append(paragraphSeparator);

                foreach (var text in paragraph.Descendants<D.Text>())
                {
                    paragraphText.Append(text.Text);
                }

                paragraphSeparator = "\n";
            }
        }

        return paragraphText.ToString();
    }

    return string.Empty;
}

// Determines whether the shape is a title shape.
static bool IsTitleShape(Shape shape)
{
    PlaceholderShape? placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>();

    if (placeholderShape is not null && placeholderShape.Type is not null && placeholderShape.Type.HasValue)
    {
        switch ((PlaceholderValues)placeholderShape.Type)
        {
            // Any title shape.
            case PlaceholderValues.Title:

            // A centered title.
            case PlaceholderValues.CenteredTitle:
                return true;

            default:
                return false;
        }
    }

    return false;
}
