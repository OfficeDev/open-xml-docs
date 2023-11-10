
using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Text;

if (args is [{ } sldText, { } slideIndex])
{
    GetSlideIdAndText(out string text, sldText, int.Parse(slideIndex));
}

if (args is [{ } presentationFile])
{
    CountSlides(presentationFile);
}

static int CountSlides(string presentationFile)
{
    // Open the presentation as read-only.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        // Pass the presentation to the next CountSlides method
        // and return the slide count.
        return CountSlidesFromPresentation(presentationDocument);
    }
}

// Count the slides in the presentation.
static int CountSlidesFromPresentation(PresentationDocument presentationDocument)
{
    // Check for a null document object.
    if (presentationDocument is null)
    {
        throw new ArgumentNullException("presentationDocument");
    }

    int slidesCount = 0;

    // Get the presentation part of document.
    PresentationPart? presentationPart = presentationDocument.PresentationPart;
    // Get the slide count from the SlideParts.
    if (presentationPart is not null)
    {
        slidesCount = presentationPart.SlideParts.Count();
    }

    // Return the slide count to the previous method.
    return slidesCount;
}

static void GetSlideIdAndText(out string sldText, string docName, int index)
{
    using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
    {
        // Get the relationship ID of the first slide.
        PresentationPart? part = ppt.PresentationPart;
        OpenXmlElementList? slideIds = part?.Presentation?.SlideIdList?.ChildElements;

        if (part is null || slideIds is null || slideIds.Count == 0)
        {
            sldText = "";
            return;
        }

        string? relId = ((SlideId)slideIds[index]).RelationshipId;

        if (relId is null)
        {
            sldText = "";
            return;
        }

        // Get the slide part from the relationship ID.
        SlidePart slide = (SlidePart)part.GetPartById(relId);

        // Build a StringBuilder object.
        StringBuilder paragraphText = new StringBuilder();

        // Get the inner text of the slide:
        IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
        foreach (A.Text text in texts)
        {
            paragraphText.Append(text.Text);
        }
        sldText = paragraphText.ToString();
    }
}
