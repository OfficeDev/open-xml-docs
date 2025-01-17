
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Linq;
using static SlideMover;

bool fromIsValid = int.TryParse(args[1], out int from);
bool toIsValid = int.TryParse(args[2], out int to);

if (fromIsValid && toIsValid)
{
    MoveSlide(args[0], from, to);
}

public class SlideMover
{
    //<Snippet0>
    // Counting the slides in the presentation.
    public static int CountSlides(string presentationFile)
    // <Snippet1>
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
        // </Snippet1>
        {
            // <Snippet2>
            // Pass the presentation to the next CountSlides method
            // and return the slide count.
            return CountSlides(presentationDocument);
            // </Snippet2>
        }
    }

    // Count the slides in the presentation.
    static int CountSlides(PresentationDocument presentationDocument)
    {
        // <Snippet3>
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
        // </Snippet3>
    }

    // <Snippet4>
    // Move a slide to a different position in the slide order in the presentation.
    public static void MoveSlide(string presentationFile, int from, int to)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            MoveSlide(presentationDocument, from, to);
        }
    }
    // </Snippet4>

    // <Snippet5>
    // Move a slide to a different position in the slide order in the presentation.
    static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
    {
        if (presentationDocument is null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        // Call the CountSlides method to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        // Verify that both from and to positions are within range and different from one another.
        if (from < 0 || from >= slidesCount)
        {
            throw new ArgumentOutOfRangeException("from");
        }

        if (to < 0 || from >= slidesCount || to == from)
        {
            throw new ArgumentOutOfRangeException("to");
        }
        // </Snippet5>

        // <Snippet6>
        // Get the presentation part from the presentation document.
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        // The slide count is not zero, so the presentation must contain slides.            
        Presentation? presentation = presentationPart?.Presentation;

        if (presentation is null)
        {
            throw new ArgumentNullException(nameof(presentation));
        }

        SlideIdList? slideIdList = presentation.SlideIdList;

        if (slideIdList is null)
        {
            throw new ArgumentNullException(nameof(slideIdList));
        }

        // Get the slide ID of the source slide.
        SlideId? sourceSlide = slideIdList.ChildElements[from] as SlideId;

        if (sourceSlide is null)
        {
            throw new ArgumentNullException(nameof(sourceSlide));
        }

        SlideId? targetSlide = null;

        // Identify the position of the target slide after which to move the source slide.
        if (to == 0)
        {
            targetSlide = null;
        }
        else if (from < to)
        {
            targetSlide = slideIdList.ChildElements[to] as SlideId;
        }
        else
        {
            targetSlide = slideIdList.ChildElements[to - 1] as SlideId;
        }
        // </Snippet6>
        // <Snippet7>
        // Remove the source slide from its current position.
        sourceSlide.Remove();

        // Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide);
        // </Snippet7>
    }
    // </Snippet0>

}