
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;

CountSlides(args[0]);

// Get the presentation object and pass it to the next CountSlides method.
static int CountSlides(string presentationFile)
{
    // Open the presentation as read-only.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        // Pass the presentation to the next CountSlide method
        // and return the slide count.
        return CountSlidesFromPresentation(presentationDocument);
    }
}

// Count the slides in the presentation.
static int CountSlidesFromPresentation(PresentationDocument presentationDocument)
{
    int slidesCount = 0;

    // Get the presentation part of document.
    PresentationPart presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();

    // Get the slide count from the SlideParts.
    if (presentationPart is not null)
    {
        slidesCount = presentationPart.SlideParts.Count();
    }

    // Return the slide count to the previous method.
    return slidesCount;
}
//
// Get the presentation object and pass it to the next DeleteSlide method.
static void DeleteSlide(string presentationFile, int slideIndex)
{
    // Open the source document as read/write.

    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        // Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
        DeleteSlideFromPresentation(presentationDocument, slideIndex);
    }
}

// Delete the specified slide from the presentation.
static void DeleteSlideFromPresentation(PresentationDocument presentationDocument, int slideIndex)
{
    if (presentationDocument is null)
    {
        throw new ArgumentNullException("presentationDocument");
    }

    // Use the CountSlides sample to get the number of slides in the presentation.
    int slidesCount = CountSlidesFromPresentation(presentationDocument);

    if (slideIndex < 0 || slideIndex >= slidesCount)
    {
        throw new ArgumentOutOfRangeException("slideIndex");
    }

    // Get the presentation part from the presentation document. 
    PresentationPart? presentationPart = presentationDocument.PresentationPart;

    // Get the presentation from the presentation part.
    Presentation? presentation = presentationPart?.Presentation;

    // Get the list of slide IDs in the presentation.
    SlideIdList? slideIdList = presentation?.SlideIdList;

    // Get the slide ID of the specified slide
    SlideId? slideId = slideIdList?.ChildElements[slideIndex] as SlideId;

    // Get the relationship ID of the slide.
    string? slideRelId = slideId?.RelationshipId;

    // If there's no relationship ID, there's no slide to delete.
    if (slideRelId is null)
    {
        return;
    }

    // Remove the slide from the slide list.
    slideIdList!.RemoveChild(slideId);

    //
    // Remove references to the slide from all custom shows.
    if (presentation!.CustomShowList is not null)
    {
        // Iterate through the list of custom shows.
        foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
        {
            if (customShow.SlideList is not null)
            {
                // Declare a link list of slide list entries.
                LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                {
                    // Find the slide reference to remove from the custom show.
                    if (slideListEntry.Id is not null && slideListEntry.Id == slideRelId)
                    {
                        slideListEntries.AddLast(slideListEntry);
                    }
                }

                // Remove all references to the slide from the custom show.
                foreach (SlideListEntry slideListEntry in slideListEntries)
                {
                    customShow.SlideList.RemoveChild(slideListEntry);
                }
            }
        }
    }

    // Save the modified presentation.
    presentation.Save();

    // Get the slide part for the specified slide.
    SlidePart slidePart = (SlidePart)presentationPart!.GetPartById(slideRelId);

    // Remove the slide part.
    presentationPart.DeletePart(slidePart);
}
