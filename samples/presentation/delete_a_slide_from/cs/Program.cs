using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;

public class Program
{
    public static void Main(string[] args)
    {
        int count = CountSlides(args[0]);

        Console.WriteLine($"{count} slides found");

        DeleteSlide(args[0], 0);
    }

    // <Snippet0>
    // Get the presentation object and pass it to the next CountSlides method.
    static int CountSlides(string presentationFile)
    {
        // <Snippet1>
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            //</Snippet1>
        {
            // <Snippet3>
            // Pass the presentation to the next CountSlide method
            // and return the slide count.
            return CountSlides(presentationDocument);
            // </Snippet3>
        }
    }

    // Count the slides in the presentation.
    static int CountSlides(PresentationDocument presentationDocument)
    {
        // <Snippet4>
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
        // </Snippet4>
    }
    
    // <Snippet5>
    // Get the presentation object and pass it to the next DeleteSlide method.
    static void DeleteSlide(string presentationFile, int slideIndex)
    {
        // <Snippet2>
        // Open the source document as read/write.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            // </Snippet2>
        {
            // Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
            DeleteSlide(presentationDocument, slideIndex);
        }
    }
    // </Snippet5>
    // <Snippet6>
    // Delete the specified slide from the presentation.
    static void DeleteSlide(PresentationDocument presentationDocument, int slideIndex)
    {
        if (presentationDocument is null)
        {
            throw new ArgumentNullException(nameof(presentationDocument));
        }

        // Use the CountSlides sample to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

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
        // </Snippet6>

        // <Snippet7>
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
        // </Snippet7>

        // <Snippet8>
        // Get the slide part for the specified slide.
        SlidePart slidePart = (SlidePart)presentationPart!.GetPartById(slideRelId);

        // Remove the slide part.
        presentationPart.DeletePart(slidePart);
        // </Snippet8>
    }
    // </Snippet0>
}
