
using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

ApplyThemeToPresentation(args[0], args[1]);

// Apply a new theme to the presentation. 
static void ApplyThemeToPresentation(string presentationFile, string themePresentation)
{
    using (PresentationDocument themeDocument = PresentationDocument.Open(themePresentation, false))
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        ApplyThemeToPresentationDocument(presentationDocument, themeDocument);
    }
}

// Apply a new theme to the presentation. 
static void ApplyThemeToPresentationDocument(PresentationDocument presentationDocument, PresentationDocument themeDocument)
{
    // Get the presentation part of the presentation document.
    PresentationPart presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();

    // Get the existing slide master part.
    SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.ElementAt(0);


    string relationshipId = presentationPart.GetIdOfPart(slideMasterPart);

    // Get the new slide master part.
    PresentationPart themeDocPresentationPart = themeDocument.PresentationPart ?? themeDocument.AddPresentationPart();
    SlideMasterPart newSlideMasterPart = themeDocPresentationPart.SlideMasterParts.ElementAt(0);

    if (presentationPart.ThemePart is not null)
    {
        // Remove the existing theme part.
        presentationPart.DeletePart(presentationPart.ThemePart);
    }

    // Remove the old slide master part.
    presentationPart.DeletePart(slideMasterPart);

    // Import the new slide master part, and reuse the old relationship ID.
    newSlideMasterPart = presentationPart.AddPart(newSlideMasterPart, relationshipId);

    if (newSlideMasterPart.ThemePart is not null)
    {
        // Change to the new theme part.
        presentationPart.AddPart(newSlideMasterPart.ThemePart);
    }

    Dictionary<string, SlideLayoutPart> newSlideLayouts = new Dictionary<string, SlideLayoutPart>();

    foreach (var slideLayoutPart in newSlideMasterPart.SlideLayoutParts)
    {
        string? newLayoutType = GetSlideLayoutType(slideLayoutPart);

        if (newLayoutType is not null)
        {
            newSlideLayouts.Add(newLayoutType, slideLayoutPart);
        }
    }

    string? layoutType = null;
    SlideLayoutPart? newLayoutPart = null;

    // Insert the code for the layout for this example.
    string defaultLayoutType = "Title and Content";

    // Remove the slide layout relationship on all slides. 
    foreach (var slidePart in presentationPart.SlideParts)
    {
        layoutType = null;

        if (slidePart.SlideLayoutPart is not null)
        {
            // Determine the slide layout type for each slide.
            layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart);

            // Delete the old layout part.
            slidePart.DeletePart(slidePart.SlideLayoutPart);
        }

        if (layoutType is not null && newSlideLayouts.TryGetValue(layoutType, out newLayoutPart))
        {
            // Apply the new layout part.
            slidePart.AddPart(newLayoutPart);
        }
        else
        {
            newLayoutPart = newSlideLayouts[defaultLayoutType];

            // Apply the new default layout part.
            slidePart.AddPart(newLayoutPart);
        }
    }
}

// Get the slide layout type.
static string? GetSlideLayoutType(SlideLayoutPart slideLayoutPart)
{
    CommonSlideData? slideData = slideLayoutPart.SlideLayout?.CommonSlideData;

    // Remarks: If this is used in production code, check for a null reference.

    return slideData?.Name;
}
