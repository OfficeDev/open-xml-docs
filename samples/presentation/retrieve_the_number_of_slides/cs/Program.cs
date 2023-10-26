#nullable disable

using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;

static int RetrieveNumberOfSlides(string fileName,
    bool includeHidden = true)
{
    int slidesCount = 0;

    using (PresentationDocument doc =
        PresentationDocument.Open(fileName, false))
    {
        // Get the presentation part of the document.
        PresentationPart presentationPart = doc.PresentationPart;
        if (presentationPart != null)
        {
            if (includeHidden)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }
            else
            {
                // Each slide can include a Show property, which if hidden 
                // will contain the value "0". The Show property may not 
                // exist, and most likely will not, for non-hidden slides.
                var slides = presentationPart.SlideParts.Where(
                    (s) => (s.Slide != null) &&
                      ((s.Slide.Show == null) || (s.Slide.Show.HasValue &&
                      s.Slide.Show.Value)));
                slidesCount = slides.Count();
            }
        }
    }
    return slidesCount;
}