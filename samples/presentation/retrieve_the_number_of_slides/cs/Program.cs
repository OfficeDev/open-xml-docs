using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO.Enumeration;
using System.Linq;

if (args is [ { } fileName, {} includeHidden])
{
    RetrieveNumberOfSlides(fileName, includeHidden);
}
else if (args is [{ } fileName2])
{
    RetrieveNumberOfSlides(fileName2);
}

static int RetrieveNumberOfSlides(string fileName, string includeHidden = "true")
{
    int slidesCount = 0;

    using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
    {
        if (doc is not null && doc.PresentationPart is not null)
        {
            // Get the presentation part of the document.
            PresentationPart presentationPart = doc.PresentationPart;
            if (presentationPart is not null)
            {
                if (includeHidden.ToLower() == "true")
                {
                    slidesCount = presentationPart.SlideParts.Count();
                }
                else
                {
                    // Each slide can include a Show property, which if hidden 
                    // will contain the value "0". The Show property may not 
                    // exist, and most likely will not, for non-hidden slides.
                    var slides = presentationPart.SlideParts.Where(
                        (s) => (s.Slide is not null) &&
                          ((s.Slide.Show is null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));

                    slidesCount = slides.Count();
                }
            }
        }
    }

    Console.WriteLine($"Slide Count: {slidesCount}");

    return slidesCount;
}