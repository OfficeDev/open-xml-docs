using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;

// <Snippet0>
// <Snippet2>
if (args is [{ } fileName, { } includeHidden])
{
    RetrieveNumberOfSlides(fileName, includeHidden);
}
else if (args is [{ } fileName2])
{
    RetrieveNumberOfSlides(fileName2);
}
// </Snippet2>

// <Snippet1>
static int RetrieveNumberOfSlides(string fileName, string includeHidden = "true")
// </Snippet1>
{
    int slidesCount = 0;
    // <Snippet3>
    using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
    {
        if (doc.PresentationPart is not null)
        {
            // Get the presentation part of the document.
            PresentationPart presentationPart = doc.PresentationPart;
            // </Snippet3>

            if (presentationPart is not null)
            {
                // <Snippet4>
                if (includeHidden.ToUpper() == "TRUE")
                {
                    slidesCount = presentationPart.SlideParts.Count();
                }
                else
                {
                    // </Snippet4>
                    // <Snippet5>
                    // Each slide can include a Show property, which if hidden 
                    // will contain the value "0". The Show property may not 
                    // exist, and most likely will not, for non-hidden slides.
                    var slides = presentationPart.SlideParts.Where(
                        (s) => (s.Slide is not null) &&
                          ((s.Slide.Show is null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));

                    slidesCount = slides.Count();
                    // </Snippet5>
                }
            }
        }
    }

    Console.WriteLine($"Slide Count: {slidesCount}");

    return slidesCount;
}
// </Snippet0>
