
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using Drawing = DocumentFormat.OpenXml.Drawing;
// <Snippet3>
if (args is [{ } fileName])
{
    foreach (string link in GetAllExternalHyperlinksInPresentation(fileName))
    {
        Console.WriteLine(link);
    }
}
// </Snippet3>

// <Snippet>
// Returns all the external hyperlinks in the slides of a presentation.
static IEnumerable<String> GetAllExternalHyperlinksInPresentation(string fileName)
{
    // Declare a list of strings.
    List<string> ret = new List<string>();

    // <Snippet1>
    // Open the presentation file as read-only.
    using (PresentationDocument document = PresentationDocument.Open(fileName, false))
    // </Snippet1>
    {
        // If there is no PresentationPart then there are no hyperlinks
        if (document.PresentationPart is null)
        {
            return ret;
        }

        // <Snippet2>
        // Iterate through all the slide parts in the presentation part.
        foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
        {
            IEnumerable<Drawing.HyperlinkType> links = slidePart.Slide.Descendants<Drawing.HyperlinkType>();

            // Iterate through all the links in the slide part.
            foreach (Drawing.HyperlinkType link in links)
            {
                // Iterate through all the external relationships in the slide part. 
                foreach (HyperlinkRelationship relation in slidePart.HyperlinkRelationships)
                {
                    // If the relationship ID matches the link ID…
                    if (relation.Id.Equals(link.Id))
                    {
                        // Add the URI of the external relationship to the list of strings.
                        ret.Add(relation.Uri.AbsoluteUri);
                    }
                }
            }
        }
        // </Snippet2>
    }

    // Return the list of strings.
    return ret;
}
// </Snippet>