
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;

GetSlideIdAndText(out string text, args[0], int.Parse(args[1]));

static void GetSlideIdAndText(out string sldText, string docName, int index)
{
    using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
    {
        // Get the relationship ID of the first slide.
        PresentationPart? part = ppt.PresentationPart;
        OpenXmlElementList slideIds = part?.Presentation?.SlideIdList?.ChildElements ?? default;

        // If there are no slide IDs then there are no slides.
        if (slideIds.Count() < 1)
        {
            sldText = "";
            return;
        }

        string? relId = (slideIds[index] as SlideId)?.RelationshipId;

        if (relId is null)
        {
            sldText = "";
            return;
        }

        // Get the slide part from the relationship ID.
        SlidePart slide = (SlidePart)part!.GetPartById(relId);

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
