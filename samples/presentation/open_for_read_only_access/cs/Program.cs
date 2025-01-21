
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;

// <Snippet4>
try
{
    string file = args[0];
    bool isInt = int.TryParse(args[1], out int i);

    if (isInt)
    {
        GetSlideIdAndText(out string sldText, file, i);
        Console.WriteLine($"The text in slide #{i + 1} is {sldText}");
    }
}
catch(ArgumentOutOfRangeException exp) {
    Console.Error.WriteLine(exp.Message);
}
// </Snippet4>

// <Snippet0>
static void GetSlideIdAndText(out string sldText, string docName, int index)
{
    // <Snippet1>
    using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
    {
        // </Snippet1>
        // <Snippet2>
        // Get the relationship ID of the first slide.
        PresentationPart? part = ppt.PresentationPart;
        OpenXmlElementList slideIds = part?.Presentation?.SlideIdList?.ChildElements ?? default;

        // If there are no slide IDs then there are no slides.
        if (slideIds.Count == 0)
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
        // </Snippet2>

        // <Snippet3>
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
        // </Snippet3>
    }
}
// </Snippet0>
