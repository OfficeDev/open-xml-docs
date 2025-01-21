using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

SetPPTShapeColor(args[0]);

// <Snippet0>
// Change the fill color of a shape.
// The test file must have a shape on the first slide.
static void SetPPTShapeColor(string docName)
{
    // <Snippet1>
    using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
    // </Snippet1>
    {
        // <Snippet2>
        // Get the relationship ID of the first slide.
        PresentationPart presentationPart = ppt.PresentationPart ?? ppt.AddPresentationPart();
        presentationPart.Presentation.SlideIdList ??= new SlideIdList();
        SlideId? slideId = presentationPart.Presentation.SlideIdList.GetFirstChild<SlideId>();

        if (slideId is not null)
        {
            string? relId = slideId.RelationshipId;

            if (relId is not null)
            {
                // Get the slide part from the relationship ID.
                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relId);
                // </Snippet2>
                // <Snippet3>
                // Get or add the shape tree
                slidePart.Slide.CommonSlideData ??= new CommonSlideData();

                // Get the shape tree that contains the shape to change.
                slidePart.Slide.CommonSlideData.ShapeTree ??= new ShapeTree();

                // Get the first shape in the shape tree.
                Shape? shape = slidePart.Slide.CommonSlideData.ShapeTree.GetFirstChild<Shape>();

                if (shape is not null)
                {
                    // Get or add the shape properties element of the shape.
                    shape.ShapeProperties ??= new ShapeProperties();

                    // Get or add the fill reference.
                    Drawing.SolidFill? solidFill = shape.ShapeProperties.GetFirstChild<Drawing.SolidFill>();

                    if (solidFill is null)
                    {
                        shape.ShapeProperties.AddChild(new Drawing.SolidFill());
                        solidFill = shape.ShapeProperties.GetFirstChild<Drawing.SolidFill>();
                    }

                    // Set the fill color to SchemeColor
                    solidFill!.SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Accent2 };
                }
                // </Snippet3>
            }
        }
    }
}
// </Snippet0>
