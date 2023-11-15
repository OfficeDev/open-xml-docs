using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

SetPPTShapeColor(args[0]);

// Change the fill color of a shape.
// The test file must have a filled shape as the first shape on the first slide.
static void SetPPTShapeColor(string docName)
{
    using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
    {
        // Get the relationship ID of the first slide.
        PresentationPart presentationPart = ppt.PresentationPart ?? ppt.AddPresentationPart();
        SlideIdList slideIdList = presentationPart.Presentation.SlideIdList ?? presentationPart.Presentation.AppendChild(new SlideIdList());
        SlideId? slideId = slideIdList.GetFirstChild<SlideId>();

        if (slideId is not null)
        {
            string? relId = slideId.RelationshipId;

            if (relId is not null)
            {
                // Get the slide part from the relationship ID.
                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relId);

                if (slidePart is not null && slidePart.Slide is not null && slidePart.Slide.CommonSlideData is not null && slidePart.Slide.CommonSlideData.ShapeTree is not null)
                {

                    // Get the shape tree that contains the shape to change.
                    ShapeTree tree = slidePart.Slide.CommonSlideData.ShapeTree;

                    // Get the first shape in the shape tree.
                    Shape? shape = tree.GetFirstChild<Shape>();

                    if (shape is not null && shape.ShapeStyle is not null && shape.ShapeStyle.FillReference is not null)
                    {
                        // Get the style of the shape.
                        ShapeStyle style = shape.ShapeStyle;

                        // Get the fill reference.
                        Drawing.FillReference fillRef = style.FillReference;

                        // Set the fill color to SchemeColor Accent 6;
                        fillRef.SchemeColor = new Drawing.SchemeColor();
                        fillRef.SchemeColor.Val = Drawing.SchemeColorValues.Accent6;

                        // Save the modified slide.
                        slidePart.Slide.Save();
                    }
                }
            }
        }
    }
}