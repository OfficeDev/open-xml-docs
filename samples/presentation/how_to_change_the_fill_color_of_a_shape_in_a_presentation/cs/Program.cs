
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

// Change the fill color of a shape.
// The test file must have a filled shape as the first shape on the first slide.
static void SetPPTShapeColor(string docName)
{
    using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
    {
        // Get the relationship ID of the first slide.
        PresentationPart part = ppt.PresentationPart;
        OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
        string relId = (slideIds[0] as SlideId).RelationshipId;

        // Get the slide part from the relationship ID.
        SlidePart slide = (SlidePart)part.GetPartById(relId);

        if (slide != null)
        {
            // Get the shape tree that contains the shape to change.
            ShapeTree tree = slide.Slide.CommonSlideData.ShapeTree;

            // Get the first shape in the shape tree.
            Shape shape = tree.GetFirstChild<Shape>();

            if (shape != null)
            {
                // Get the style of the shape.
                ShapeStyle style = shape.ShapeStyle;

                // Get the fill reference.
                Drawing.FillReference fillRef = style.FillReference;

                // Set the fill color to SchemeColor Accent 6;
                fillRef.SchemeColor = new Drawing.SchemeColor();
                fillRef.SchemeColor.Val = Drawing.SchemeColorValues.Accent6;

                // Save the modified slide.
                slide.Slide.Save();
            }
        }
    }
}