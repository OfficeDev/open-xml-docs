using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

CreatePresentation(args[0]);

static void CreatePresentation(string filepath)
{
    // Create a presentation at a specified file path. The presentation document type is pptx, by default.
    using (PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation))
    {
        PresentationPart presentationPart = presentationDoc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        CreatePresentationParts(presentationPart);
    }
}
static void CreatePresentationParts(PresentationPart presentationPart)
{
    SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
    SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
    SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
    NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
    DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

    presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

    // Code to create other parts of the presentation file goes here.
}
