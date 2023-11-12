Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation

Module MyModule
    Sub CreatePresentation(ByVal filepath As String)

        ' Create a presentation at a specified file path. The presentation document type is pptx, by default.
        Dim presentationDoc As PresentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)
        Dim presentationPart As PresentationPart = presentationDoc.AddPresentationPart()
        presentationPart.Presentation = New Presentation()

        CreatePresentationPartsFromPresentation(presentationPart)

        ' Dispose the presentation handle.
        presentationDoc.Dispose()
    End Sub
    Sub CreatePresentationPartsFromPresentation(ByVal presentationPart As PresentationPart)
        Dim slideMasterIdList1 As New SlideMasterIdList(New SlideMasterId() With {
         .Id = 2147483648UI,
         .RelationshipId = "rId1"
        })
        Dim slideIdList1 As New SlideIdList(New SlideId() With {
         .Id = 256UI,
         .RelationshipId = "rId2"
        })
        Dim slideSize1 As New SlideSize() With {
         .Cx = 9144000,
         .Cy = 6858000,
         .Type = SlideSizeValues.Screen4x3
        }
        Dim notesSize1 As New NotesSize() With {
         .Cx = 6858000,
         .Cy = 9144000
        }
        Dim defaultTextStyle1 As New DefaultTextStyle()

        presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1)

        ' Code to create other parts of the presentation file goes here.
    End Sub
End Module
