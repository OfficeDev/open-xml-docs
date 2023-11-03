
    Public Shared Sub CreatePresentation(ByVal filepath As String)

                ' Create a presentation at a specified file path. The presentation document type is pptx, by default.
                Dim presentationDoc As PresentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)
                Dim presentationPart As PresentationPart = presentationDoc.AddPresentationPart()
                presentationPart.Presentation = New Presentation()

                CreatePresentationParts(presentationPart)

                ' Close the presentation handle.
                presentationDoc.Close()
            End Sub
    Private Shared Sub CreatePresentationParts(ByVal presentationPart As PresentationPart)
                Dim slideMasterIdList1 As New SlideMasterIdList(New SlideMasterId() With { _
                 .Id = DirectCast(2147483648UI, UInt32Value), _
                 .RelationshipId = "rId1" _
                })
                Dim slideIdList1 As New SlideIdList(New SlideId() With { _
                 .Id = DirectCast(256UI, UInt32Value), _
                 .RelationshipId = "rId2" _
                })
                Dim slideSize1 As New SlideSize() With { _
                 .Cx = 9144000, _
                 .Cy = 6858000, _
                 .Type = SlideSizeValues.Screen4x3 _
                }
                Dim notesSize1 As New NotesSize() With { _
                 .Cx = 6858000, _
                 .Cy = 9144000 _
                }
                Dim defaultTextStyle1 As New DefaultTextStyle()

                presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1)

             ' Code to create other parts of the presentation file goes here.
            End Sub


    Public Shared Sub CreatePresentation(ByVal filepath As String)

                ' Create a presentation at a specified file path. The presentation document type is pptx, by default.
                Dim presentationDoc As PresentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)
                Dim presentationPart As PresentationPart = presentationDoc.AddPresentationPart()
                presentationPart.Presentation = New Presentation()

                CreatePresentationParts(presentationPart)

                ' Close the presentation handle.
                presentationDoc.Close()
            End Sub
    Private Shared Sub CreatePresentationParts(ByVal presentationPart As PresentationPart)
                Dim slideMasterIdList1 As New SlideMasterIdList(New SlideMasterId() With { _
                 .Id = DirectCast(2147483648UI, UInt32Value), _
                 .RelationshipId = "rId1" _
                })
                Dim slideIdList1 As New SlideIdList(New SlideId() With { _
                 .Id = DirectCast(256UI, UInt32Value), _
                 .RelationshipId = "rId2" _
                })
                Dim slideSize1 As New SlideSize() With { _
                 .Cx = 9144000, _
                 .Cy = 6858000, _
                 .Type = SlideSizeValues.Screen4x3 _
                }
                Dim notesSize1 As New NotesSize() With { _
                 .Cx = 6858000, _
                 .Cy = 9144000 _
                }
                Dim defaultTextStyle1 As New DefaultTextStyle()

                presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1)

             ' Code to create other parts of the presentation file goes here.
            End Sub
