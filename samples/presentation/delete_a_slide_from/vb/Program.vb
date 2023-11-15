Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation


Module MyModule

    Sub Main(args As String())
    End Sub

    ' Count the number of slides in the presentation.
    Public Function CountSlides(ByVal presentationFile As String) As Integer
        ' Open the presentation as read-only.
        Using presentationDocument__1 As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' Pass the presentation to the next CountSlides method
            ' and return the slide count.
            Return CountSlides(presentationDocument__1)
        End Using
    End Function
    ' Count the slides in the presentation.
    Public Function CountSlides(ByVal presentationDocument As PresentationDocument) As Integer
        ' Check for a null document object.
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        Dim slidesCount As Integer = 0

        ' Get the presentation part of document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        If presentationPart IsNot Nothing AndAlso presentationPart.Presentation IsNot Nothing Then
            ' Get the Presentation object from the presentation part.
            Dim presentation As Presentation = presentationPart.Presentation

            ' Verify that the presentation contains slides. 
            If presentation.SlideIdList IsNot Nothing Then

                ' Get the slide count from the slide ID list. 
                slidesCount = presentation.SlideIdList.Elements(Of SlideId)().Count()
            End If
        End If

        ' Return the slide count to the previous method.
        Return slidesCount
    End Function
    ' Delete the specified slide from the presentation.
    Public Sub DeleteSlide(ByVal presentationFile As String, ByVal slideIndex As Integer)

        ' Open the source document as read/write.
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, True)

        Using (presentationDocument)

            ' Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
            DeleteSlide2(presentationDocument, slideIndex)

        End Using

    End Sub
    ' Delete the specified slide in the presentation.
    Public Sub DeleteSlide2(ByVal presentationDocument As PresentationDocument, ByVal slideIndex As Integer)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Use the CountSlides code example to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)
        If ((slideIndex < 0) OrElse (slideIndex >= slidesCount)) Then
            Throw New ArgumentOutOfRangeException("slideIndex")
        End If

        ' Get the presentation part from the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the presentation from the presentation part. 
        Dim presentation As Presentation = presentationPart.Presentation

        ' Get the list of slide IDs in the presentation.
        Dim slideIdList As SlideIdList = presentation.SlideIdList

        ' Get the slide ID of the specified slide.
        Dim slideId As SlideId = CType(slideIdList.ChildElements(slideIndex), SlideId)

        ' Get the relationship ID of the specified slide.
        Dim slideRelId As String = slideId.RelationshipId

        ' Remove the slide from the slide list.
        slideIdList.RemoveChild(slideId)
        ' Remove references to the slide from all custom shows.
        If ((presentation.CustomShowList) IsNot Nothing) Then

            ' Iterate through the list of custom shows.
            For Each customShow As System.Object In presentation.CustomShowList.Elements(Of _
                                   DocumentFormat.OpenXml.Presentation.CustomShow)()

                If ((customShow.SlideList) IsNot Nothing) Then

                    ' Declare a linked list.
                    Dim slideListEntries As LinkedList(Of SlideListEntry) = New LinkedList(Of SlideListEntry)

                    ' Iterate through all the slides in the custom show.
                    For Each slideListEntry As SlideListEntry In customShow.SlideList.Elements

                        ' Find the slide reference to be removed from the custom show.
                        If (((slideListEntry.Id) IsNot Nothing) _
                                    AndAlso (slideListEntry.Id = slideRelId)) Then

                            ' Add that slide reference to the end of the linked list.
                            slideListEntries.AddLast(slideListEntry)
                        End If
                    Next

                    ' Remove references to the slide from the custom show.
                    For Each slideListEntry As SlideListEntry In slideListEntries
                        customShow.SlideList.RemoveChild(slideListEntry)
                    Next
                End If
            Next
        End If

        ' Save the change to the presentation part.
        presentation.Save()

        ' Get the slide part for the specified slide.
        Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slideRelId), SlidePart)

        ' Remove the slide part.
        presentationPart.DeletePart(slidePart)

    End Sub
End Module