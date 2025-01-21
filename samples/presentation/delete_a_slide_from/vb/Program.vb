Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation

Public Module Program
    Public Sub Main(args As String())
        Dim count As Integer = CountSlides(args(0))

        Console.WriteLine($"{count} slides found")

        DeleteSlide(args(0), 0)
    End Sub

    ' <Snippet0>
    ' Get the presentation object and pass it to the next CountSlides method.
    Private Function CountSlides(presentationFile As String) As Integer
        ' <Snippet1>
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' </Snippet1>
            ' <Snippet3>
            ' Pass the presentation to the next CountSlide method
            ' and return the slide count.
            Return CountSlides(presentationDocument)
            ' </Snippet3>
        End Using
    End Function

    ' Count the slides in the presentation.
    Private Function CountSlides(presentationDocument As PresentationDocument) As Integer
        ' <Snippet4>
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        Dim slidesCount As Integer = 0

        ' Get the presentation part of document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the slide count from the SlideParts.
        If presentationPart IsNot Nothing Then
            slidesCount = presentationPart.SlideParts.Count()
        End If

        ' Return the slide count to the previous method.
        Return slidesCount
        ' </Snippet4>
    End Function

    ' <Snippet5>
    ' Get the presentation object and pass it to the next DeleteSlide method.
    Private Sub DeleteSlide(presentationFile As String, slideIndex As Integer)
        ' <Snippet2>
        ' Open the source document as read/write.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
            ' </Snippet2>
            ' Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
            DeleteSlide(presentationDocument, slideIndex)
        End Using
    End Sub
    ' </Snippet5>

    ' <Snippet6>
    ' Delete the specified slide from the presentation.
    Private Sub DeleteSlide(presentationDocument As PresentationDocument, slideIndex As Integer)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        ' Use the CountSlides sample to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        If slideIndex < 0 OrElse slideIndex >= slidesCount Then
            Throw New ArgumentOutOfRangeException(NameOf(slideIndex))
        End If

        ' Get the presentation part from the presentation document. 
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the presentation from the presentation part.
        Dim presentation As Presentation = presentationPart?.Presentation

        ' Get the list of slide IDs in the presentation.
        Dim slideIdList As SlideIdList = presentation?.SlideIdList

        ' Get the slide ID of the specified slide
        Dim slideId As SlideId = TryCast(slideIdList?.ChildElements(slideIndex), SlideId)

        ' Get the relationship ID of the slide.
        Dim slideRelId As String = slideId?.RelationshipId

        ' If there's no relationship ID, there's no slide to delete.
        If slideRelId Is Nothing Then
            Return
        End If

        ' Remove the slide from the slide list.
        slideIdList.RemoveChild(slideId)
        ' </Snippet6>

        ' <Snippet7>
        ' Remove references to the slide from all custom shows.
        If presentation.CustomShowList IsNot Nothing Then
            ' Iterate through the list of custom shows.
            For Each customShow In presentation.CustomShowList.Elements(Of CustomShow)()
                If customShow.SlideList IsNot Nothing Then
                    ' Declare a link list of slide list entries.
                    Dim slideListEntries As New LinkedList(Of SlideListEntry)()
                    For Each slideListEntry As SlideListEntry In customShow.SlideList.Elements()
                        ' Find the slide reference to remove from the custom show.
                        If slideListEntry.Id IsNot Nothing AndAlso slideListEntry.Id = slideRelId Then
                            slideListEntries.AddLast(slideListEntry)
                        End If
                    Next

                    ' Remove all references to the slide from the custom show.
                    For Each slideListEntry As SlideListEntry In slideListEntries
                        customShow.SlideList.RemoveChild(slideListEntry)
                    Next
                End If
            Next
        End If
        ' </Snippet7>

        ' <Snippet8>
        ' Get the slide part for the specified slide.
        Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slideRelId), SlidePart)

        ' Remove the slide part.
        presentationPart.DeletePart(slidePart)
        ' </Snippet8>
    End Sub
    ' </Snippet0>
End Module
