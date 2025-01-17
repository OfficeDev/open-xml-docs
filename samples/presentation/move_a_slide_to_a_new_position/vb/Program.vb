Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation

Module Program
    Sub Main(args As String())
        Dim from As Integer
        Dim toIndex As Integer
        Dim fromIsValid As Boolean = Integer.TryParse(args(1), from)
        Dim toIsValid As Boolean = Integer.TryParse(args(2), toIndex)

        If fromIsValid AndAlso toIsValid Then
            SlideMover.MoveSlide(args(0), from, toIndex)
        End If
    End Sub
End Module

Public Class SlideMover
    ' <Snippet0>
    ' Counting the slides in the presentation.
    Public Shared Function CountSlides(presentationFile As String) As Integer
        ' <Snippet1>
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' </Snippet1>
            ' <Snippet2>
            ' Pass the presentation to the next CountSlides method
            ' and return the slide count.
            Return CountSlides(presentationDocument)
            ' </Snippet2>
        End Using
    End Function

    ' Count the slides in the presentation.
    Private Shared Function CountSlides(presentationDocument As PresentationDocument) As Integer
        ' <Snippet3>
        Dim slidesCount As Integer = 0

        ' Get the presentation part of document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the slide count from the SlideParts.
        If presentationPart IsNot Nothing Then
            slidesCount = presentationPart.SlideParts.Count()
        End If

        ' Return the slide count to the previous method.
        Return slidesCount
        ' </Snippet3>
    End Function

    ' <Snippet4>
    ' Move a slide to a different position in the slide order in the presentation.
    Public Shared Sub MoveSlide(presentationFile As String, from As Integer, toIndex As Integer)
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
            MoveSlide(presentationDocument, from, toIndex)
        End Using
    End Sub
    ' </Snippet4>

    ' <Snippet5>
    ' Move a slide to a different position in the slide order in the presentation.
    Private Shared Sub MoveSlide(presentationDocument As PresentationDocument, from As Integer, toIndex As Integer)
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentationDocument))
        End If

        ' Call the CountSlides method to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        ' Verify that both from and to positions are within range and different from one another.
        If from < 0 OrElse from >= slidesCount Then
            Throw New ArgumentOutOfRangeException(NameOf(from))
        End If

        If toIndex < 0 OrElse from >= slidesCount OrElse toIndex = from Then
            Throw New ArgumentOutOfRangeException(NameOf(toIndex))
        End If
        ' </Snippet5>

        ' <Snippet6>
        ' Get the presentation part from the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' The slide count is not zero, so the presentation must contain slides.            
        Dim presentation As Presentation = presentationPart?.Presentation

        If presentation Is Nothing Then
            Throw New ArgumentNullException(NameOf(presentation))
        End If

        Dim slideIdList As SlideIdList = presentation.SlideIdList

        If slideIdList Is Nothing Then
            Throw New ArgumentNullException(NameOf(slideIdList))
        End If

        ' Get the slide ID of the source slide.
        Dim sourceSlide As SlideId = TryCast(slideIdList.ChildElements(from), SlideId)

        If sourceSlide Is Nothing Then
            Throw New ArgumentNullException(NameOf(sourceSlide))
        End If

        Dim targetSlide As SlideId = Nothing

        ' Identify the position of the target slide after which toIndex move the source slide.
        If toIndex = 0 Then
            targetSlide = Nothing
        ElseIf from < toIndex Then
            targetSlide = TryCast(slideIdList.ChildElements(toIndex), SlideId)
        Else
            targetSlide = TryCast(slideIdList.ChildElements(toIndex - 1), SlideId)
        End If
        ' </Snippet6>
        ' <Snippet7>
        ' Remove the source slide from its current position.
        sourceSlide.Remove()

        ' Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide)
        ' </Snippet7>
    End Sub
    ' </Snippet0>
End Class


