Imports System
Imports System.Linq
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Packaging


Module MyModule
' Count the slides in the presentation.
    Public Function CountSlides(ByVal presentationFile As String) As Integer

        ' Open the presentation as read-only.
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, False)
        Using (presentationDocument)

            ' Pass the presentation to the next CountSlide method
            ' and return the slide count.
            Return CountSlides(presentationDocument)
        End Using
    End Function
    ' Count the slides in the presentation.
    Public Function CountSlides(ByVal presentationDocument As PresentationDocument) As Integer

        ' Check for a null document object.
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If
        Dim slidesCount As Integer = 0

        ' Get the presentation part of the document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart
        If ((Not (presentationPart) Is Nothing) AndAlso (Not (presentationPart.Presentation) Is Nothing)) Then

            ' Get the Presentation object from the presentation part.
            Dim presentation As Presentation = presentationPart.Presentation
            If (Not (presentation.SlideIdList) Is Nothing) Then

                ' Get the slide count from the slide ID list.
                slidesCount = presentation.SlideIdList.Elements.Count()

            End If
        End If

        ' Return the slide count to the previous function.
        Return slidesCount
    End Function
    ' Move a slide to a different position in the slide order in the presentation.
    Public Sub MoveSlide(ByVal presentationFile As String, ByVal from As Integer, ByVal moveTo As Integer)
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, True)

        Using (presentationDocument)
            MoveSlide(presentationDocument, from, moveTo)
        End Using

    End Sub
    ' Move a slide to a different position in the slide order in the presentation.
    Public Sub MoveSlide(ByVal presentationDocument As PresentationDocument, ByVal from As Integer, ByVal moveTo As Integer)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Use the CountSlides sample to get the number of slides in the presentation.
        Dim slidesCount As Integer = CountSlides(presentationDocument)

        ' Verify that both from and to positions are within range and different from one another.
        If ((from < 0) OrElse (from >= slidesCount)) Then
            Throw New ArgumentOutOfRangeException("from")
        End If

        If ((moveTo < 0) _
                    OrElse ((from >= slidesCount) _
                    OrElse (moveTo = from))) Then
            Throw New ArgumentOutOfRangeException("moveTo")
        End If

        ' Get the presentation part from the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' The slide count is not zero, so the presentation must contain slides. 
        Dim presentation As Presentation = presentationPart.Presentation
        Dim slideIdList As SlideIdList = presentation.SlideIdList

        ' Get the slide ID of the source slide.
        Dim sourceSlide As SlideId = CType(slideIdList.ChildElements(from), SlideId)
        Dim targetSlide As SlideId = Nothing

        ' Identify the position of the target slide after which to move the source slide.
        If (moveTo = 0) Then
            targetSlide = Nothing
        ElseIf (from < moveTo) Then
            targetSlide = CType(slideIdList.ChildElements(moveTo), SlideId)
        Else
            targetSlide = CType(slideIdList.ChildElements((moveTo - 1)), SlideId)
        End If

        ' Remove the source slide from its current position.
        sourceSlide.Remove()

        ' Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide)

        ' Save the modified presentation.
        presentation.Save()

    End Sub
End Module
