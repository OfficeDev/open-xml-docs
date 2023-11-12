Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports D = DocumentFormat.OpenXml.Drawing


Module MyModule
' Get a list of the titles of all the slides in the presentation.
    Public Function GetSlideTitles(ByVal presentationFile As String) As IList(Of String)

        ' Open the presentation as read-only.
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, False)
        Using (presentationDocument)
            Return GetSlideTitles(presentationDocument)
        End Using

    End Function
    ' Get a list of the titles of all the slides in the presentation.
    Public Function GetSlideTitles(ByVal presentationDocument As PresentationDocument) As IList(Of String)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Get a PresentationPart object from the PresentationDocument object.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart
        If ((Not (presentationPart) Is Nothing) _
           AndAlso (Not (presentationPart.Presentation) Is Nothing)) Then

            ' Get a Presentation object from the PresentationPart object.
            Dim presentation As Presentation = presentationPart.Presentation
            If (Not (presentation.SlideIdList) Is Nothing) Then

                Dim titlesList As List(Of String) = New List(Of String)

                ' Get the title of each slide in the slide order.
                For Each slideId As Object In presentation.SlideIdList.Elements(Of SlideId)()

                    Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slideId.RelationshipId.ToString()), SlidePart)

                    ' Get the slide title.
                    Dim title As String = GetSlideTitle(slidePart)

                    ' An empty title can also be added.
                    titlesList.Add(title)
                Next
                Return titlesList
            End If
        End If
        Return Nothing
    End Function
    ' Get the title string of the slide.
    Public Function GetSlideTitle(ByVal slidePart As SlidePart) As String
        If (slidePart Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Declare a paragraph separator.
        Dim paragraphSeparator As String = Nothing
        If (Not (slidePart.Slide) Is Nothing) Then

            ' Find all the title shapes.
            Dim shapes = From shape In slidePart.Slide.Descendants(Of Shape)() _
             Where (IsTitleShape(shape)) _
             Select shape

            Dim paragraphText As StringBuilder = New StringBuilder

            For Each shape As Object In shapes

                ' Get the text in each paragraph in this shape.
                For Each paragraph As Object In shape.TextBody.Descendants(Of D.Paragraph)()

                    ' Add a line break.
                    paragraphText.Append(paragraphSeparator)

                    For Each text As Object In paragraph.Descendants(Of D.Text)()
                        paragraphText.Append(text.Text)
                    Next

                    paragraphSeparator = "" & vbLf
                Next
            Next
            Return paragraphText.ToString
        End If
        Return String.Empty
    End Function
    ' Determines whether the shape is a title shape.
    Private Function IsTitleShape(ByVal shape As Shape) As Boolean
        Dim placeholderShape As Object = _
         shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild(Of PlaceholderShape)()
        If ((Not (placeholderShape) Is Nothing) _
           AndAlso ((Not (placeholderShape.Type) Is Nothing) _
           AndAlso placeholderShape.Type.HasValue)) Then
            Select Case placeholderShape.Type.Value

                ' Any title shape
                Case PlaceholderValues.Title
                    Return True

                    ' A centered title.
                Case PlaceholderValues.CenteredTitle
                    Return True
                Case Else
                    Return False
            End Select
        End If
        Return False
    End Function
End Module
