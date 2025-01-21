Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports D = DocumentFormat.OpenXml.Drawing

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        For Each title As String In GetSlideTitles(args(0))
            Console.WriteLine(title)
        Next
        ' </Snippet2>
    End Sub

    ' <Snippet>
    ' Get a list of the titles of all the slides in the presentation.
    Function GetSlideTitles(presentationFile As String) As IList(Of String)
        ' <Snippet1>
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' </Snippet1>
            Dim titles As IList(Of String) = GetSlideTitlesFromPresentation(presentationDocument)

            Return If(titles, Enumerable.Empty(Of String)().ToList())
        End Using
    End Function

    ' Get a list of the titles of all the slides in the presentation.
    Function GetSlideTitlesFromPresentation(presentationDocument As PresentationDocument) As IList(Of String)
        ' Get a PresentationPart object from the PresentationDocument object.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        If presentationPart IsNot Nothing AndAlso presentationPart.Presentation IsNot Nothing Then
            ' Get a Presentation object from the PresentationPart object.
            Dim presentation As Presentation = presentationPart.Presentation

            If presentation.SlideIdList IsNot Nothing Then
                Dim titlesList As New List(Of String)()

                ' Get the title of each slide in the slide order.
                For Each slideId As SlideId In presentation.SlideIdList.Elements(Of SlideId)()
                    If slideId.RelationshipId Is Nothing Then
                        Continue For
                    End If

                    Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slideId.RelationshipId), SlidePart)

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
    Function GetSlideTitle(slidePart As SlidePart) As String
        If slidePart Is Nothing Then
            Throw New ArgumentNullException(NameOf(slidePart))
        End If

        ' Declare a paragraph separator.
        Dim paragraphSeparator As String = Nothing

        If slidePart.Slide IsNot Nothing Then
            ' Find all the title shapes.
            Dim shapes = From shape In slidePart.Slide.Descendants(Of Shape)()
                         Where IsTitleShape(shape)
                         Select shape

            Dim paragraphText As New StringBuilder()

            For Each shape In shapes
                Dim paragraphs = shape.TextBody?.Descendants(Of D.Paragraph)()
                If paragraphs Is Nothing Then
                    Continue For
                End If

                ' Get the text in each paragraph in this shape.
                For Each paragraph In paragraphs
                    ' Add a line break.
                    paragraphText.Append(paragraphSeparator)

                    For Each text In paragraph.Descendants(Of D.Text)()
                        paragraphText.Append(text.Text)
                    Next

                    paragraphSeparator = vbLf
                Next
            Next

            Return paragraphText.ToString()
        End If

        Return String.Empty
    End Function

    ' Determines whether the shape is a title shape.
    Function IsTitleShape(shape As Shape) As Boolean
        Dim placeholderShape As PlaceholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild(Of PlaceholderShape)()

        If placeholderShape IsNot Nothing AndAlso placeholderShape.Type IsNot Nothing AndAlso placeholderShape.Type.HasValue Then
            Return placeholderShape.Type = PlaceholderValues.Title OrElse placeholderShape.Type = PlaceholderValues.CenteredTitle
        End If

        Return False
    End Function
    ' </Snippet>
End Module

