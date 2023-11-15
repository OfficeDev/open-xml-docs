Imports System.Text
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation


Module MyModule

    Sub Main(args As String())
    End Sub

    ' Get all the text in a slide.
    Public Function GetAllTextInSlide(ByVal presentationFile As String, ByVal slideIndex As Integer) As String()
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, False)
            ' Pass the presentation and the slide index
            ' to the next GetAllTextInSlide method, and
            ' then return the array of strings it returns. 
            Return GetAllTextInSlide(presentationDocument, slideIndex)
        End Using
    End Function
    Public Function GetAllTextInSlide(ByVal presentationDocument As PresentationDocument, ByVal slideIndex As Integer) As String()
        ' Verify that the presentation document exists.
        If presentationDocument Is Nothing Then
            Throw New ArgumentNullException("presentationDocument")
        End If

        ' Verify that the slide index is not out of range.
        If slideIndex < 0 Then
            Throw New ArgumentOutOfRangeException("slideIndex")
        End If

        ' Get the presentation part of the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Verify that the presentation part and presentation exist.
        If presentationPart IsNot Nothing AndAlso presentationPart.Presentation IsNot Nothing Then
            ' Get the Presentation object from the presentation part.
            Dim presentation As Presentation = presentationPart.Presentation

            ' Verify that the slide ID list exists.
            If presentation.SlideIdList IsNot Nothing Then
                ' Get the collection of slide IDs from the slide ID list.
                Dim slideIds = presentation.SlideIdList.ChildElements

                ' If the slide ID is in range...
                If slideIndex < slideIds.Count Then
                    ' Get the relationship ID of the slide.
                    Dim slidePartRelationshipId As String = (TryCast(slideIds(slideIndex), SlideId)).RelationshipId

                    ' Get the specified slide part from the relationship ID.
                    Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slidePartRelationshipId), SlidePart)

                    ' Pass the slide part to the next method, and
                    ' then return the array of strings that method
                    ' returns to the previous method.
                    Return GetAllTextInSlide(slidePart)
                End If
            End If
        End If

        ' Else, return null.
        Return Nothing
    End Function
    Public Function GetAllTextInSlide(ByVal slidePart As SlidePart) As String()
        ' Verify that the slide part exists.
        If slidePart Is Nothing Then
            Throw New ArgumentNullException("slidePart")
        End If

        ' Create a new linked list of strings.
        Dim texts As New LinkedList(Of String)()

        ' If the slide exists...
        If slidePart.Slide IsNot Nothing Then
            ' Iterate through all the paragraphs in the slide.
            For Each paragraph In slidePart.Slide.Descendants(Of DocumentFormat.OpenXml.Drawing.Paragraph)()
                ' Create a new string builder.                    
                Dim paragraphText As New StringBuilder()

                ' Iterate through the lines of the paragraph.
                For Each Text In paragraph.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)()
                    ' Append each line to the previous lines.
                    paragraphText.Append(Text.Text)
                Next Text

                If paragraphText.Length > 0 Then
                    ' Add each paragraph to the linked list.
                    texts.AddLast(paragraphText.ToString())
                End If
            Next paragraph
        End If

        If texts.Count > 0 Then
            ' Return an array of strings.
            Return texts.ToArray()
        Else
            Return Nothing
        End If
    End Function
End Module