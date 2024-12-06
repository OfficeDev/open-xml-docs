Imports System.Text
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation

Module Program
    Sub Main(args As String())
        If args.Length = 2 Then
            Dim filePath As String = args(0)
            Dim slideIndex As Integer = Integer.Parse(args(1))

            ' <Snippet5>
            For Each text As String In TextInSlide.GetAllTextInSlide(filePath, slideIndex)
                Console.WriteLine(text)
            Next
            ' </Snippet5>
        Else
            Console.WriteLine("Usage: <program> <presentationFile> <slideIndex>")
        End If
    End Sub
End Module

Public Class TextInSlide
    ' <Snippet>
    ' <Snippet2>
    ' Get all the text in a slide.
    Public Shared Function GetAllTextInSlide(presentationFile As String, slideIndex As Integer) As String()
        ' <Snippet1>
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' </Snippet1>
            ' Pass the presentation and the slide index
            ' to the next GetAllTextInSlide method, and
            ' then return the array of strings it returns. 
            Return GetAllTextInSlide(presentationDocument, slideIndex)
        End Using
    End Function
    ' </Snippet2>
    ' <Snippet3>
    Private Shared Function GetAllTextInSlide(presentationDocument As PresentationDocument, slideIndex As Integer) As String()
        ' Verify that the slide index is not out of range.
        If slideIndex < 0 Then
            Throw New ArgumentOutOfRangeException(NameOf(slideIndex))
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
                Dim slideIds As DocumentFormat.OpenXml.OpenXmlElementList = presentation.SlideIdList.ChildElements

                ' If the slide ID is in range...
                If slideIndex < slideIds.Count Then
                    ' Get the relationship ID of the slide.
                    Dim slidePartRelationshipId As String = CType(slideIds(slideIndex), SlideId).RelationshipId

                    If slidePartRelationshipId Is Nothing Then
                        Return Array.Empty(Of String)()
                    End If

                    ' Get the specified slide part from the relationship ID.
                    Dim slidePart As SlidePart = CType(presentationPart.GetPartById(slidePartRelationshipId), SlidePart)

                    ' Pass the slide part to the next method, and
                    ' then return the array of strings that method
                    ' returns to the previous method.
                    Return GetAllTextInSlide(slidePart)
                End If
            End If
        End If

        ' Else, return an empty array.
        Return Array.Empty(Of String)()
    End Function
    ' </Snippet3>
    ' <Snippet4>
    Private Shared Function GetAllTextInSlide(slidePart As SlidePart) As String()
        ' Verify that the slide part exists.
        If slidePart Is Nothing Then
            Throw New ArgumentNullException(NameOf(slidePart))
        End If

        ' Create a new linked list of strings.
        Dim texts As New LinkedList(Of String)()

        ' If the slide exists...
        If slidePart.Slide IsNot Nothing Then
            ' Iterate through all the paragraphs in the slide.
            For Each paragraph As DocumentFormat.OpenXml.Drawing.Paragraph In slidePart.Slide.Descendants(Of DocumentFormat.OpenXml.Drawing.Paragraph)()
                ' Create a new string builder.                    
                Dim paragraphText As New StringBuilder()

                ' Iterate through the lines of the paragraph.
                For Each text As DocumentFormat.OpenXml.Drawing.Text In paragraph.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)()
                    ' Append each line to the previous lines.
                    paragraphText.Append(text.Text)
                Next

                If paragraphText.Length > 0 Then
                    ' Add each paragraph to the linked list.
                    texts.AddLast(paragraphText.ToString())
                End If
            Next
        End If

        ' Return an array of strings.
        Return texts.ToArray()
    End Function
    ' </Snippet4>
    ' </Snippet>
End Class
