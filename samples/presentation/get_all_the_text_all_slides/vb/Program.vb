Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports A = DocumentFormat.OpenXml.Drawing

Module Program
    Sub Main(args As String())
        If args.Length = 2 Then
            Dim filePath As String = args(0)
            Dim slideIndex As Integer = Integer.Parse(args(1))
            Dim text As String
            GetSlideIdAndText(text, filePath, slideIndex)
            Console.WriteLine($"Slide #{slideIndex + 1} contains: {text}")
        End If

        ' <Snippet2>
        If args.Length = 1 Then
            Dim path As String = args(0)
            Dim numberOfSlides As Integer = CountSlides(path)
            Console.WriteLine($"Number of slides = {numberOfSlides}")

            For i As Integer = 0 To numberOfSlides - 1
                Dim text As String
                GetSlideIdAndText(text, path, i)
                Console.WriteLine($"Slide #{i + 1} contains: {text}")
            Next
        End If
        ' </Snippet2>
    End Sub

    ' <Snippet>
    Function CountSlides(presentationFile As String) As Integer
        ' <Snippet1>
        ' Open the presentation as read-only.
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
            ' </Snippet1>
            ' Pass the presentation to the next CountSlides method
            ' and return the slide count.
            Return CountSlidesFromPresentation(presentationDocument)
        End Using
    End Function

    ' Count the slides in the presentation.
    Function CountSlidesFromPresentation(presentationDocument As PresentationDocument) As Integer
        ' Check for a null document object.
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
    End Function

    Sub GetSlideIdAndText(ByRef sldText As String, docName As String, index As Integer)
        Using ppt As PresentationDocument = PresentationDocument.Open(docName, False)
            ' Get the relationship ID of the first slide.
            Dim part As PresentationPart = ppt.PresentationPart
            Dim slideIds As OpenXmlElementList = If(part?.Presentation?.SlideIdList?.ChildElements, Nothing)

            If part Is Nothing OrElse slideIds.Count = 0 Then
                sldText = ""
                Return
            End If

            Dim relId As String = CType(slideIds(index), SlideId).RelationshipId

            If relId Is Nothing Then
                sldText = ""
                Return
            End If

            ' Get the slide part from the relationship ID.
            Dim slide As SlidePart = CType(part.GetPartById(relId), SlidePart)

            ' Build a StringBuilder object.
            Dim paragraphText As New StringBuilder()

            ' Get the inner text of the slide:
            Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
            For Each text As A.Text In texts
                paragraphText.Append(text.Text)
            Next
            sldText = paragraphText.ToString()
        End Using
    End Sub
    ' </Snippet>
End Module

