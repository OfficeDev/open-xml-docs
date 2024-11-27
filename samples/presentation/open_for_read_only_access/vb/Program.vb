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
        ' <Snippet4>
        Try
            Dim file As String = args(0)
            Dim i As Integer
            Dim isInt As Boolean = Integer.TryParse(args(1), i)

            If isInt Then
                Dim sldText As String
                GetSlideIdAndText(sldText, file, i)
                Console.WriteLine($"The text in slide #{i + 1} is {sldText}")
            End If
        Catch exp As ArgumentOutOfRangeException
            Console.Error.WriteLine(exp.Message)
        End Try
        ' </Snippet4>
    End Sub

    ' <Snippet0>
    Sub GetSlideIdAndText(ByRef sldText As String, docName As String, index As Integer)
        ' <Snippet1>
        Using ppt As PresentationDocument = PresentationDocument.Open(docName, False)
            ' </Snippet1>
            ' <Snippet2>
            ' Get the relationship ID of the first slide.
            Dim part As PresentationPart = ppt.PresentationPart
            Dim slideIds As OpenXmlElementList = If(part?.Presentation?.SlideIdList?.ChildElements, New OpenXmlElementList())

            ' If there are no slide IDs then there are no slides.
            If slideIds.Count = 0 Then
                sldText = ""
                Return
            End If

            Dim relId As String = TryCast(slideIds(index), SlideId)?.RelationshipId

            If relId Is Nothing Then
                sldText = ""
                Return
            End If
            ' </Snippet2>

            ' <Snippet3>
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
            ' </Snippet3>
        End Using
    End Sub
End Module
' </Snippet0>
