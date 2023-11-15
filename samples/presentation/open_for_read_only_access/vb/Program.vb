Imports System.Text
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports A = DocumentFormat.OpenXml.Drawing


Module MyModule

    Sub Main(args As String())
    End Sub

    Public Sub GetSlideIdAndText(ByRef sldText As String, ByVal docName As String, ByVal index As Integer)
        Using ppt As PresentationDocument = PresentationDocument.Open(docName, False)
            ' Get the relationship ID of the first slide.
            Dim part As PresentationPart = ppt.PresentationPart
            Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements
            Dim relId As String = TryCast(slideIds(index), SlideId).RelationshipId
            relId = TryCast(slideIds(index), SlideId).RelationshipId

            ' Get the slide part from the relationship ID.
            Dim slide As SlidePart = DirectCast(part.GetPartById(relId), SlidePart)

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
End Module