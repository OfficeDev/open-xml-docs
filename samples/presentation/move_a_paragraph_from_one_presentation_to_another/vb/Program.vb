Imports System.Linq
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Packaging
Imports Drawing = DocumentFormat.OpenXml.Drawing


Module MyModule
' Moves a paragraph range in a TextBody shape in the source document
    ' to another TextBody shape in the target document.
    Public Sub MoveParagraphToPresentation(ByVal sourceFile As String, ByVal targetFile As String)

        ' Open the source file.
        Dim sourceDoc As PresentationDocument = PresentationDocument.Open(sourceFile, True)

        ' Open the target file.
        Dim targetDoc As PresentationDocument = PresentationDocument.Open(targetFile, True)

        ' Get the first slide in the source presentation.
        Dim slide1 As SlidePart = GetFirstSlide(sourceDoc)

        ' Get the first TextBody shape in it.
        Dim textBody1 As TextBody = slide1.Slide.Descendants(Of TextBody).First()

        ' Get the first paragraph in the TextBody shape.
        ' Note: Drawing is the alias of the namespace DocumentFormat.OpenXml.Drawing
        Dim p1 As Drawing.Paragraph = textBody1.Elements(Of Drawing.Paragraph).First()

        ' Get the first slide in the target presentation.
        Dim slide2 As SlidePart = GetFirstSlide(targetDoc)

        ' Get the first TextBody shape in it.
        Dim textBody2 As TextBody = slide2.Slide.Descendants(Of TextBody).First()

        ' Clone the source paragraph and insert the cloned paragraph into the target TextBody shape.
        textBody2.Append(p1.CloneNode(True))

        ' Remove the source paragraph from the source file.
        textBody1.RemoveChild(Of Drawing.Paragraph)(p1)

        ' Replace it with an empty one, because a paragraph is required for a TextBody shape.
        textBody1.AppendChild(Of Drawing.Paragraph)(New Drawing.Paragraph())

        ' Save the slide in the source file.
        slide1.Slide.Save()

        ' Save the slide in the target file.
        slide2.Slide.Save()

    End Sub
    ' Get the slide part of the first slide in the presentation document.
    Public Function GetFirstSlide(ByVal presentationDoc As PresentationDocument) As SlidePart

        ' Get relationship ID of the first slide.
        Dim part As PresentationPart = presentationDoc.PresentationPart
        Dim slideId As SlideId = part.Presentation.SlideIdList.GetFirstChild(Of SlideId)()
        Dim relId As String = slideId.RelationshipId

        ' Get the slide part by the relationship ID.
        Dim slidePart As SlidePart = CType(part.GetPartById(relId), SlidePart)

        Return slidePart

    End Function
End Module
