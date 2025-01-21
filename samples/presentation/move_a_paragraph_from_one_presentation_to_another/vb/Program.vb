Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports System
Imports System.Linq
Imports Drawing = DocumentFormat.OpenXml.Drawing

Module Program
    Sub Main(args As String())
        MoveParagraphToPresentation(args(0), args(1))
    End Sub

    ' <Snippet>
    ' <Snippet2>
    ' Moves a paragraph range in a TextBody shape in the source document
    ' to another TextBody shape in the target document.
    Sub MoveParagraphToPresentation(sourceFile As String, targetFile As String)
        ' Open the source file as read/write.
        ' <Snippet1>
        Using sourceDoc As PresentationDocument = PresentationDocument.Open(sourceFile, True)
            ' </Snippet1>
            ' Open the target file as read/write.
            Using targetDoc As PresentationDocument = PresentationDocument.Open(targetFile, True)
                ' Get the first slide in the source presentation.
                Dim slide1 As SlidePart = GetFirstSlide(sourceDoc)

                ' Get the first TextBody shape in it.
                Dim textBody1 As TextBody = slide1.Slide.Descendants(Of TextBody)().First()

                ' Get the first paragraph in the TextBody shape.
                ' Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
                Dim p1 As Drawing.Paragraph = textBody1.Elements(Of Drawing.Paragraph)().First()

                ' Get the first slide in the target presentation.
                Dim slide2 As SlidePart = GetFirstSlide(targetDoc)

                ' Get the first TextBody shape in it.
                Dim textBody2 As TextBody = slide2.Slide.Descendants(Of TextBody)().First()

                ' Clone the source paragraph and insert the cloned paragraph into the target TextBody shape.
                ' Passing "true" creates a deep clone, which creates a copy of the 
                ' Paragraph object and everything directly or indirectly referenced by that object.
                textBody2.Append(p1.CloneNode(True))

                ' Remove the source paragraph from the source file.
                textBody1.RemoveChild(p1)

                ' Replace the removed paragraph with a placeholder.
                textBody1.AppendChild(New Drawing.Paragraph())
            End Using
        End Using
    End Sub
    ' </Snippet2>
    ' <Snippet3>
    ' Get the slide part of the first slide in the presentation document.
    Function GetFirstSlide(presentationDocument As PresentationDocument) As SlidePart
        ' Get relationship ID of the first slide
        Dim part As PresentationPart = If(presentationDocument.PresentationPart, presentationDocument.AddPresentationPart())
        Dim slideIdList As SlideIdList = If(part.Presentation.SlideIdList, part.Presentation.AppendChild(New SlideIdList()))
        Dim slideId As SlideId = If(part.Presentation.SlideIdList?.GetFirstChild(Of SlideId)(), slideIdList.AppendChild(New SlideId()))
        Dim relId As String = slideId.RelationshipId

        If relId Is Nothing Then
            Throw New ArgumentNullException(NameOf(relId))
        End If

        ' Get the slide part by the relationship ID.
        Dim slidePart As SlidePart = CType(part.GetPartById(relId), SlidePart)

        Return slidePart
    End Function
    ' </Snippet3>
    ' </Snippet>
End Module

