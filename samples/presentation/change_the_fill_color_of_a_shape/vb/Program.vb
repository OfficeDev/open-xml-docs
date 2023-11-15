Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports Drawing = DocumentFormat.OpenXml.Drawing

Module Program
    Sub Main(args As String())
    End Sub



    ' Change the fill color of a shape.
    ' The test file must have a filled shape as the first shape on the first slide.
    Public Sub SetPPTShapeColor(ByVal docName As String)
        Using ppt As PresentationDocument = PresentationDocument.Open(docName, True)
            ' Get the relationship ID of the first slide.
            Dim part As PresentationPart = ppt.PresentationPart
            Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements
            Dim relId As String = TryCast(slideIds(0), SlideId).RelationshipId

            ' Get the slide part from the relationship ID.
            Dim slide As SlidePart = DirectCast(part.GetPartById(relId), SlidePart)

            If slide IsNot Nothing Then
                ' Get the shape tree that contains the shape to change.
                Dim tree As ShapeTree = slide.Slide.CommonSlideData.ShapeTree

                ' Get the first shape in the shape tree.
                Dim shape As Shape = tree.GetFirstChild(Of Shape)()

                If shape IsNot Nothing Then
                    ' Get the style of the shape.
                    Dim style As ShapeStyle = shape.ShapeStyle

                    ' Get the fill reference.
                    Dim fillRef As Drawing.FillReference = style.FillReference

                    ' Set the fill color to SchemeColor Accent 6;
                    fillRef.SchemeColor = New Drawing.SchemeColor()
                    fillRef.SchemeColor.Val = Drawing.SchemeColorValues.Accent6

                    ' Save the modified slide.
                    slide.Slide.Save()
                End If
            End If
        End Using
    End Sub
End Module