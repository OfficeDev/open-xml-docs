Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports Drawing = DocumentFormat.OpenXml.Drawing

Module Program
    Sub Main(args As String())
        SetPPTShapeColor(args(0))
    End Sub

    ' <Snippet0>
    ' Change the fill color of a shape.
    ' The test file must have a shape on the first slide.
    Sub SetPPTShapeColor(docName As String)
        ' <Snippet1>
        Using ppt As PresentationDocument = PresentationDocument.Open(docName, True)
            ' </Snippet1>
            ' <Snippet2>
            ' Get the relationship ID of the first slide.
            Dim presentationPart As PresentationPart = If(ppt.PresentationPart, ppt.AddPresentationPart())
            presentationPart.Presentation.SlideIdList = If(presentationPart.Presentation.SlideIdList, New SlideIdList())
            Dim slideId As SlideId = presentationPart.Presentation.SlideIdList.GetFirstChild(Of SlideId)()

            If slideId IsNot Nothing Then
                Dim relId As String = slideId.RelationshipId

                If relId IsNot Nothing Then
                    ' Get the slide part from the relationship ID.
                    Dim slidePart As SlidePart = CType(presentationPart.GetPartById(relId), SlidePart)
                    ' </Snippet2>
                    ' <Snippet3>
                    ' Get or add the shape tree
                    slidePart.Slide.CommonSlideData = If(slidePart.Slide.CommonSlideData, New CommonSlideData())

                    ' Get the shape tree that contains the shape to change.
                    slidePart.Slide.CommonSlideData.ShapeTree = If(slidePart.Slide.CommonSlideData.ShapeTree, New ShapeTree())

                    ' Get the first shape in the shape tree.
                    Dim shape As Shape = slidePart.Slide.CommonSlideData.ShapeTree.GetFirstChild(Of Shape)()

                    If shape IsNot Nothing Then
                        ' Get or add the shape properties element of the shape.
                        shape.ShapeProperties = If(shape.ShapeProperties, New ShapeProperties())

                        ' Get or add the fill reference.
                        Dim solidFill As Drawing.SolidFill = shape.ShapeProperties.GetFirstChild(Of Drawing.SolidFill)()

                        ' Add solid fill element if it is missing and assign it to solidFill
                        If solidFill Is Nothing Then
                            shape.ShapeProperties.AddChild(New Drawing.SolidFill())
                            solidFill = shape.ShapeProperties.GetFirstChild(Of Drawing.SolidFill)()
                        End If

                        ' Set the fill color to SchemeColor
                        solidFill.SchemeColor = New Drawing.SchemeColor() With {.Val = Drawing.SchemeColorValues.Accent2}
                    End If
                    ' </Snippet3>
                End If
            End If
        End Using
    End Sub
    ' </Snippet0>
End Module
