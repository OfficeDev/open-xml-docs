Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports P = DocumentFormat.OpenXml.Presentation

Module MyModule

    Sub Main(args As String())
    End Sub

    Function CreateSlideLayoutPart(ByVal slidePart1 As SlidePart) As SlideLayoutPart
        Dim slideLayoutPart1 As SlideLayoutPart = slidePart1.AddNewPart(Of SlideLayoutPart)("rId1")
        Dim slideLayout As New SlideLayout(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With {
                  .Id = CType(1UI, UInt32Value),
                  .Name = ""
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()),
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With {
                  .Id = CType(2UI, UInt32Value),
                  .Name = ""
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {
                  .NoGrouping = True
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(),
                    New ListStyle(), New Paragraph(New EndParagraphRunProperties()))))), New ColorMapOverride(New MasterColorMapping()))
        slideLayoutPart1.SlideLayout = slideLayout
        Return slideLayoutPart1
    End Function
End Module