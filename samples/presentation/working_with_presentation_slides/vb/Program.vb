

Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports P = DocumentFormat.OpenXml.Presentation

Module MyModule
    Function CreateSlidePart(ByVal presentationPart As PresentationPart) As SlidePart
        Dim slidePart1 As SlidePart = presentationPart.AddNewPart(Of SlidePart)("rId2")
        slidePart1.Slide = New Slide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With {
                 .Id = CType(1UI, UInt32Value),
                  .Name = ""
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New GroupShapeProperties(New TransformGroup()),
                   New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With {
                  .Id = CType(2UI, UInt32Value),
                  .Name = "Title 1"
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {
                  .NoGrouping = True
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(),
                   New ListStyle(), New Paragraph(New EndParagraphRunProperties() With {
                  .Language = "en-US"
                }))))), New ColorMapOverride(New MasterColorMapping()))
        Return slidePart1
    End Function
End Module
