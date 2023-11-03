
    Private Shared Function CreateSlidePart(ByVal presentationPart As PresentationPart) As SlidePart
                Dim slidePart1 As SlidePart = presentationPart.AddNewPart(Of SlidePart)("rId2")
                slidePart1.Slide = New Slide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                 .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New GroupShapeProperties(New TransformGroup()), _
                   New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), _
                   New ListStyle(), New Paragraph(New EndParagraphRunProperties() With { _
                  .Language = "en-US" _
                }))))), New ColorMapOverride(New MasterColorMapping()))
                Return slidePart1
            End Function


    New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), _
                   New ListStyle(), New Paragraph(New EndParagraphRunProperties() With { _
                  .Language = "en-US" })))
