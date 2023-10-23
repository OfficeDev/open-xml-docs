Module Program `
  Sub Main(args As String())`
  End Sub`

  
    Private Shared Function CreateSlideLayoutPart(ByVal slidePart1 As SlidePart) As SlideLayoutPart
                Dim slideLayoutPart1 As SlideLayoutPart = slidePart1.AddNewPart(Of SlideLayoutPart)("rId1")
                Dim slideLayout As New SlideLayout(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), _
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), _
                    New ListStyle(), New Paragraph(New EndParagraphRunProperties()))))), New ColorMapOverride(New MasterColorMapping()))
                slideLayoutPart1.SlideLayout = slideLayout
                Return slideLayoutPart1
            End Function

    Private Shared Function CreateSlideLayoutPart(ByVal slidePart1 As SlidePart) As SlideLayoutPart
                Dim slideLayoutPart1 As SlideLayoutPart = slidePart1.AddNewPart(Of SlideLayoutPart)("rId1")
                Dim slideLayout As New SlideLayout(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), _
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), _
                    New ListStyle(), New Paragraph(New EndParagraphRunProperties()))))), New ColorMapOverride(New MasterColorMapping()))
                slideLayoutPart1.SlideLayout = slideLayout
                Return slideLayoutPart1
            End Function
End Module