Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports D = DocumentFormat.OpenXml.Drawing
Imports P = DocumentFormat.OpenXml.Presentation

Module MyModule

    Sub Main(args As String())
    End Sub

    Function CreateSlideMasterPart(ByVal slideLayoutPart1 As SlideLayoutPart) As SlideMasterPart
        Dim slideMasterPart1 As SlideMasterPart = slideLayoutPart1.AddNewPart(Of SlideMasterPart)("rId1")
        Dim slideMaster As New SlideMaster(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With {
                  .Id = CType(1UI, UInt32Value),
                  .Name = ""
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()),
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With {
                  .Id = CType(2UI, UInt32Value),
                  .Name = "Title Placeholder 1"
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {
                  .NoGrouping = True
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {
                  .Type = PlaceholderValues.Title
                })), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph())))), New P.ColorMap() With {
                  .Background1 = D.ColorSchemeIndexValues.Light1,
                  .Text1 = D.ColorSchemeIndexValues.Dark1,
                  .Background2 = D.ColorSchemeIndexValues.Light2,
                  .Text2 = D.ColorSchemeIndexValues.Dark2,
                  .Accent1 = D.ColorSchemeIndexValues.Accent1,
                  .Accent2 = D.ColorSchemeIndexValues.Accent2,
                  .Accent3 = D.ColorSchemeIndexValues.Accent3,
                  .Accent4 = D.ColorSchemeIndexValues.Accent4,
                  .Accent5 = D.ColorSchemeIndexValues.Accent5,
                  .Accent6 = D.ColorSchemeIndexValues.Accent6,
                  .Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                  .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                }, New SlideLayoutIdList(New SlideLayoutId() With {
                  .Id = CType(2147483649UI, UInt32Value),
                  .RelationshipId = "rId1"
                }), New TextStyles(New TitleStyle(), New BodyStyle(), New OtherStyle()))
        slideMasterPart1.SlideMaster = slideMaster

        Return slideMasterPart1
    End Function
End Module