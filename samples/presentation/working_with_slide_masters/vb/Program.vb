
    Private Shared Function CreateSlideMasterPart(ByVal slideLayoutPart1 As SlideLayoutPart) As SlideMasterPart
                Dim slideMasterPart1 As SlideMasterPart = slideLayoutPart1.AddNewPart(Of SlideMasterPart)("rId1")
                Dim slideMaster As New SlideMaster(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), _
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title Placeholder 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With { _
                  .Type = PlaceholderValues.Title _
                })), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph())))), New P.ColorMap() With { _
                  .Background1 = D.ColorSchemeIndexValues.Light1, _
                  .Text1 = D.ColorSchemeIndexValues.Dark1, _
                  .Background2 = D.ColorSchemeIndexValues.Light2, _
                  .Text2 = D.ColorSchemeIndexValues.Dark2, _
                  .Accent1 = D.ColorSchemeIndexValues.Accent1, _
                  .Accent2 = D.ColorSchemeIndexValues.Accent2, _
                  .Accent3 = D.ColorSchemeIndexValues.Accent3, _
                  .Accent4 = D.ColorSchemeIndexValues.Accent4, _
                  .Accent5 = D.ColorSchemeIndexValues.Accent5, _
                  .Accent6 = D.ColorSchemeIndexValues.Accent6, _
                  .Hyperlink = D.ColorSchemeIndexValues.Hyperlink, _
                  .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink _
                }, New SlideLayoutIdList(New SlideLayoutId() With { _
                  .Id = CType(2147483649UI, UInt32Value), _
                  .RelationshipId = "rId1" _
                }), New TextStyles(New TitleStyle(), New BodyStyle(), New OtherStyle()))
                slideMasterPart1.SlideMaster = slideMaster

                Return slideMasterPart1
            End Function


    Private Shared Function CreateSlideMasterPart(ByVal slideLayoutPart1 As SlideLayoutPart) As SlideMasterPart
                Dim slideMasterPart1 As SlideMasterPart = slideLayoutPart1.AddNewPart(Of SlideMasterPart)("rId1")
                Dim slideMaster As New SlideMaster(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), _
                    New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title Placeholder 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With { _
                  .Type = PlaceholderValues.Title _
                })), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph())))), New P.ColorMap() With { _
                  .Background1 = D.ColorSchemeIndexValues.Light1, _
                  .Text1 = D.ColorSchemeIndexValues.Dark1, _
                  .Background2 = D.ColorSchemeIndexValues.Light2, _
                  .Text2 = D.ColorSchemeIndexValues.Dark2, _
                  .Accent1 = D.ColorSchemeIndexValues.Accent1, _
                  .Accent2 = D.ColorSchemeIndexValues.Accent2, _
                  .Accent3 = D.ColorSchemeIndexValues.Accent3, _
                  .Accent4 = D.ColorSchemeIndexValues.Accent4, _
                  .Accent5 = D.ColorSchemeIndexValues.Accent5, _
                  .Accent6 = D.ColorSchemeIndexValues.Accent6, _
                  .Hyperlink = D.ColorSchemeIndexValues.Hyperlink, _
                  .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink _
                }, New SlideLayoutIdList(New SlideLayoutId() With { _
                  .Id = CType(2147483649UI, UInt32Value), _
                  .RelationshipId = "rId1" _
                }), New TextStyles(New TitleStyle(), New BodyStyle(), New OtherStyle()))
                slideMasterPart1.SlideMaster = slideMaster

                Return slideMasterPart1
            End Function
