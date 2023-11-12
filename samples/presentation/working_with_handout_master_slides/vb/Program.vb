

Module MyModule
Private Shared Function CreateHandoutMasterPart(ByVal presentationPart As PresentationPart) As HandoutMasterPart
            Dim handoutMasterPart1 As HandoutMasterPart = presentationPart.AddNewPart(Of HandoutMasterPart)("rId3")
            handoutMasterPart1.HandoutMaster = New HandoutMaster(New CommonSlideData(New ShapeTree(New _
                P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(1UI, UInt32Value), _
             .Name = "" _
            }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New _
                GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(2UI, UInt32Value), _
             .Name = "Title 1" _
            }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
             .NoGrouping = True _
            }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New _
                P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties() With { _
             .Language = "en-US" _
            }))))), New P.ColorMap() With { _
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
            })
            Return handoutMasterPart1
        End Function


    Private Shared Function CreateHandoutMasterPart(ByVal presentationPart As PresentationPart) As HandoutMasterPart
            Dim handoutMasterPart1 As HandoutMasterPart = presentationPart.AddNewPart(Of HandoutMasterPart)("rId3")
            handoutMasterPart1.HandoutMaster = New HandoutMaster(New CommonSlideData(New ShapeTree(New _
                P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(1UI, UInt32Value), _
             .Name = "" _
            }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New _
                GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
             .Id = DirectCast(2UI, UInt32Value), _
             .Name = "Title 1" _
            }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
             .NoGrouping = True _
            }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New _
                P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties() With { _
             .Language = "en-US" _
            }))))), New P.ColorMap() With { _
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
            })
            Return handoutMasterPart1
        End Function
End Module
