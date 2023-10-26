Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports D = DocumentFormat.OpenXml.Drawing
Imports P = DocumentFormat.OpenXml.Presentation

Namespace CreatePresentationDocument
    Class Program
        Public Shared Sub Main(ByVal args As String())

            Dim filepath As String = "C:\Users\username\Documents\PresentationFromFilename.pptx"
            CreatePresentation(filepath)

        End Sub

        Public Shared Sub CreatePresentation(ByVal filepath As String)
            ' Create a presentation at a specified file path. The presentation document type is pptx, by default.
            Dim presentationDoc As PresentationDocument = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)
            Dim presentationPart As PresentationPart = presentationDoc.AddPresentationPart()
            presentationPart.Presentation = New Presentation()

            CreatePresentationParts(presentationPart)

            'Dispose the presentation handle
            presentationDoc.Dispose()
        End Sub

        Private Shared Sub CreatePresentationParts(ByVal presentationPart As PresentationPart)
            Dim slideMasterIdList1 As New SlideMasterIdList(New SlideMasterId() With {
             .Id = CType(2147483648UI, UInt32Value),
             .RelationshipId = "rId1"
            })
            Dim slideIdList1 As New SlideIdList(New SlideId() With {
             .Id = CType(256UI, UInt32Value), .RelationshipId = "rId2"
            })
            Dim slideSize1 As New SlideSize() With {
             .Cx = 9144000,
             .Cy = 6858000,
             .Type = SlideSizeValues.Screen4x3
            }
            Dim notesSize1 As New NotesSize() With {
             .Cx = 6858000,
             .Cy = 9144000
            }
            Dim defaultTextStyle1 As New DefaultTextStyle()

            Dim slidePart1 As SlidePart
            Dim slideLayoutPart1 As SlideLayoutPart
            Dim slideMasterPart1 As SlideMasterPart
            Dim themePart1 As ThemePart

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1)

            slidePart1 = CreateSlidePart(presentationPart)
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1)
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1)
            themePart1 = CreateTheme(slideMasterPart1)

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1")
            presentationPart.AddPart(slideMasterPart1, "rId1")
            presentationPart.AddPart(themePart1, "rId5")
        End Sub

        Private Shared Function CreateSlidePart(ByVal presentationPart As PresentationPart) As SlidePart
            Dim slidePart1 As SlidePart = presentationPart.AddNewPart(Of SlidePart)("rId2")
            slidePart1.Slide = New Slide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With {
             .Id = CType(1UI, UInt32Value),
              .Name = ""
            }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With {
              .Id = CType(2UI, UInt32Value),
              .Name = "Title 1"
            }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {
              .NoGrouping = True
            }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties() With {
              .Language = "en-US"
            }))))), New ColorMapOverride(New MasterColorMapping()))
            Return slidePart1
        End Function

        Private Shared Function CreateSlideLayoutPart(ByVal slidePart1 As SlidePart) As SlideLayoutPart
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

        Private Shared Function CreateSlideMasterPart(ByVal slideLayoutPart1 As SlideLayoutPart) As SlideMasterPart
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

        Private Shared Function CreateTheme(ByVal slideMasterPart1 As SlideMasterPart) As ThemePart
            Dim themePart1 As ThemePart = slideMasterPart1.AddNewPart(Of ThemePart)("rId5")
            Dim theme1 As New D.Theme() With {
              .Name = "Office Theme"
            }

            Dim themeElements1 As New D.ThemeElements(New D.ColorScheme(New D.Dark1Color(New D.SystemColor() With {
              .Val = D.SystemColorValues.WindowText,
              .LastColor = "000000"
            }), New D.Light1Color(New D.SystemColor() With {
              .Val = D.SystemColorValues.Window,
              .LastColor = "FFFFFF"
            }), New D.Dark2Color(New D.RgbColorModelHex() With {
              .Val = "1F497D"
            }), New D.Light2Color(New D.RgbColorModelHex() With {
              .Val = "EEECE1"
            }), New D.Accent1Color(New D.RgbColorModelHex() With {
              .Val = "4F81BD"
            }), New D.Accent2Color(New D.RgbColorModelHex() With {
              .Val = "C0504D"
            }),
             New D.Accent3Color(New D.RgbColorModelHex() With {
              .Val = "9BBB59"
            }), New D.Accent4Color(New D.RgbColorModelHex() With {
              .Val = "8064A2"
            }), New D.Accent5Color(New D.RgbColorModelHex() With {
              .Val = "4BACC6"
            }), New D.Accent6Color(New D.RgbColorModelHex() With {
              .Val = "F79646"
            }), New D.Hyperlink(New D.RgbColorModelHex() With {
              .Val = "0000FF"
            }), New D.FollowedHyperlinkColor(New D.RgbColorModelHex() With {
              .Val = "800080"
            })) With {
              .Name = "Office"
            }, New D.FontScheme(New D.MajorFont(New D.LatinFont() With {
              .Typeface = "Calibri"
            }, New D.EastAsianFont() With {
              .Typeface = ""
            }, New D.ComplexScriptFont() With {
              .Typeface = ""
            }), New D.MinorFont(New D.LatinFont() With {
              .Typeface = "Calibri"
            }, New D.EastAsianFont() With {
              .Typeface = ""
            }, New D.ComplexScriptFont() With {
              .Typeface = ""
            })) With {
              .Name = "Office"
            }, New D.FormatScheme(New D.FillStyleList(New D.SolidFill(New D.SchemeColor() With {
              .Val = D.SchemeColorValues.PhColor
            }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With {
              .Val = 50000
            }, New D.SaturationModulation() With {
              .Val = 300000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }) With {
              .Position = 0
            }, New D.GradientStop(New D.SchemeColor(New D.Tint() With {
              .Val = 37000
            }, New D.SaturationModulation() With {
              .Val = 300000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }) With {
              .Position = 35000
            }, New D.GradientStop(New D.SchemeColor(New D.Tint() With {
              .Val = 15000
            }, New D.SaturationModulation() With {
              .Val = 350000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }) With {
              .Position = 100000
            }), New D.LinearGradientFill() With {
              .Angle = 16200000,
              .Scaled = True
            }), New D.NoFill(), New D.PatternFill(), New D.GroupFill()), New D.LineStyleList(New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With {
              .Val = 95000
            }, New D.SaturationModulation() With {
              .Val = 105000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }), New D.PresetDash() With {
              .Val = D.PresetLineDashValues.Solid
            }) With {
              .Width = 9525,
              .CapType = D.LineCapValues.Flat,
              .CompoundLineType = D.CompoundLineValues.[Single],
              .Alignment = D.PenAlignmentValues.Center
            }, New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With {
              .Val = 95000
            }, New D.SaturationModulation() With {
              .Val = 105000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }), New D.PresetDash() With {
              .Val = D.PresetLineDashValues.Solid
            }) With {
              .Width = 9525,
              .CapType = D.LineCapValues.Flat,
              .CompoundLineType = D.CompoundLineValues.[Single],
              .Alignment = D.PenAlignmentValues.Center
            }, New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With {
              .Val = 95000
            }, New D.SaturationModulation() With {
              .Val = 105000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }), New D.PresetDash() With {
              .Val = D.PresetLineDashValues.Solid
            }) With {
              .Width = 9525,
              .CapType = D.LineCapValues.Flat,
              .CompoundLineType = D.CompoundLineValues.[Single],
              .Alignment = D.PenAlignmentValues.Center
            }), New D.EffectStyleList(New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With {
              .Val = 38000
            }) With {
              .Val = "000000"
            }) With {
              .BlurRadius = 40000L,
              .Distance = 20000L,
              .Direction = 5400000,
              .RotateWithShape = False
            })), New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With {
              .Val = 38000
            }) With {
              .Val = "000000"
            }) With {
              .BlurRadius = 40000L,
              .Distance = 20000L,
              .Direction = 5400000,
              .RotateWithShape = False
            })), New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With {
              .Val = 38000
            }) With {
              .Val = "000000"
            }) With {
              .BlurRadius = 40000L,
              .Distance = 20000L,
              .Direction = 5400000,
              .RotateWithShape = False
            }))), New D.BackgroundFillStyleList(New D.SolidFill(New D.SchemeColor() With {
              .Val = D.SchemeColorValues.PhColor
            }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With {
              .Val = 50000
            }, New D.SaturationModulation() With {
              .Val = 300000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }) With {
              .Position = 0
            }, New D.GradientStop(New D.SchemeColor(New D.Tint() With {
              .Val = 50000
            }, New D.SaturationModulation() With {
              .Val = 300000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }) With {
              .Position = 0
            }, New D.GradientStop(New D.SchemeColor(New D.Tint() With {
              .Val = 50000
            }, New D.SaturationModulation() With {
              .Val = 300000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }) With {
              .Position = 0
            }), New D.LinearGradientFill() With {
              .Angle = 16200000,
              .Scaled = True
            }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With {
              .Val = 50000
            }, New D.SaturationModulation() With {
              .Val = 300000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }) With {
              .Position = 0
            }, New D.GradientStop(New D.SchemeColor(New D.Tint() With {
              .Val = 50000
            }, New D.SaturationModulation() With {
              .Val = 300000
            }) With {
              .Val = D.SchemeColorValues.PhColor
            }) With {
              .Position = 0
            }), New D.LinearGradientFill() With {
              .Angle = 16200000,
              .Scaled = True
            }))) With {
              .Name = "Office"
            })

            theme1.Append(themeElements1)
            theme1.Append(New D.ObjectDefaults())
            theme1.Append(New D.ExtraColorSchemeList())

            themePart1.Theme = theme1
            Return themePart1

        End Function

    End Class

End Namespace
