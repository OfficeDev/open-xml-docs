Module Program `
  Sub Main(args As String())`
  End Sub`

  
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXML.Drawing 
    Imports DocumentFormat.OpenXml.Presentation
    Imports P = DocumentFormat.OpenXml.Presentation
    Imports D = DocumentFormat.OpenXml.Drawing

    Imports System.Collections.Generic
    Imports System.Linq
    Imports System.Text
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Drawing
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Presentation
    Imports P = DocumentFormat.OpenXml.Presentation
    Imports D = DocumentFormat.OpenXml.Drawing


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

                'Close the presentation handle
                presentationDoc.Close()
            End Sub

            Private Shared Sub CreatePresentationParts(ByVal presentationPart As PresentationPart)
                Dim slideMasterIdList1 As New SlideMasterIdList(New SlideMasterId() With { _
                 .Id = CType(2147483648UI, UInt32Value), _
                 .RelationshipId = "rId1" _
                })
                Dim slideIdList1 As New SlideIdList(New SlideId() With { _
                 .Id = CType(256UI, UInt32Value), .RelationshipId = "rId2" _
                })
                Dim slideSize1 As New SlideSize() With { _
                 .Cx = 9144000, _
                 .Cy = 6858000, _
                 .Type = SlideSizeValues.Screen4x3 _
                }
                Dim notesSize1 As New NotesSize() With { _
                 .Cx = 6858000, _
                 .Cy = 9144000 _
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
                slidePart1.Slide = New Slide(New CommonSlideData(New ShapeTree(New P.NonVisualGroupShapeProperties(New P.NonVisualDrawingProperties() With { _
                 .Id = CType(1UI, UInt32Value), _
                  .Name = "" _
                }, New P.NonVisualGroupShapeDrawingProperties(), New ApplicationNonVisualDrawingProperties()), New GroupShapeProperties(New TransformGroup()), New P.Shape(New P.NonVisualShapeProperties(New P.NonVisualDrawingProperties() With { _
                  .Id = CType(2UI, UInt32Value), _
                  .Name = "Title 1" _
                }, New P.NonVisualShapeDrawingProperties(New ShapeLocks() With { _
                  .NoGrouping = True _
                }), New ApplicationNonVisualDrawingProperties(New PlaceholderShape())), New P.ShapeProperties(), New P.TextBody(New BodyProperties(), New ListStyle(), New Paragraph(New EndParagraphRunProperties() With { _
                  .Language = "en-US" _
                }))))), New ColorMapOverride(New MasterColorMapping()))
                Return slidePart1
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

            Private Shared Function CreateTheme(ByVal slideMasterPart1 As SlideMasterPart) As ThemePart
                Dim themePart1 As ThemePart = slideMasterPart1.AddNewPart(Of ThemePart)("rId5")
                Dim theme1 As New D.Theme() With { _
                  .Name = "Office Theme" _
                }

                Dim themeElements1 As New D.ThemeElements(New D.ColorScheme(New D.Dark1Color(New D.SystemColor() With { _
                  .Val = D.SystemColorValues.WindowText, _
                  .LastColor = "000000" _
                }), New D.Light1Color(New D.SystemColor() With { _
                  .Val = D.SystemColorValues.Window, _
                  .LastColor = "FFFFFF" _
                }), New D.Dark2Color(New D.RgbColorModelHex() With { _
                  .Val = "1F497D" _
                }), New D.Light2Color(New D.RgbColorModelHex() With { _
                  .Val = "EEECE1" _
                }), New D.Accent1Color(New D.RgbColorModelHex() With { _
                  .Val = "4F81BD" _
                }), New D.Accent2Color(New D.RgbColorModelHex() With { _
                  .Val = "C0504D" _
                }), _
                 New D.Accent3Color(New D.RgbColorModelHex() With { _
                  .Val = "9BBB59" _
                }), New D.Accent4Color(New D.RgbColorModelHex() With { _
                  .Val = "8064A2" _
                }), New D.Accent5Color(New D.RgbColorModelHex() With { _
                  .Val = "4BACC6" _
                }), New D.Accent6Color(New D.RgbColorModelHex() With { _
                  .Val = "F79646" _
                }), New D.Hyperlink(New D.RgbColorModelHex() With { _
                  .Val = "0000FF" _
                }), New D.FollowedHyperlinkColor(New D.RgbColorModelHex() With { _
                  .Val = "800080" _
                })) With { _
                  .Name = "Office" _
                }, New D.FontScheme(New D.MajorFont(New D.LatinFont() With { _
                  .Typeface = "Calibri" _
                }, New D.EastAsianFont() With { _
                  .Typeface = "" _
                }, New D.ComplexScriptFont() With { _
                  .Typeface = "" _
                }), New D.MinorFont(New D.LatinFont() With { _
                  .Typeface = "Calibri" _
                }, New D.EastAsianFont() With { _
                  .Typeface = "" _
                }, New D.ComplexScriptFont() With { _
                  .Typeface = "" _
                })) With { _
                  .Name = "Office" _
                }, New D.FormatScheme(New D.FillStyleList(New D.SolidFill(New D.SchemeColor() With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 37000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 35000 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 15000 _
                }, New D.SaturationModulation() With { _
                  .Val = 350000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 100000 _
                }), New D.LinearGradientFill() With { _
                  .Angle = 16200000, _
                  .Scaled = True _
                }), New D.NoFill(), New D.PatternFill(), New D.GroupFill()), New D.LineStyleList(New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With { _
                  .Val = 95000 _
                }, New D.SaturationModulation() With { _
                  .Val = 105000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.PresetDash() With { _
                  .Val = D.PresetLineDashValues.Solid _
                }) With { _
                  .Width = 9525, _
                  .CapType = D.LineCapValues.Flat, _
                  .CompoundLineType = D.CompoundLineValues.[Single], _
                  .Alignment = D.PenAlignmentValues.Center _
                }, New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With { _
                  .Val = 95000 _
                }, New D.SaturationModulation() With { _
                  .Val = 105000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.PresetDash() With { _
                  .Val = D.PresetLineDashValues.Solid _
                }) With { _
                  .Width = 9525, _
                  .CapType = D.LineCapValues.Flat, _
                  .CompoundLineType = D.CompoundLineValues.[Single], _
                  .Alignment = D.PenAlignmentValues.Center _
                }, New D.Outline(New D.SolidFill(New D.SchemeColor(New D.Shade() With { _
                  .Val = 95000 _
                }, New D.SaturationModulation() With { _
                  .Val = 105000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.PresetDash() With { _
                  .Val = D.PresetLineDashValues.Solid _
                }) With { _
                  .Width = 9525, _
                  .CapType = D.LineCapValues.Flat, _
                  .CompoundLineType = D.CompoundLineValues.[Single], _
                  .Alignment = D.PenAlignmentValues.Center _
                }), New D.EffectStyleList(New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With { _
                  .Val = 38000 _
                }) With { _
                  .Val = "000000" _
                }) With { _
                  .BlurRadius = 40000L, _
                  .Distance = 20000L, _
                  .Direction = 5400000, _
                  .RotateWithShape = False _
                })), New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With { _
                  .Val = 38000 _
                }) With { _
                  .Val = "000000" _
                }) With { _
                  .BlurRadius = 40000L, _
                  .Distance = 20000L, _
                  .Direction = 5400000, _
                  .RotateWithShape = False _
                })), New D.EffectStyle(New D.EffectList(New D.OuterShadow(New D.RgbColorModelHex(New D.Alpha() With { _
                  .Val = 38000 _
                }) With { _
                  .Val = "000000" _
                }) With { _
                  .BlurRadius = 40000L, _
                  .Distance = 20000L, _
                  .Direction = 5400000, _
                  .RotateWithShape = False _
                }))), New D.BackgroundFillStyleList(New D.SolidFill(New D.SchemeColor() With { _
                  .Val = D.SchemeColorValues.PhColor _
                }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }), New D.LinearGradientFill() With { _
                  .Angle = 16200000, _
                  .Scaled = True _
                }), New D.GradientFill(New D.GradientStopList(New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }, New D.GradientStop(New D.SchemeColor(New D.Tint() With { _
                  .Val = 50000 _
                }, New D.SaturationModulation() With { _
                  .Val = 300000 _
                }) With { _
                  .Val = D.SchemeColorValues.PhColor _
                }) With { _
                  .Position = 0 _
                }), New D.LinearGradientFill() With { _
                  .Angle = 16200000, _
                  .Scaled = True _
                }))) With { _
                  .Name = "Office" _
                })

                theme1.Append(themeElements1)
                theme1.Append(New D.ObjectDefaults())
                theme1.Append(New D.ExtraColorSchemeList())

                themePart1.Theme = theme1
                Return themePart1

            End Function

        End Class

    End Namespace
End Module