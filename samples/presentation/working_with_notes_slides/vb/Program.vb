Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports insert_a_new_slideto_vb.SlideHelpers
Imports D = DocumentFormat.OpenXml.Drawing
Imports P = DocumentFormat.OpenXml.Presentation

Module Program
    Sub Main(args As String())
        Dim pptxPath As String = args(0)
        ' <Snippet0>
        ' <Snippet1>
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(pptxPath, True)
            If presentationDocument Is Nothing Then
                Throw New Exception("Presentation Document does not exist")
            End If

            ' Get the first slide in the presentation or use the InsertNewSlide.InsertNewSlideIntoPresentation helper method to insert a new slide.
            Dim slidePart As SlidePart = If(presentationDocument.PresentationPart?.SlideParts.FirstOrDefault(), InsertNewSlide.InsertNewSlideIntoPresentation(presentationDocument, 1, "my new slide"))

            ' Add a new NoteSlidePart if one does not already exist
            Dim notesSlidePart As NotesSlidePart = If(slidePart.NotesSlidePart, slidePart.AddNewPart(Of NotesSlidePart)())
            ' </Snippet1>

            ' <Snippet2>
            ' Add a NoteSlide to the NoteSlidePart if one does not already exist.
            If notesSlidePart.NotesSlide Is Nothing Then
                notesSlidePart.NotesSlide = New P.NotesSlide(
                    New P.CommonSlideData(
                        New P.ShapeTree(
                            New P.NonVisualGroupShapeProperties(
                                New P.NonVisualDrawingProperties() With {.Id = 1, .Name = ""},
                                New P.NonVisualGroupShapeDrawingProperties(),
                                New P.ApplicationNonVisualDrawingProperties()),
                            New P.GroupShapeProperties(
                                New D.Transform2D(
                                    New D.Offset() With {.X = 0, .Y = 0},
                                    New D.Extents() With {.Cx = 0, .Cy = 0},
                                    New D.ChildOffset() With {.X = 0, .Y = 0},
                                    New D.ChildExtents() With {.Cx = 0, .Cy = 0}))
)))
                ' </Snippet2>
                ' <Snippet3>
                notesSlidePart.NotesSlide.AddChild(New P.Shape(
                                New P.NonVisualShapeProperties(
                                    New P.NonVisualDrawingProperties() With {.Id = 3, .Name = "test Placeholder 3"},
                                    New P.NonVisualShapeDrawingProperties(),
                                    New P.ApplicationNonVisualDrawingProperties()),
                                New P.ShapeProperties(),
                                New P.TextBody(
                                    New D.BodyProperties(),
                                    New D.Paragraph(
                                        New D.Run(
                                            New D.Text("This is a test note!"))))))
            End If

            notesSlidePart.AddPart(slidePart)
            ' </Snippet3>

            ' <Snippet4>
            ' Add the required NotesMasterPart if it is missing
            Dim notesMasterPart As NotesMasterPart = If(notesSlidePart.NotesMasterPart, notesSlidePart.AddNewPart(Of NotesMasterPart)())

            ' Add a NotesMaster to the NotesMasterPart if not present
            If notesMasterPart.NotesMaster Is Nothing Then
                notesMasterPart.NotesMaster = New NotesMaster(
                    New P.CommonSlideData(
                        New P.ShapeTree(
                            New P.NonVisualGroupShapeProperties(
                                New P.NonVisualDrawingProperties() With {.Id = 1, .Name = "tacocat"},
                                New P.NonVisualGroupShapeDrawingProperties(),
                                New P.ApplicationNonVisualDrawingProperties(
                                    New P.PlaceholderShape() With {.Type = PlaceholderValues.Body, .Index = 1})),
                            New P.GroupShapeProperties())),
                    New P.ColorMap() With {
                        .Background1 = D.ColorSchemeIndexValues.Light1,
                        .Background2 = D.ColorSchemeIndexValues.Light2,
                        .Text1 = D.ColorSchemeIndexValues.Dark1,
                        .Text2 = D.ColorSchemeIndexValues.Dark2,
                        .Accent1 = D.ColorSchemeIndexValues.Accent1,
                        .Accent2 = D.ColorSchemeIndexValues.Accent2,
                        .Accent3 = D.ColorSchemeIndexValues.Accent3,
                        .Accent4 = D.ColorSchemeIndexValues.Accent4,
                        .Accent5 = D.ColorSchemeIndexValues.Accent5,
                        .Accent6 = D.ColorSchemeIndexValues.Accent6,
                        .Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                        .FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                    })
            End If

            ' Add a new ThemePart for the NotesMasterPart
            Dim themePart As ThemePart = If(notesMasterPart.ThemePart, notesMasterPart.AddNewPart(Of ThemePart)())

            ' Add the Theme if it is missing
            If themePart.Theme Is Nothing Then
                themePart.Theme = New Theme(
                    New ThemeElements(
                        New ColorScheme(
                            New Dark1Color(
                                New SystemColor() With {.Val = SystemColorValues.WindowText}),
                            New Light1Color(
                                New SystemColor() With {.Val = SystemColorValues.Window}),
                            New Dark2Color(
                                New RgbColorModelHex() With {.Val = "f1d7be"}),
                            New Light2Color(
                                New RgbColorModelHex() With {.Val = "171717"}),
                            New Accent1Color(
                                New RgbColorModelHex() With {.Val = "ea9f7d"}),
                            New Accent2Color(
                                New RgbColorModelHex() With {.Val = "168ecd"}),
                            New Accent3Color(
                                New RgbColorModelHex() With {.Val = "e694db"}),
                            New Accent4Color(
                                New RgbColorModelHex() With {.Val = "f0612a"}),
                            New Accent5Color(
                                New RgbColorModelHex() With {.Val = "5fd46c"}),
                            New Accent6Color(
                                New RgbColorModelHex() With {.Val = "b158d1"}),
                            New D.Hyperlink(
                                New RgbColorModelHex() With {.Val = "699f82"}),
                            New FollowedHyperlinkColor(
                                New RgbColorModelHex() With {.Val = "699f82"})) With {.Name = "Office2"},
                        New D.FontScheme(
                            New MajorFont(
                                New LatinFont(),
                                New EastAsianFont(),
                                New ComplexScriptFont()),
                            New MinorFont(
                                New LatinFont(),
                                New EastAsianFont(),
                                New ComplexScriptFont())) With {.Name = "Office2"},
                        New FormatScheme(
                            New FillStyleList(
                                New NoFill(),
                                New SolidFill(),
                                New D.GradientFill(),
                                New D.BlipFill(),
                                New D.PatternFill(),
                                New GroupFill()),
                            New LineStyleList(
                                New D.Outline(),
                                New D.Outline(),
                                New D.Outline()),
                            New EffectStyleList(
                                New EffectStyle(
                                    New EffectList()),
                                New EffectStyle(
                                    New EffectList()),
                                New EffectStyle(
                                    New EffectList())),
                            New BackgroundFillStyleList(
                                New NoFill(),
                                New SolidFill(),
                                New D.GradientFill(),
                                New D.BlipFill(),
                                New D.PatternFill(),
                                New GroupFill())) With {.Name = "Office2"}),
                    New ObjectDefaults(),
                    New ExtraColorSchemeList())
            End If
            ' </Snippet4>
        End Using
        ' </Snippet0>
    End Sub
End Module
