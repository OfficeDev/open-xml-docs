Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports D = DocumentFormat.OpenXml.Drawing
Imports P = DocumentFormat.OpenXml.Presentation

Module Program
    Sub Main(args As String())
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(args(0), True)
            Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

            If presentationPart IsNot Nothing Then
                CreateHandoutMasterPart(presentationPart)
            End If
        End Using
    End Sub

    ' <Snippet0>
    Private Function CreateHandoutMasterPart(presentationPart As PresentationPart) As HandoutMasterPart
        Dim handoutMasterPart1 As HandoutMasterPart = presentationPart.AddNewPart(Of HandoutMasterPart)()
        handoutMasterPart1.HandoutMaster = New HandoutMaster(
            New CommonSlideData(
                New ShapeTree(
                    New P.NonVisualGroupShapeProperties(
                        New P.NonVisualDrawingProperties() With {.Id = CType(1UI, UInt32Value), .Name = ""},
                        New P.NonVisualGroupShapeDrawingProperties(),
                        New ApplicationNonVisualDrawingProperties()
                    ),
                    New GroupShapeProperties(New TransformGroup()),
                    New P.Shape(
                        New P.NonVisualShapeProperties(
                            New P.NonVisualDrawingProperties() With {.Id = CType(2UI, UInt32Value), .Name = "Title 1"},
                            New P.NonVisualShapeDrawingProperties(New ShapeLocks() With {.NoGrouping = True}),
                            New ApplicationNonVisualDrawingProperties(New PlaceholderShape())
                        ),
                        New P.ShapeProperties(),
                        New P.TextBody(
                            New BodyProperties(),
                            New ListStyle(),
                            New Paragraph(New EndParagraphRunProperties() With {.Language = "en-US"})
                        )
                    )
                )
            ),
            New P.ColorMap() With {
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
            }
        )

        Return handoutMasterPart1
    End Function
    ' </Snippet0>
End Module
