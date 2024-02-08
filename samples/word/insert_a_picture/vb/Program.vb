' <Snippet>
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports A = DocumentFormat.OpenXml.Drawing
Imports DW = DocumentFormat.OpenXml.Drawing.Wordprocessing
Imports PIC = DocumentFormat.OpenXml.Drawing.Pictures

Module Program
    Sub Main(args As String())
        ' <Snippet4>
        Dim documentPath As String = args(0)
        Dim picturePath As String = args(1)

        InsertAPicture(documentPath, picturePath)
        ' </Snippet4>
    End Sub



    Public Sub InsertAPicture(ByVal document As String, ByVal fileName As String)
        ' <Snippet1>
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            ' </Snippet1>

            ' <Snippet2>
            Dim mainPart As MainDocumentPart = wordprocessingDocument.MainDocumentPart

            Dim imagePart As ImagePart = mainPart.AddImagePart(ImagePartType.Jpeg)

            Using stream As New FileStream(fileName, FileMode.Open)
                imagePart.FeedData(stream)
            End Using

            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart))
            ' </Snippet2>
        End Using
    End Sub

    Private Sub AddImageToBody(ByVal wordDoc As WordprocessingDocument, ByVal relationshipId As String)
        ' <Snippet3>
        ' Define the reference of the image.
        Dim element = New Drawing(
                              New DW.Inline(
                          New DW.Extent() With {.Cx = 990000L, .Cy = 792000L},
                          New DW.EffectExtent() With {.LeftEdge = 0L, .TopEdge = 0L, .RightEdge = 0L, .BottomEdge = 0L},
                          New DW.DocProperties() With {.Id = CType(1UI, UInt32Value), .Name = "Picture1"},
                          New DW.NonVisualGraphicFrameDrawingProperties(
                              New A.GraphicFrameLocks() With {.NoChangeAspect = True}
                              ),
                          New A.Graphic(New A.GraphicData(
                                        New PIC.Picture(
                                            New PIC.NonVisualPictureProperties(
                                                New PIC.NonVisualDrawingProperties() With {.Id = 0UI, .Name = "Koala.jpg"},
                                                New PIC.NonVisualPictureDrawingProperties()
                                                ),
                                            New PIC.BlipFill(
                                                New A.Blip(
                                                    New A.BlipExtensionList(
                                                        New A.BlipExtension() With {.Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"})
                                                    ) With {.Embed = relationshipId, .CompressionState = A.BlipCompressionValues.Print},
                                                New A.Stretch(
                                                    New A.FillRectangle()
                                                    )
                                                ),
                                            New PIC.ShapeProperties(
                                                New A.Transform2D(
                                                    New A.Offset() With {.X = 0L, .Y = 0L},
                                                    New A.Extents() With {.Cx = 990000L, .Cy = 792000L}),
                                                New A.PresetGeometry(
                                                    New A.AdjustValueList()
                                                    ) With {.Preset = A.ShapeTypeValues.Rectangle}
                                                )
                                            )
                                        ) With {.Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"}
                                    )
                                ) With {.DistanceFromTop = 0UI,
                                        .DistanceFromBottom = 0UI,
                                        .DistanceFromLeft = 0UI,
                                        .DistanceFromRight = 0UI}
                            )

        ' Append the reference to body, the element should be in a Run.
        wordDoc.MainDocumentPart.Document.Body.AppendChild(New Paragraph(New Run(element)))
        ' </Snippet3>
    End Sub
End Module
' </Snippet>
