Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports A = DocumentFormat.OpenXml.Drawing
Imports DW = DocumentFormat.OpenXml.Drawing.Wordprocessing
Imports PIC = DocumentFormat.OpenXml.Drawing.Pictures

Module Program `
  Sub Main(args As String())`
  End Sub`

  
                            
    Public Sub InsertAPicture(ByVal document As String, ByVal fileName As String)
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = wordprocessingDocument.MainDocumentPart

            Dim imagePart As ImagePart = mainPart.AddImagePart(ImagePartType.Jpeg)

            Using stream As New FileStream(fileName, FileMode.Open)
                imagePart.FeedData(stream)
            End Using

            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart))
        End Using
    End Sub

    Private Sub AddImageToBody(ByVal wordDoc As WordprocessingDocument, ByVal relationshipId As String)
        ' Define the reference of the image.
        Dim element = New Drawing( _
                              New DW.Inline( _
                          New DW.Extent() With {.Cx = 990000L, .Cy = 792000L}, _
                          New DW.EffectExtent() With {.LeftEdge = 0L, .TopEdge = 0L, .RightEdge = 0L, .BottomEdge = 0L}, _
                          New DW.DocProperties() With {.Id = CType(1UI, UInt32Value), .Name = "Picture1"}, _
                          New DW.NonVisualGraphicFrameDrawingProperties( _
                              New A.GraphicFrameLocks() With {.NoChangeAspect = True} _
                              ), _
                          New A.Graphic(New A.GraphicData( _
                                        New PIC.Picture( _
                                            New PIC.NonVisualPictureProperties( _
                                                New PIC.NonVisualDrawingProperties() With {.Id = 0UI, .Name = "Koala.jpg"}, _
                                                New PIC.NonVisualPictureDrawingProperties() _
                                                ), _
                                            New PIC.BlipFill( _
                                                New A.Blip( _
                                                    New A.BlipExtensionList( _
                                                        New A.BlipExtension() With {.Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"}) _
                                                    ) With {.Embed = relationshipId, .CompressionState = A.BlipCompressionValues.Print}, _
                                                New A.Stretch( _
                                                    New A.FillRectangle() _
                                                    ) _
                                                ), _
                                            New PIC.ShapeProperties( _
                                                New A.Transform2D( _
                                                    New A.Offset() With {.X = 0L, .Y = 0L}, _
                                                    New A.Extents() With {.Cx = 990000L, .Cy = 792000L}), _
                                                New A.PresetGeometry( _
                                                    New A.AdjustValueList() _
                                                    ) With {.Preset = A.ShapeTypeValues.Rectangle} _
                                                ) _
                                            ) _
                                        ) With {.Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"} _
                                    ) _
                                ) With {.DistanceFromTop = 0UI, _
                                        .DistanceFromBottom = 0UI, _
                                        .DistanceFromLeft = 0UI, _
                                        .DistanceFromRight = 0UI} _
                            )

        ' Append the reference to body, the element should be in a Run.
        wordDoc.MainDocumentPart.Document.Body.AppendChild(New Paragraph(New Run(element)))
    End Sub
End Module