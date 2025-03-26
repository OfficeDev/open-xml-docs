Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Presentation
Imports A = DocumentFormat.OpenXml.Drawing
Imports P14 = DocumentFormat.OpenXml.Office2010.PowerPoint
Imports ShapeTree = DocumentFormat.OpenXml.Presentation.ShapeTree
Imports ShapeProperties = DocumentFormat.OpenXml.Presentation.ShapeProperties
Imports NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
Imports NonVisualPictureProperties = DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties
Imports NonVisualPictureDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties
Imports Picture = DocumentFormat.OpenXml.Presentation.Picture
Imports BlipFill = DocumentFormat.OpenXml.Presentation.BlipFill
Imports DocumentFormat.OpenXml.Packaging
Imports ApplicationNonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties
Imports System.IO

Module Program
    Sub Main(args As String())
        AddVideo(args(0), args(1), args(2))
    End Sub

    Sub AddVideo(filePath As String, videoFilePath As String, coverPicPath As String)
        Dim imgEmbedId As String = "rId4"
        Dim embedId As String = "rId3"
        Dim mediaEmbedId As String = "rId2"
        Dim shapeId As UInt32Value = 5
        ' <Snippet1>
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(filePath, True)
            ' </Snippet1>
            If presentationDocument.PresentationPart Is Nothing OrElse presentationDocument.PresentationPart.Presentation.SlideIdList Is Nothing Then
                Throw New NullReferenceException("Presentation Part is empty or there are no slides in it")
            End If

            ' Get presentation part
            Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

            ' Get slides ids
            Dim slidesIds As OpenXmlElementList = presentationPart.Presentation.SlideIdList.ChildElements

            ' Get relationshipId of the last slide
            Dim videoSldRelationshipId As String = CType(slidesIds(0), SlideId).RelationshipId

            If videoSldRelationshipId Is Nothing Then
                Throw New NullReferenceException("Slide id not found")
            End If

            ' Get slide part by relationshipID
            Dim slidePart As SlidePart = CType(presentationPart.GetPartById(videoSldRelationshipId), SlidePart)

            ' Create video Media Data Part (content type, extension)
            Dim mediaDataPart As MediaDataPart = presentationDocument.CreateMediaDataPart("video/mp4", ".mp4")

            ' Get the video file and feed the stream
            Using mediaDataPartStream As Stream = File.OpenRead(videoFilePath)
                mediaDataPart.FeedData(mediaDataPartStream)
            End Using

            ' Adds a VideoReferenceRelationship to the MainDocumentPart
            slidePart.AddVideoReferenceRelationship(mediaDataPart, embedId)

            ' Adds a MediaReferenceRelationship to the SlideLayoutPart
            slidePart.AddMediaReferenceRelationship(mediaDataPart, mediaEmbedId)

            Dim nonVisualDrawingProperties As New NonVisualDrawingProperties() With {
                .Id = shapeId,
                .Name = "video"
            }
            Dim videoFromFile As New A.VideoFromFile() With {
                .Link = embedId
            }

            Dim appNonVisualDrawingProperties As New ApplicationNonVisualDrawingProperties()
            appNonVisualDrawingProperties.Append(videoFromFile)

            ' Adds sample image to the slide with id to be used as reference in blip
            Dim imagePart As ImagePart = slidePart.AddImagePart(ImagePartType.Png, imgEmbedId)
            Using data As Stream = File.OpenRead(coverPicPath)
                imagePart.FeedData(data)
            End Using

            If slidePart.Slide.CommonSlideData.ShapeTree Is Nothing Then
                Throw New NullReferenceException("Presentation shape tree is empty")
            End If

            ' Getting existing shape tree object from PowerPoint
            Dim shapeTree As ShapeTree = slidePart.Slide.CommonSlideData.ShapeTree

            ' Specifies the existence of a picture within a presentation
            Dim picture As New Picture()
            Dim nonVisualPictureProperties As New NonVisualPictureProperties()

            Dim hyperlinkOnClick As New A.HyperlinkOnClick() With {
                .Id = "",
                .Action = "ppaction://media"
            }
            nonVisualDrawingProperties.Append(hyperlinkOnClick)

            Dim nonVisualPictureDrawingProperties As New NonVisualPictureDrawingProperties()
            Dim pictureLocks As New A.PictureLocks() With {
                .NoChangeAspect = True
            }
            nonVisualPictureDrawingProperties.Append(pictureLocks)

            Dim appNonVisualDrawingPropertiesExtensionList As New ApplicationNonVisualDrawingPropertiesExtensionList()
            Dim appNonVisualDrawingPropertiesExtension As New ApplicationNonVisualDrawingPropertiesExtension() With {
                .Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"
            }

            Dim media As New P14.Media() With {
                .Embed = mediaEmbedId
            }
            media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main")

            appNonVisualDrawingPropertiesExtension.Append(media)
            appNonVisualDrawingPropertiesExtensionList.Append(appNonVisualDrawingPropertiesExtension)
            appNonVisualDrawingProperties.Append(appNonVisualDrawingPropertiesExtensionList)

            nonVisualPictureProperties.Append(nonVisualDrawingProperties)
            nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties)
            nonVisualPictureProperties.Append(appNonVisualDrawingProperties)

            ' Prepare shape properties to display picture
            Dim blipFill As New BlipFill()
            Dim blip As New A.Blip() With {
                .Embed = imgEmbedId
            }

            Dim stretch As New A.Stretch()
            Dim fillRectangle As New A.FillRectangle()
            Dim transform2D As New A.Transform2D()
            Dim offset As New A.Offset() With {
                .X = 1524000L,
                .Y = 857250L
            }
            Dim extents As New A.Extents() With {
                .Cx = 9144000L,
                .Cy = 5143500L
            }
            Dim presetGeometry As New A.PresetGeometry() With {
                .Preset = A.ShapeTypeValues.Rectangle
            }
            Dim adjValueList As New A.AdjustValueList()

            stretch.Append(fillRectangle)
            blipFill.Append(blip)
            blipFill.Append(stretch)
            transform2D.Append(offset)
            transform2D.Append(extents)
            presetGeometry.Append(adjValueList)

            Dim shapeProperties As New ShapeProperties()
            shapeProperties.Append(transform2D)
            shapeProperties.Append(presetGeometry)

            ' Adds all elements to the slide's shape tree
            picture.Append(nonVisualPictureProperties)
            picture.Append(blipFill)
            picture.Append(shapeProperties)

            shapeTree.Append(picture)
        End Using
    End Sub
End Module