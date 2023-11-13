Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports Drawing = DocumentFormat.OpenXml.Drawing


Module MyModule

    Sub Main(args As String())
    End Sub

    ' Insert a slide into the specified presentation.
    Public Sub InsertNewSlide(ByVal presentationFile As String, ByVal position As Integer, ByVal slideTitle As String)

        ' Open the source document as read/write. 
        Dim presentationDocument As PresentationDocument = presentationDocument.Open(presentationFile, True)

        Using (presentationDocument)

            'Pass the source document and the position and title of the slide to be inserted to the next method.
            InsertNewSlide(presentationDocument, position, slideTitle)

        End Using

    End Sub
    ' Insert a slide into the specified presentation.
    Public Sub InsertNewSlide(ByVal presentationDocument As PresentationDocument, ByVal position As Integer, ByVal slideTitle As String)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If
        If (slideTitle Is Nothing) Then
            Throw New ArgumentNullException("slideTitle")
        End If

        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Verify that the presentation is not empty.
        If (presentationPart Is Nothing) Then
            Throw New InvalidOperationException("The presentation document is empty.")
        End If

        ' Declare and instantiate a new slide.
        Dim slide As Slide = New Slide(New CommonSlideData(New ShapeTree))
        Dim drawingObjectId As UInteger = 1

        ' Construct the slide content.
        ' Specify the non-visual properties of the new slide.
        Dim nonVisualProperties As DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties = slide.CommonSlideData.ShapeTree.AppendChild(New _
            DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties())
        nonVisualProperties.NonVisualDrawingProperties = New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 1, .Name = ""}
        nonVisualProperties.NonVisualGroupShapeDrawingProperties = New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties()
        nonVisualProperties.ApplicationNonVisualDrawingProperties = New ApplicationNonVisualDrawingProperties()

        ' Specify the group shape properties of the new slide.
        slide.CommonSlideData.ShapeTree.AppendChild(New DocumentFormat.OpenXml.Presentation.GroupShapeProperties())
        ' Declare and instantiate the title shape of the new slide.
        Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = slide.CommonSlideData.ShapeTree.AppendChild _
            (New DocumentFormat.OpenXml.Presentation.Shape())
        drawingObjectId = (drawingObjectId + 1)

        ' Specify the required shape properties for the title shape. 
        titleShape.NonVisualShapeProperties = New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(New _
            DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "Title"}, _
            New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties _
            (New Drawing.ShapeLocks() With {.NoGrouping = True}), _
            New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Type = PlaceholderValues.Title}))

        titleShape.ShapeProperties = New DocumentFormat.OpenXml.Presentation.ShapeProperties()

        ' Specify the text of the title shape.
        titleShape.TextBody = New DocumentFormat.OpenXml.Presentation.TextBody(New Drawing.BodyProperties, _
             New Drawing.ListStyle, New Drawing.Paragraph _
             (New Drawing.Run(New Drawing.Text() With {.Text = slideTitle})))
        ' Declare and instantiate the body shape of the new slide.
        Dim bodyShape As DocumentFormat.OpenXml.Presentation.Shape = slide.CommonSlideData.ShapeTree.AppendChild _
            (New DocumentFormat.OpenXml.Presentation.Shape())
        drawingObjectId = (drawingObjectId + 1)

        ' Specify the required shape properties for the body shape.
        bodyShape.NonVisualShapeProperties = New NonVisualShapeProperties(New NonVisualDrawingProperties() With {.Id = drawingObjectId, .Name = "ContentPlaceholder"}, _
             New NonVisualShapeDrawingProperties(New Drawing.ShapeLocks() With {.NoGrouping = True}), _
             New ApplicationNonVisualDrawingProperties(New PlaceholderShape() With {.Index = 1}))

        bodyShape.ShapeProperties = New ShapeProperties()

        ' Specify the text of the body shape.
        bodyShape.TextBody = New TextBody(New Drawing.BodyProperties, New Drawing.ListStyle, New Drawing.Paragraph)
        ' Create the slide part for the new slide.
        Dim slidePart As SlidePart = presentationPart.AddNewPart(Of SlidePart)()

        ' Save the new slide part.
        slide.Save(slidePart)

        ' Modify the slide ID list in the presentation part.
        ' The slide ID list should not be null.
        Dim slideIdList As SlideIdList = presentationPart.Presentation.SlideIdList

        ' Find the highest slide ID in the current list.
        Dim maxSlideId As UInteger = 1
        Dim prevSlideId As SlideId = Nothing

        For Each slideId As SlideId In slideIdList.ChildElements
            If (CType(slideId.Id, UInteger) > maxSlideId) Then
                maxSlideId = slideId.Id
            End If
            position = (position - 1)
            If (position = 0) Then
                prevSlideId = slideId
            End If
        Next

        maxSlideId = (maxSlideId + 1)

        ' Get the ID of the previous slide.
        Dim lastSlidePart As SlidePart = Nothing

        If (prevSlideId IsNot Nothing) Then
            lastSlidePart = CType(presentationPart.GetPartById(prevSlideId.RelationshipId), SlidePart)
        Else
            lastSlidePart = CType(presentationPart.GetPartById(CType(slideIdList.ChildElements(0), SlideId).RelationshipId), SlidePart)
        End If


        ' Use the same slide layout as that of the previous slide.
        If ((lastSlidePart.SlideLayoutPart) IsNot Nothing) Then
            slidePart.AddPart(lastSlidePart.SlideLayoutPart)
        End If

        ' Insert the new slide into the slide list after the previous slide.
        Dim newSlideId As SlideId = slideIdList.InsertAfter(New SlideId, prevSlideId)
        newSlideId.Id = maxSlideId
        newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart)

        ' Save the modified presentation.
        presentationPart.Presentation.Save()

    End Sub
End Module