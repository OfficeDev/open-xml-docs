' <Snippet9>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation


Module MyModule
    ' <Snippet8>
    Sub Main(args As String())
        ApplyThemeToPresentation(args(0), args(1))
    End Sub
    ' </Snippet8>
    ' <Snippet2>
    ' Apply a new theme to the presentation. 
    Public Sub ApplyThemeToPresentation(ByVal presentationFile As String, ByVal themePresentation As String)
        ' <Snippet1>
        Dim themeDocument As PresentationDocument = PresentationDocument.Open(themePresentation, False)
        Dim presentationDoc As PresentationDocument = PresentationDocument.Open(presentationFile, True)
        Using (themeDocument)
            Using (presentationDoc)
                ' </Snippet1>
                ApplyThemeToPresentation(presentationDoc, themeDocument)
            End Using
        End Using

    End Sub
    ' </Snippet2>
    ' <Snippet3>
    ' Apply a new theme to the presentation. 
    Public Sub ApplyThemeToPresentation(ByVal presentationDocument As PresentationDocument, ByVal themeDocument As PresentationDocument)
        If (presentationDocument Is Nothing) Then
            Throw New ArgumentNullException("presentationDocument")
        End If
        If (themeDocument Is Nothing) Then
            Throw New ArgumentNullException("themeDocument")
        End If

        ' Get the presentation part of the presentation document.
        Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

        ' Get the existing slide master part.
        Dim slideMasterPart As SlideMasterPart = presentationPart.SlideMasterParts.ElementAt(0)

        Dim relationshipId As String = presentationPart.GetIdOfPart(slideMasterPart)

        ' Get the new slide master part.
        Dim newSlideMasterPart As SlideMasterPart = themeDocument.PresentationPart.SlideMasterParts.ElementAt(0)
        ' </Snippet3>
        ' <Snippet4>
        ' Remove the theme part.
        presentationPart.DeletePart(presentationPart.ThemePart)

        ' Remove the old slide master part.
        presentationPart.DeletePart(slideMasterPart)

        ' Import the new slide master part, and reuse the old relationship ID.
        newSlideMasterPart = presentationPart.AddPart(newSlideMasterPart, relationshipId)

        ' Change to the new theme part.
        presentationPart.AddPart(newSlideMasterPart.ThemePart)
        ' </Snippet4>
        ' <Snippet5>
        ' <Snippet6>
        Dim newSlideLayouts As Dictionary(Of String, SlideLayoutPart) = New Dictionary(Of String, SlideLayoutPart)()
        For Each slideLayoutPart As Object In newSlideMasterPart.SlideLayoutParts
            newSlideLayouts.Add(GetSlideLayoutType(slideLayoutPart), slideLayoutPart)
        Next
        Dim layoutType As String = Nothing
        Dim newLayoutPart As SlideLayoutPart = Nothing

        ' Insert the code for the layout for this example.
        Dim defaultLayoutType As String = "Title and Content"
        ' </Snippet5>
        ' Remove the slide layout relationship on all slides. 
        For Each slidePart As Object In presentationPart.SlideParts
            layoutType = Nothing
            If ((slidePart.SlideLayoutPart) IsNot Nothing) Then

                ' Determine the slide layout type for each slide.
                layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart)

                ' Delete the old layout part.
                slidePart.DeletePart(slidePart.SlideLayoutPart)
            End If

            If (((layoutType) IsNot Nothing) AndAlso newSlideLayouts.TryGetValue(layoutType, newLayoutPart)) Then

                ' Apply the new layout part.
                slidePart.AddPart(newLayoutPart)
            Else
                newLayoutPart = newSlideLayouts(defaultLayoutType)

                ' Apply the new default layout part.
                slidePart.AddPart(newLayoutPart)
            End If
            ' </Snippet6>
        Next
    End Sub
    ' <Snippet7>
    ' Get the type of the slide layout.
    Public Function GetSlideLayoutType(ByVal slideLayoutPart As SlideLayoutPart) As String
        Dim slideData As CommonSlideData = slideLayoutPart.SlideLayout.CommonSlideData

        ' Remarks: If this is used in production code, check for a null reference.
        Return slideData.Name
    End Function
    ' </Snippet7>
End Module
' </Snippet9>