Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation

Module Program
    Sub Main(args As String())
        AddTransitionToSlides(args(0))
    End Sub

    ' <Snippet0>
    Sub AddTransitionToSlides(filePath As String)
        ' <Snippet1>
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(filePath, True)
            ' </Snippet1>
            ' Check if the presentation part and slide list are available
            If presentationDocument.PresentationPart Is Nothing OrElse presentationDocument.PresentationPart.Presentation.SlideIdList Is Nothing Then
                Throw New NullReferenceException("Presentation part is empty or there are no slides")
            End If

            ' Get the presentation part
            Dim presentationPart As PresentationPart = presentationDocument.PresentationPart

            ' Get the list of slide IDs
            Dim slidesIds As OpenXmlElementList = presentationPart.Presentation.SlideIdList.ChildElements

            ' <Snippet2>
            ' Define the transition start time and duration in milliseconds
            Dim startTransitionAfterMs As String = "3000"
            Dim durationMs As String = "2000"

            ' Set to true if you want to advance to the next slide on mouse click
            Dim advanceOnClick As Boolean = True

            ' Iterate through each slide ID to get slides parts
            For Each slideId As SlideId In slidesIds
                ' Get the relationship ID of the slide
                Dim relId As String = slideId.RelationshipId.ToString()

                If relId Is Nothing Then
                    Throw New NullReferenceException("RelationshipId not found")
                End If

                ' Get the slide part using the relationship ID
                Dim slidePart As SlidePart = CType(presentationDocument.PresentationPart.GetPartById(relId), SlidePart)

                ' Remove existing transitions if any
                If slidePart.Slide.Transition IsNot Nothing Then
                    slidePart.Slide.Transition.Remove()
                End If

                ' Check if there are any AlternateContent elements
                If slidePart.Slide.Descendants(Of AlternateContent)().ToList().Count > 0 Then
                    ' Get all AlternateContent elements
                    Dim alternateContents As List(Of AlternateContent) = slidePart.Slide.Descendants(Of AlternateContent)().ToList()
                    For Each alternateContent In alternateContents
                        ' Remove transitions in AlternateContentChoice within AlternateContent
                        Dim childElements As List(Of OpenXmlElement) = alternateContent.ChildElements.ToList()

                        For Each element In childElements
                            Dim transitions As List(Of Transition) = element.Descendants(Of Transition)().ToList()
                            For Each transition In transitions
                                transition.Remove()
                            Next
                        Next
                        ' Add new transitions to AlternateContentChoice and AlternateContentFallback
                        alternateContent.GetFirstChild(Of AlternateContentChoice)()
                        Dim choiceTransition = New Transition(New RandomBarTransition()) With {
                        .Duration = durationMs,
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                        Dim fallbackTransition = New Transition(New RandomBarTransition()) With {
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                        alternateContent.GetFirstChild(Of AlternateContentChoice)().Append(choiceTransition)
                        alternateContent.GetFirstChild(Of AlternateContentFallback)().Append(fallbackTransition)
                    Next
                    ' </Snippet2>
                    ' <Snippet3>
                    ' Add transition if there is none
                Else
                    ' Check if there is a transition appended to the slide and set it to null
                    If slidePart.Slide.Transition IsNot Nothing Then
                        slidePart.Slide.Transition = Nothing
                    End If

                    ' Create a new AlternateContent element
                    Dim alternateContent As New AlternateContent()
                    alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")

                    ' Create a new AlternateContentChoice element and add the transition
                    Dim alternateContentChoice As New AlternateContentChoice() With {
                    .Requires = "p14"
                }
                    Dim choiceTransition = New Transition(New RandomBarTransition()) With {
                        .Duration = durationMs,
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                    alternateContentChoice.Append(choiceTransition)

                    ' Create a new AlternateContentFallback element and add the transition
                    Dim fallbackTransition = New Transition(New RandomBarTransition()) With {
                        .AdvanceAfterTime = startTransitionAfterMs,
                        .AdvanceOnClick = advanceOnClick,
                        .Speed = TransitionSpeedValues.Slow}
                    Dim alternateContentFallback As New AlternateContentFallback(fallbackTransition)

                    alternateContentFallback.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main")
                    alternateContentFallback.AddNamespaceDeclaration("p16", "http://schemas.microsoft.com/office/powerpoint/2015/main")
                    alternateContentFallback.AddNamespaceDeclaration("adec", "http://schemas.microsoft.com/office/drawing/2017/decorative")
                    alternateContentFallback.AddNamespaceDeclaration("a16", "http://schemas.microsoft.com/office/drawing/2014/main")

                    ' Append the AlternateContentChoice and AlternateContentFallback to the AlternateContent
                    alternateContent.Append(alternateContentChoice)
                    alternateContent.Append(alternateContentFallback)
                    slidePart.Slide.Append(alternateContent)
                End If
                ' </Snippet3>
            Next
        End Using
    End Sub

End Module
' </Snippet0>
