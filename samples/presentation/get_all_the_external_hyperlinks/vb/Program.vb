Imports DocumentFormat.OpenXml.Packaging
Imports Drawing = DocumentFormat.OpenXml.Drawing


Module MyModule

    Sub Main(args As String())
    End Sub

    ' Returns all the external hyperlinks in the slides of a presentation.
    Public Function GetAllExternalHyperlinksInPresentation(ByVal fileName As String) As IEnumerable

        ' Declare a list of strings.
        Dim ret As List(Of String) = New List(Of String)

        ' Open the presentation file as read-only.
        Dim document As PresentationDocument = PresentationDocument.Open(fileName, False)

        Using (document)

            ' Iterate through all the slide parts in the presentation part.
            For Each slidePart As SlidePart In document.PresentationPart.SlideParts
                Dim links As IEnumerable = slidePart.Slide.Descendants(Of Drawing.HyperlinkType)()

                ' Iterate through all the links in the slide part.
                For Each link As Drawing.HyperlinkType In links

                    ' Iterate through all the external relationships in the slide part.
                    For Each relation As HyperlinkRelationship In slidePart.HyperlinkRelationships
                        ' If the relationship ID matches the link ID…
                        If relation.Id.Equals(link.Id) Then

                            ' Add the URI of the external relationship to the list of strings.
                            ret.Add(relation.Uri.AbsoluteUri)
                        End If
                    Next
                Next
            Next


            ' Return the list of strings.
            Return ret

        End Using
    End Function
End Module