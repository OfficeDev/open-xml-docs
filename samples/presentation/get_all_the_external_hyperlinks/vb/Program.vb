Imports DocumentFormat.OpenXml.Packaging
Imports System
Imports System.Collections.Generic
Imports Drawing = DocumentFormat.OpenXml.Drawing

Module Program
    Sub Main(args As String())
        ' <Snippet3>
        If args.Length = 1 Then
            Dim fileName As String = args(0)
            For Each link As String In GetAllExternalHyperlinksInPresentation(fileName)
                Console.WriteLine(link)
            Next
        End If
        ' </Snippet3>
    End Sub

    ' <Snippet>
    ' Returns all the external hyperlinks in the slides of a presentation.
    Function GetAllExternalHyperlinksInPresentation(fileName As String) As IEnumerable(Of String)
        ' Declare a list of strings.
        Dim ret As New List(Of String)()

        ' <Snippet1>
        ' Open the presentation file as read-only.
        Using document As PresentationDocument = PresentationDocument.Open(fileName, False)
            ' </Snippet1>
            ' If there is no PresentationPart then there are no hyperlinks
            If document.PresentationPart Is Nothing Then
                Return ret
            End If

            ' <Snippet2>
            ' Iterate through all the slide parts in the presentation part.
            For Each slidePart As SlidePart In document.PresentationPart.SlideParts
                Dim links As IEnumerable(Of Drawing.HyperlinkType) = slidePart.Slide.Descendants(Of Drawing.HyperlinkType)()

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
            ' </Snippet2>
        End Using

        ' Return the list of strings.
        Return ret
    End Function
    ' </Snippet>
End Module
