Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module MyModule

    ' <Snippet0>
    ' Remove all of the headers and footers from a document.
    ' <Snippet1>
    Sub RemoveHeadersAndFooters(filename As String)
        ' </Snippet1>

        ' <Snippet3>
        ' Given a document name, remove all of the headers and footers
        ' from the document.
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filename, True)

            If doc.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is null.")
            End If

            ' Get a reference to the main document part.
            Dim docPart = doc.MainDocumentPart
            ' </Snippet3>

            ' <Snippet4>
            ' Count the header and footer parts and continue if there 
            ' are any.
            If docPart.HeaderParts.Count() > 0 OrElse docPart.FooterParts.Count() > 0 Then
                ' </Snippet4>

                ' <Snippet5>
                ' Remove the header and footer parts.
                docPart.DeleteParts(docPart.HeaderParts)
                docPart.DeleteParts(docPart.FooterParts)
                ' </Snippet5>

                ' <Snippet6>
                ' Get a reference to the root element of the main
                ' document part.
                Dim document As Document = docPart.Document

                ' Remove all references to the headers and footers.

                ' First, create a list of all descendants of type
                ' HeaderReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim headers = document.Descendants(Of HeaderReference)().ToList()

                For Each header In headers
                    header.Remove()
                Next

                ' First, create a list of all descendants of type
                ' FooterReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim footers = document.Descendants(Of FooterReference)().ToList()

                For Each footer In footers
                    footer.Remove()
                Next
                ' </Snippet6>
            End If
        End Using
    End Sub
    ' </Snippet0>

    Sub Main(args As String())
        ' <Snippet2>
        Dim filename As String = args(0)

        RemoveHeadersAndFooters(filename)
        ' </Snippet2>
    End Sub

End Module

