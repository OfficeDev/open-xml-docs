' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        Dim filename As String = args(0)

        RemoveHeadersAndFooters(filename)
        ' </Snippet2>
    End Sub


    ' To remove all of the headers and footers in a document.
    ' <Snippet1>
    Public Sub RemoveHeadersAndFooters(ByVal filename As String)
        ' </Snippet1>

        ' <Snippet3>
        ' Given a document name, remove all of the headers and footers
        ' from the document.
        Using doc = WordprocessingDocument.Open(filename, True)
            ' </Snippet3>

            ' <Snippet4>
            If doc.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is Nothing")
            End If

            ' Get a reference to the main document part.
            Dim docPart = doc.MainDocumentPart

            ' Count the header and footer parts and continue if there 
            ' are any.
            If (docPart.HeaderParts.Count > 0) Or (docPart.FooterParts.Count > 0) Then
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
                Dim headers = document.Descendants(Of HeaderReference).ToList()

                For Each header In headers
                    header.Remove()
                Next

                ' First, create a list of all descendants of type
                ' FooterReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim footers = document.Descendants(Of FooterReference).ToList()

                For Each footer In footers
                    footer.Remove()
                Next

                ' Save the changes.
                document.Save()
            End If
            ' </Snippet6>
        End Using
    End Sub
End Module
' </Snippet0>
