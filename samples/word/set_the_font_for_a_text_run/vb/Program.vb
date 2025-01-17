Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module MyModule

    ' Set the font for a text run.
    ' <Snippet0>
    Sub SetRunFont(fileName As String)
        ' Open a Wordprocessing document for editing.
        Using package As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' <Snippet1>
            ' Set the font to Arial to the first Run.
            ' Use an object initializer for RunProperties and rPr.
            Dim rPr As New RunProperties(New RunFonts() With {
                .Ascii = "Arial"
            })
            ' </Snippet1>

            ' <Snippet2>
            If package.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is null.")
            End If

            Dim r As Run = package.MainDocumentPart.Document.Descendants(Of Run)().First()
            r.PrependChild(Of RunProperties)(rPr)
            ' </Snippet2>
        End Using
    End Sub
    ' </Snippet0>

    Sub Main(args As String())
        SetRunFont(args(0))
    End Sub

End Module

