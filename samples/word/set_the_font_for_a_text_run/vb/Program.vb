' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        SetRunFont(args(0))
    End Sub



    ' Set the font for a text run.
    Public Sub SetRunFont(ByVal fileName As String)
        ' Open a Wordprocessing document for editing.
        Dim package As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
        Using (package)

            ' <Snippet1>
            ' Set the font to Arial to the first Run.
            ' Use an object initializer for RunProperties and rPr.
            Dim rPr As RunProperties = New RunProperties(New RunFonts With {.Ascii = "Arial"})
            ' </Snippet1>

            ' <Snippet2>
            Dim r As Run = package.MainDocumentPart.Document.Descendants(Of Run).First
            r.PrependChild(Of RunProperties)(rPr)
            ' </Snippet2>

            ' Save changes to the main document part.
            package.MainDocumentPart.Document.Save()
        End Using
    End Sub
End Module
' </Snippet0>
