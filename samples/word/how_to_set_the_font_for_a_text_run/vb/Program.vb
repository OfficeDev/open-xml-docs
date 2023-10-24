Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub



    ' Set the font for a text run.
    Public Sub SetRunFont(ByVal fileName As String)
        ' Open a Wordprocessing document for editing.
        Dim package As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
        Using (package)
            ' Set the font to Arial to the first Run.
            Dim rPr As RunProperties = New RunProperties(New RunFonts With {.Ascii = "Arial"})
            Dim r As Run = package.MainDocumentPart.Document.Descendants(Of Run).First

            r.PrependChild(Of RunProperties)(rPr)

            ' Save changes to the main document part.
            package.MainDocumentPart.Document.Save()
        End Using
    End Sub
End Module