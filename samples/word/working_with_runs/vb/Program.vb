' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module MyModule

    Sub Main(args As String())
        WriteToWordDoc(args(0), args(1))
    End Sub

    Public Sub WriteToWordDoc(ByVal filepath As String, ByVal txt As String)
        ' Open a WordprocessingDocument for editing using the filepath.
        Using wordprocessingDocument As WordprocessingDocument =
            WordprocessingDocument.Open(filepath, True)
            ' Assign a reference to the existing document body.
            Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body

            ' Add new text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())

            ' Apply bold formatting to the run.
            Dim runProperties As RunProperties = run.AppendChild(New RunProperties(New Bold()))
            run.AppendChild(New Text(txt))
        End Using
    End Sub
End Module
' </Snippet0>