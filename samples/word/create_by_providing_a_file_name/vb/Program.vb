Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule

    Sub Main(args As String())
        ' <Snippet2>
        CreateWordprocessingDocument(args(0))
        ' </Snippet2>
    End Sub

    ' <Snippet0>
    Public Sub CreateWordprocessingDocument(ByVal filepath As String)
        ' Create a document by supplying the filepath.
        ' <Snippet1>
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document)
            ' </Snippet1>

            ' Add a main document part. 
            Dim mainPart As MainDocumentPart = wordDocument.AddMainDocumentPart()

            ' Create the document structure and add some text.
            mainPart.Document = New Document()
            Dim body As Body = mainPart.Document.AppendChild(New Body())
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Create text in body - CreateWordprocessingDocument"))
        End Using
    End Sub
    ' </Snippet0>
End Module