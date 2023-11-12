Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule
    Public Sub CreateWordprocessingDocument(ByVal filepath As String)
        ' Create a document by supplying the filepath.
        Using wordDocument As WordprocessingDocument =
            WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document)

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
End Module
