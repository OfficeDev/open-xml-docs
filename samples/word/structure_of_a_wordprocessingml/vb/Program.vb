Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module MyModule
    Sub CreateWordDoc(filepath As String, msg As String)
        Using doc As WordprocessingDocument = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
            ' Add a main document part. 
            Dim mainPart As MainDocumentPart = doc.AddMainDocumentPart()

            ' Create the document structure and add some text.
            mainPart.Document = New Document()
            Dim body As Body = mainPart.Document.AppendChild(New Body())
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())

            ' String msg contains the text, "Hello, Word!"
            run.AppendChild(New Text(msg))
        End Using
    End Sub
End Module
