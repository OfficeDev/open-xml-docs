' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule

    Sub Main(args As String())
        ' <Snippet4>
        Dim file As String = args(0)
        Dim txt As String = args(1)

        OpenAndAddTextToWordDocument(file, txt)
        ' </Snippet4>
    End Sub

    Public Sub OpenAndAddTextToWordDocument(ByVal filepath As String, ByVal txt As String)

        ' <Snippet1>
        ' Open a WordprocessingDocument for editing using the filepath.
        Dim wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)

        If wordprocessingDocument Is Nothing Then
            Throw New ArgumentNullException(NameOf(wordprocessingDocument))
        End If
        ' </Snippet1>

        ' <Snippet2>
        ' Assign a reference to the existing document body. 
        Dim mainDocumentPart As MainDocumentPart = If(wordprocessingDocument.MainDocumentPart, wordprocessingDocument.AddMainDocumentPart())

        If wordprocessingDocument.MainDocumentPart.Document Is Nothing Then
            wordprocessingDocument.MainDocumentPart.Document = New Document()
        End If

        If wordprocessingDocument.MainDocumentPart.Document.Body Is Nothing Then
            wordprocessingDocument.MainDocumentPart.Document.Body = New Body()
        End If

        Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body
        ' </Snippet2>

        ' <Snippet3>
        ' Add new text.
        Dim para As Paragraph = body.AppendChild(New Paragraph)
        Dim run As Run = para.AppendChild(New Run)
        run.AppendChild(New Text(txt))
        ' </Snippet3>

        ' There is not using, so Dispose the handle explicitly.
        wordprocessingDocument.Dispose()
    End Sub
End Module
' </Snippet0>
