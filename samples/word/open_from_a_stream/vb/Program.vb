' <Snippet0>
Imports System.IO
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule

    Sub Main(args As String())
        ' <Snippet4>
        Dim filePath As String = args(0)
        Dim txt As String = args(1)

        Using fileStream As FileStream = New FileStream(filePath, FileMode.Open)
            OpenAndAddToWordprocessingStream(fileStream, txt)
        End Using
        ' </Snippet4>
    End Sub

    Public Sub OpenAndAddToWordprocessingStream(ByVal stream As Stream, ByVal txt As String)

        ' <Snippet1>
        ' Open a WordProcessingDocument based on a stream.
        Dim wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(stream, True)
        ' </Snippet1>

        ' <Snippet2>
        ' Assign a reference to the document body. 
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

        ' Dispose the document handle.
        wordprocessingDocument.Dispose()

        ' Caller must close the stream.
    End Sub
End Module
' </Snippet0>
