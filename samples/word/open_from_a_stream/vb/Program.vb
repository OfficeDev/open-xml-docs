Imports System.IO
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule

    Sub Main(args As String())
    End Sub

    Public Sub OpenAndAddToWordprocessingStream(ByVal stream As Stream, ByVal txt As String)
        ' Open a WordProcessingDocument based on a stream.
        Dim wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(stream, True)

        ' Assign a reference to the existing document body.
        Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body

        ' Add new text.
        Dim para As Paragraph = body.AppendChild(New Paragraph)
        Dim run As Run = para.AppendChild(New Run)
        run.AppendChild(New Text(txt))

        ' Dispose the document handle.
        wordprocessingDocument.Dispose()

        ' Caller must close the stream.
    End Sub
End Module