Imports System
Imports System.IO
Imports DocumentFormat.OpenXml.Packaging


Module MyModule
' To get the contents of a document part.
    Public Function GetCommentsFromDocument(ByVal document As String) As String
        Dim comments As String = Nothing
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, False)
        Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
        Dim WordprocessingCommentsPart As WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart
        Dim streamReader As StreamReader = New StreamReader(WordprocessingCommentsPart.GetStream)
        comments = streamReader.ReadToEnd
        Return comments
    End Function
End Module
