Imports DocumentFormat.OpenXml.Packaging
Imports System
Imports System.IO

Module Program
    Sub Main(args As String())
        ' <Snippet4>
        Dim document As String = args(0)
        GetCommentsFromDocument(document)
        ' </Snippet4>
    End Sub

    ' To get the contents of a document part.
    ' <Snippet2>
    ' <Snippet0>
    Function GetCommentsFromDocument(document As String) As String
        Dim comments As String = Nothing

        ' <Snippet1>
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, False)
            ' </Snippet1>
            If wordDoc Is Nothing Then
                Throw New ArgumentNullException(NameOf(wordDoc))
            End If

            Dim mainPart As MainDocumentPart = If(wordDoc.MainDocumentPart, wordDoc.AddMainDocumentPart())
            Dim WordprocessingCommentsPart As WordprocessingCommentsPart = If(mainPart.WordprocessingCommentsPart, mainPart.AddNewPart(Of WordprocessingCommentsPart)())
            ' </Snippet2>

            ' <Snippet3>
            Using streamReader As New StreamReader(WordprocessingCommentsPart.GetStream())
                comments = streamReader.ReadToEnd()
            End Using
        End Using

        Return comments
        ' </Snippet3>
    End Function
    ' </Snippet0>
End Module


