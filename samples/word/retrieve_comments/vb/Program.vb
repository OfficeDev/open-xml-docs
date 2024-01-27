' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        GetCommentsFromDocument(args(0))
    End Sub

    Public Sub GetCommentsFromDocument(ByVal fileName As String)

        ' <Snippet1>
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(fileName, False)

            If wordDoc.MainDocumentPart Is Nothing Or wordDoc.MainDocumentPart.WordprocessingCommentsPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or WordprocessingCommentsPart is null.")
            End If
            ' </Snippet1>

            ' <Snippet2>
            Dim commentsPart As WordprocessingCommentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart

            If commentsPart IsNot Nothing AndAlso commentsPart.Comments IsNot Nothing Then
                For Each comment As Comment In
                    commentsPart.Comments.Elements(Of Comment)()
                    Console.WriteLine(comment.InnerText)
                Next
            End If
            ' </Snippet2>

        End Using
    End Sub
End Module
' </Snippet0>
