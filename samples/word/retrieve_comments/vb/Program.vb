Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub



    Public Sub GetCommentsFromDocument(ByVal fileName As String)
        Using wordDoc As WordprocessingDocument =
            WordprocessingDocument.Open(fileName, False)

            Dim commentsPart As WordprocessingCommentsPart =
                wordDoc.MainDocumentPart.WordprocessingCommentsPart

            If commentsPart IsNot Nothing AndAlso
                commentsPart.Comments IsNot Nothing Then
                For Each comment As Comment In
                    commentsPart.Comments.Elements(Of Comment)()
                    Console.WriteLine(comment.InnerText)
                Next
            End If
        End Using
    End Sub
End Module