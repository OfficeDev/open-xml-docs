' <Snippet>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        If (args.Length >= 2) Then
            Dim fileName As String = args(0)
            Dim author As String = args(1)

            DeleteComments(fileName, author)
        ElseIf args.Length = 1 Then
            Dim fileName As String = args(0)

            DeleteComments(fileName)
        End If
        ' </Snippet2>
    End Sub


    ' <Snippet1>
    ' Delete comments by a specific author. Pass an empty string for the author 
    ' to delete all comments, by all authors.
    Public Sub DeleteComments(ByVal fileName As String, Optional ByVal author As String = "")
        ' </Snippet1>

        ' <Snippet3>
        ' Get an existing Wordprocessing document.
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' Set commentPart to the document 
            ' WordprocessingCommentsPart, if it exists.
            Dim commentPart As WordprocessingCommentsPart = document.MainDocumentPart.WordprocessingCommentsPart

            ' If no WordprocessingCommentsPart exists, there can be no
            ' comments. Stop execution and return from the method.
            If (commentPart Is Nothing) Then
                Return
            End If
            ' </Snippet3>

            ' Create a list of comments by the specified author, or
            ' if the author name is empty, all authors.
            ' <Snippet4>
            Dim commentsToDelete As List(Of Comment) = commentPart.Comments.Elements(Of Comment)().ToList()
            ' </Snippet4>

            ' <Snippet5>
            If Not String.IsNullOrEmpty(author) Then
                commentsToDelete = commentsToDelete.Where(Function(c) c.Author = author).ToList()
            End If
            ' </Snippet5>

            ' <Snippet6>
            Dim commentIds As IEnumerable(Of String) = commentsToDelete.Where(Function(r) r IsNot Nothing And r.Id.HasValue).Select(Function(r) r.Id.Value)
            ' </Snippet6>

            ' <Snippet7>
            ' Delete each comment in commentToDelete from the Comments 
            ' collection.
            For Each c As Comment In commentsToDelete
                If (c IsNot Nothing) Then
                    c.Remove()
                End If
            Next

            ' Save the comment part change.
            commentPart.Comments.Save()
            ' </Snippet7>

            ' <Snippet8>
            Dim doc As Document = document.MainDocumentPart.Document
            ' </Snippet8>

            ' <Snippet9>
            ' Delete CommentRangeStart for each 
            ' deleted comment in the main document.
            Dim commentRangeStartToDelete As List(Of CommentRangeStart) = doc.Descendants(Of CommentRangeStart).
                                                                            Where(Function(c) commentIds.Contains(c.Id.Value)).
                                                                            ToList()

            For Each c As CommentRangeStart In commentRangeStartToDelete
                c.Remove()
            Next

            ' Delete CommentRangeEnd for each deleted comment in main document.
            Dim commentRangeEndToDelete As List(Of CommentRangeEnd) = doc.Descendants(Of CommentRangeEnd).
                                                                        Where(Function(c) commentIds.Contains(c.Id.Value)).
                                                                        ToList()

            For Each c As CommentRangeEnd In commentRangeEndToDelete
                c.Remove()
            Next

            ' Delete CommentReference for each deleted comment in the main document.
            Dim commentRangeReferenceToDelete As List(Of CommentReference) = doc.Descendants(Of CommentReference).
                                                                                Where(Function(c) commentIds.Contains(c.Id.Value)).
                                                                                ToList()

            For Each c As CommentReference In commentRangeReferenceToDelete
                c.Remove()
            Next

            ' Save changes back to the MainDocumentPart part.
            doc.Save()
            ' </Snippet9>
        End Using
    End Sub
End Module
' </Snippet>
