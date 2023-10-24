Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub


    ' Delete comments by a specific author. Pass an empty string for the author 
    ' to delete all comments, by all authors.
    Public Sub DeleteComments(ByVal fileName As String,
        Optional ByVal author As String = "")

        ' Get an existing Wordprocessing document.
        Using document As WordprocessingDocument =
            WordprocessingDocument.Open(fileName, True)
            ' Set commentPart to the document 
            ' WordprocessingCommentsPart, if it exists.
            Dim commentPart As WordprocessingCommentsPart =
                document.MainDocumentPart.WordprocessingCommentsPart

            ' If no WordprocessingCommentsPart exists, there can be no
            ' comments. Stop execution and return from the method.
            If (commentPart Is Nothing) Then
                Return
            End If

            ' Create a list of comments by the specified author, or
            ' if the author name is empty, all authors.
            Dim commentsToDelete As List(Of Comment) = _
                commentPart.Comments.Elements(Of Comment)().ToList()
            If Not String.IsNullOrEmpty(author) Then
                commentsToDelete = commentsToDelete.
                Where(Function(c) c.Author = author).ToList()
            End If
            Dim commentIds As IEnumerable(Of String) =
                commentsToDelete.Select(Function(r) r.Id.Value)

            ' Delete each comment in commentToDelete from the Comments 
            ' collection.
            For Each c As Comment In commentsToDelete
                c.Remove()
            Next

            ' Save the comment part change.
            commentPart.Comments.Save()

            Dim doc As Document = document.MainDocumentPart.Document

            ' Delete CommentRangeStart for each 
            ' deleted comment in the main document.
            Dim commentRangeStartToDelete As List(Of CommentRangeStart) = _
                doc.Descendants(Of CommentRangeStart). _
                Where(Function(c) commentIds.Contains(c.Id.Value)).ToList()
            For Each c As CommentRangeStart In commentRangeStartToDelete
                c.Remove()
            Next

            ' Delete CommentRangeEnd for each deleted comment in main document.
            Dim commentRangeEndToDelete As List(Of CommentRangeEnd) = _
                doc.Descendants(Of CommentRangeEnd). _
                Where(Function(c) commentIds.Contains(c.Id.Value)).ToList()
            For Each c As CommentRangeEnd In commentRangeEndToDelete
                c.Remove()
            Next

            ' Delete CommentReference for each deleted comment in the main document.
            Dim commentRangeReferenceToDelete As List(Of CommentReference) = _
                doc.Descendants(Of CommentReference). _
                Where(Function(c) commentIds.Contains(c.Id.Value)).ToList
            For Each c As CommentReference In commentRangeReferenceToDelete
                c.Remove()
            Next

            ' Save changes back to the MainDocumentPart part.
            doc.Save()
        End Using
    End Sub
End Module