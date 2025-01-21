Imports DocumentFormat.OpenXml.Office2021.PowerPoint.Comment
Imports DocumentFormat.OpenXml.Packaging
Imports Comment = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment

Module Program
    Sub Main(args As String())
        DeleteCommentsByAuthorInPresentation(args(0), args(1))
    End Sub

    ' <Snippet0>
    ' Remove all the comments in the slides by a certain author.
    Sub DeleteCommentsByAuthorInPresentation(fileName As String, author As String)
        ' <Snippet1>
        Using doc As PresentationDocument = PresentationDocument.Open(fileName, True)
            ' </Snippet1>
            ' <Snippet2>
            ' Get the modern comments.
            Dim commentAuthors As IEnumerable(Of Author) = doc.PresentationPart?.authorsPart?.AuthorList.Elements(Of Author)().Where(Function(x) x.Name IsNot Nothing AndAlso x.Name.HasValue AndAlso x.Name.Value.Equals(author))
            ' </Snippet2>

            If commentAuthors Is Nothing Then
                Return
            End If

            ' <Snippet3>
            ' Iterate through all the matching authors.
            For Each commentAuthor As Author In commentAuthors
                Dim authorId As String = commentAuthor.Id
                Dim slideParts As IEnumerable(Of SlidePart) = doc.PresentationPart?.SlideParts

                ' If there's no author ID or slide parts, return.
                If authorId Is Nothing OrElse slideParts Is Nothing Then
                    Return
                End If

                ' Iterate through all the slides and get the slide parts.
                For Each slide As SlidePart In slideParts
                    Dim slideCommentsParts As IEnumerable(Of PowerPointCommentPart) = slide.commentParts

                    ' Get the list of comments.
                    If slideCommentsParts IsNot Nothing Then
                        Dim commentsTup = slideCommentsParts.SelectMany(Function(scp) scp.CommentList.Elements(Of Comment)().Where(Function(comment) comment.AuthorId IsNot Nothing AndAlso comment.AuthorId = authorId).Select(Function(c) New Tuple(Of PowerPointCommentPart, Comment)(scp, c)))

                        For Each comment As Tuple(Of PowerPointCommentPart, Comment) In commentsTup
                            ' Delete all the comments by the specified author.
                            comment.Item1.CommentList.RemoveChild(comment.Item2)

                            ' If the commentPart has no existing comment.
                            If comment.Item1.CommentList.ChildElements.Count = 0 Then
                                ' Delete this part.
                                slide.DeletePart(comment.Item1)
                            End If
                        Next
                    End If
                Next

                ' Delete the comment author from the authors part.
                doc.PresentationPart?.authorsPart?.AuthorList.RemoveChild(commentAuthor)
            Next
            ' </Snippet3>
        End Using
    End Sub
    ' </Snippet0>
End Module
