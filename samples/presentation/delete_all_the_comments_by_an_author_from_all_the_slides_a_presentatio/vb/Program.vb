Imports System
Imports System.Linq
Imports System.Collections.Generic
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Packaging


Module MyModule
' Remove all the comments in the slides by a certain author.
    Public Sub DeleteCommentsByAuthorInPresentation(ByVal fileName As String, ByVal author As String)

        Dim doc As PresentationDocument = PresentationDocument.Open(fileName, True)

        If (String.IsNullOrEmpty(fileName) Or String.IsNullOrEmpty(author)) Then
            Throw New ArgumentNullException("File name or author name is NULL!")
        End If

        Using (doc)

            ' Get the specified comment author.
            Dim commentAuthors = doc.PresentationPart.CommentAuthorsPart. _
                CommentAuthorList.Elements(Of CommentAuthor)().Where(Function(e) _
                   e.Name.Value.Equals(author))

            ' Dim changed As Boolean = False
            For Each commentAuthor In commentAuthors

                Dim authorId = commentAuthor.Id

                ' Iterate through all the slides and get the slide parts.
                For Each slide In doc.PresentationPart.GetPartsOfType(Of SlidePart)()

                    ' Get the slide comments part of each slide.
                    For Each slideCommentsPart In slide.GetPartsOfType(Of SlideCommentsPart)()

                        ' Delete all the comments by the specified author.
                        Dim commentList = slideCommentsPart.CommentList.Elements(Of Comment)(). _
                            Where(Function(e) e.AuthorId.Value.Equals(authorId.Value))

                        Dim comments As List(Of Comment) = commentList.ToList()

                        For Each comm As Comment In comments
                            slideCommentsPart.CommentList.RemoveChild(Of Comment)(comm)
                        Next
                    Next
                Next

                ' Delete the comment author from the comment authors part.
                doc.PresentationPart.CommentAuthorsPart.CommentAuthorList.RemoveChild(Of CommentAuthor)(commentAuthor)

            Next

        End Using
    End Sub
End Module
