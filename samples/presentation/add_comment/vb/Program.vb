Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation

Module Program
    Sub Main(args As String())
        AddCommentToPresentation(args(0), args(1), args(2), String.Join(" ", args.Skip(3)))
    End Sub

    Public Sub AddCommentToPresentation(ByVal file As String,
                ByVal initials As String,
                ByVal name As String,
                ByVal text As String)

        Dim doc As PresentationDocument =
          PresentationDocument.Open(file, True)

        Using (doc)

            ' Declare a CommentAuthorsPart object.
            Dim authorsPart As CommentAuthorsPart

            ' Verify that there is an existing comment authors part.
            If (doc.PresentationPart.CommentAuthorsPart Is Nothing) Then

                ' If not, add a new one.
                authorsPart = doc.PresentationPart.AddNewPart(Of CommentAuthorsPart)()
            Else
                authorsPart = doc.PresentationPart.CommentAuthorsPart
            End If

            ' Verify that there is a comment author list in the comment authors part.
            If (authorsPart.CommentAuthorList Is Nothing) Then

                ' If not, add a new one.
                authorsPart.CommentAuthorList = New CommentAuthorList()
            End If

            ' Declare a new author ID.
            Dim authorId As UInteger = 0
            Dim author As CommentAuthor = Nothing

            ' If there are existing child elements in the comment authors list.
            If authorsPart.CommentAuthorList.HasChildren = True Then

                ' Verify that the author passed in is on the list.
                Dim authors = authorsPart.CommentAuthorList.Elements(Of CommentAuthor)().Where _
              (Function(a) a.Name = name AndAlso a.Initials = initials)

                ' If so...
                If (authors.Any()) Then

                    ' Assign the new comment author the existing ID.
                    author = authors.First()
                    authorId = author.Id
                End If

                ' If not...
                If (author Is Nothing) Then

                    ' Assign the author passed in a new ID.
                    authorId =
                authorsPart.CommentAuthorList.Elements(Of CommentAuthor)().Select(Function(a) a.Id.Value).Max()
                End If

            End If

            ' If there are no existing child elements in the comment authors list.
            If (author Is Nothing) Then

                authorId = authorId + 1

                ' Add a new child element (comment author) to the comment author list.
                author = (authorsPart.CommentAuthorList.AppendChild(Of CommentAuthor) _
              (New CommentAuthor() With {.Id = authorId,
               .Name = name,
               .Initials = initials,
               .ColorIndex = 0}))
            End If

            ' Get the first slide, using the GetFirstSlide() method.
            Dim slidePart1 As SlidePart
            slidePart1 = GetFirstSlide(doc)

            ' Declare a comments part.
            Dim commentsPart As SlideCommentsPart

            ' Verify that there is a comments part in the first slide part.
            If slidePart1.GetPartsOfType(Of SlideCommentsPart)().Count() = 0 Then

                ' If not, add a new comments part.
                commentsPart = slidePart1.AddNewPart(Of SlideCommentsPart)()
            Else

                ' Else, use the first comments part in the slide part.
                commentsPart =
              slidePart1.GetPartsOfType(Of SlideCommentsPart)().First()
            End If

            ' If the comment list does not exist.
            If (commentsPart.CommentList Is Nothing) Then

                ' Add a new comments list.
                commentsPart.CommentList = New CommentList()
            End If

            ' Get the new comment ID.
            Dim commentIdx As UInteger
            If author.LastIndex Is Nothing Then
                commentIdx = 1
            Else
                commentIdx = CType(author.LastIndex, UInteger) + 1
            End If

            author.LastIndex = commentIdx

            ' Add a new comment.
            Dim comment As Comment =
(commentsPart.CommentList.AppendChild(Of Comment)(New Comment() _
             With {.AuthorId = authorId, .Index = commentIdx, .DateTime = DateTime.Now}))

            ' Add the position child node to the comment element.
            comment.Append(New Position() With
               {.X = 100, .Y = 200}, New Text() With {.Text = text})


            ' Save comment authors part.
            authorsPart.CommentAuthorList.Save()

            ' Save comments part.
            commentsPart.CommentList.Save()

        End Using

    End Sub

    ' Get the slide part of the first slide in the presentation document.
    Public Function GetFirstSlide(ByVal presentationDocument As PresentationDocument) As SlidePart
        ' Get relationship ID of the first slide
        Dim part As PresentationPart = presentationDocument.PresentationPart
        Dim slideId As SlideId = part.Presentation.SlideIdList.GetFirstChild(Of SlideId)()
        Dim relId As String = slideId.RelationshipId

        ' Get the slide part by the relationship ID.
        Dim slidePart As SlidePart = DirectCast(part.GetPartById(relId), SlidePart)

        Return slidePart
    End Function
End Module
