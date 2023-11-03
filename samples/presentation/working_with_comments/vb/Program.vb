
     ' Declare a comments part.
          Dim commentsPart As SlideCommentsPart

          ' Verify that there is a comments part in the first slide part.
          If slidePart1.GetPartsOfType(Of SlideCommentsPart)().Count() = 0 Then

             ' If not, add a new comments part.
             commentsPart = slidePart1.AddNewPart(Of SlideCommentsPart)()
          Else

             ' Else, use the first comments part in the slide part.
             commentsPart = _
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
          Dim comment As Comment = _
           (commentsPart.CommentList.AppendChild(Of Comment)(New Comment() _
             With {.AuthorId = authorId, .Index = commentIdx, .DateTime = DateTime.Now}))

          ' Add the position child node to the comment element.
          comment.Append(New Position() With _
               {.X = 100, .Y = 200}, New Text() With {.Text = text})


     ' Declare a comments part.
          Dim commentsPart As SlideCommentsPart

          ' Verify that there is a comments part in the first slide part.
          If slidePart1.GetPartsOfType(Of SlideCommentsPart)().Count() = 0 Then

             ' If not, add a new comments part.
             commentsPart = slidePart1.AddNewPart(Of SlideCommentsPart)()
          Else

             ' Else, use the first comments part in the slide part.
             commentsPart = _
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
          Dim comment As Comment = _
           (commentsPart.CommentList.AppendChild(Of Comment)(New Comment() _
             With {.AuthorId = authorId, .Index = commentIdx, .DateTime = DateTime.Now}))

          ' Add the position child node to the comment element.
          comment.Append(New Position() With _
               {.X = 100, .Y = 200}, New Text() With {.Text = text})
