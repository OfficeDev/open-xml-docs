Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports System
Imports System.Linq

Module Program
    Sub Main(args As String())
        AddCommentToPresentation(args(0), args(1), args(2), String.Join(" ", args.Skip(3)))
    End Sub

    Sub AddCommentToPresentation(file As String, initials As String, name As String, text As String)
        Using doc As PresentationDocument = PresentationDocument.Open(file, True)
            ' Declare a CommentAuthorsPart object.
            Dim authorsPart As CommentAuthorsPart

            ' If the presentation does not contain a comment authors part, add a new one.
            Dim presentationPart As PresentationPart = If(doc.PresentationPart, doc.AddPresentationPart())

            ' Verify that there is an existing comment authors part and add a new one if not.
            authorsPart = If(presentationPart.CommentAuthorsPart, presentationPart.AddNewPart(Of CommentAuthorsPart)())

            ' Verify that there is a comment author list in the comment authors part and add one if not.
            Dim authorList As CommentAuthorList = If(authorsPart.CommentAuthorList, New CommentAuthorList())
            authorsPart.CommentAuthorList = authorList

            ' Declare a new author ID as either the max existing ID + 1 or 1 if there are no existing IDs.
            Dim authorId As UInteger = If(authorList.Elements(Of CommentAuthor)().Select(Function(a) a.Id?.Value).Max(), 0)
            authorId += 1

            ' If there is an existing author with matching name and initials, use that author otherwise create a new CommentAuthor.
            Dim foo = authorList.Elements(Of CommentAuthor)().Where(Function(a) a.Name = name AndAlso a.Initials = initials).FirstOrDefault()
            Dim author As CommentAuthor = If(foo, authorList.AppendChild(New CommentAuthor() With {
                .Id = authorId,
                .Name = name,
                .Initials = initials,
                .ColorIndex = 0
            }))
            ' get the author id
            authorId = If(author.Id, authorId)

            ' Get the first slide, using the GetFirstSlide method.
            Dim slidePart1 As SlidePart = GetFirstSlide(doc)
            ' <ExtSnippet1>
            ' Declare a comments part.
            Dim commentsPart As SlideCommentsPart

            ' Verify that there is a comments part in the first slide part.
            If slidePart1.GetPartsOfType(Of SlideCommentsPart)().Count() = 0 Then
                ' If not, add a new comments part.
                commentsPart = slidePart1.AddNewPart(Of SlideCommentsPart)()
            Else
                ' Else, use the first comments part in the slide part.
                commentsPart = slidePart1.GetPartsOfType(Of SlideCommentsPart)().First()
            End If

            ' If the comment list does not exist.
            If commentsPart.CommentList Is Nothing Then
                ' Add a new comments list.
                commentsPart.CommentList = New CommentList()
            End If

            ' Get the new comment ID.
            Dim commentIdx As UInteger = If(author.LastIndex, 1)
            author.LastIndex = commentIdx

            ' Add a new comment.
            Dim comment As Comment = commentsPart.CommentList.AppendChild(Of Comment)(
                New Comment() With {
                    .AuthorId = authorId,
                    .Index = commentIdx,
                    .DateTime = DateTime.Now
                })

            ' Add the position child node to the comment element.
            comment.Append(
                New Position() With {.X = 100, .Y = 200},
                New Text() With {.Text = text})
            ' </ExtSnippet1>
            ' Save the comment authors part.
            authorList.Save()

            ' Save the comments part.
            commentsPart.CommentList.Save()
        End Using
    End Sub

    ' Get the slide part of the first slide in the presentation document.
    Function GetFirstSlide(presentationDocument As PresentationDocument) As SlidePart
        ' Get relationship ID of the first slide
        Dim part As PresentationPart = presentationDocument?.PresentationPart
        Dim slideId As SlideId = part?.Presentation?.SlideIdList?.GetFirstChild(Of SlideId)()
        Dim relId As String = slideId?.RelationshipId
        If relId Is Nothing Then
            Throw New ArgumentNullException("The first slide does not contain a relationship ID.")
        End If
        ' Get the slide part by the relationship ID.
        Dim slidePart As SlidePart = TryCast(part?.GetPartById(relId), SlidePart)

        If slidePart Is Nothing Then
            Throw New ArgumentNullException("The slide part is null.")
        End If

        Return slidePart
    End Function
End Module
