Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub



    ' Insert a comment on the first paragraph.
    Public Sub AddCommentOnFirstParagraph(ByVal fileName As String, ByVal author As String, ByVal initials As String, ByVal comment As String)
        ' Use the file name and path passed in as an 
        ' argument to open an existing Wordprocessing document. 
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' Locate the first paragraph in the document.
            Dim firstParagraph As Paragraph = document.MainDocumentPart.Document.Descendants(Of Paragraph)().First()
            Dim comments As Comments = Nothing
            Dim id As String = "0"

            ' Verify that the document contains a 
            ' WordProcessingCommentsPart part; if not, add a new one.
            If document.MainDocumentPart.GetPartsOfType(Of WordprocessingCommentsPart).Count() > 0 Then
                comments = document.MainDocumentPart.WordprocessingCommentsPart.Comments
                If comments.HasChildren Then
                    ' Obtain an unused ID.
                    id = comments.Descendants(Of Comment)().[Select](Function(e) e.Id.Value).Max()
                End If
            Else
                ' No WordprocessingCommentsPart part exists, so add one to the package.
                Dim commentPart As WordprocessingCommentsPart = document.MainDocumentPart.AddNewPart(Of WordprocessingCommentsPart)()
                commentPart.Comments = New Comments()
                comments = commentPart.Comments
            End If

            ' Compose a new Comment and add it to the Comments part.
            Dim p As New Paragraph(New Run(New Text(comment)))
            Dim cmt As New Comment() With {.Id = id, .Author = author, .Initials = initials, .Date = DateTime.Now}
            cmt.AppendChild(p)
            comments.AppendChild(cmt)
            comments.Save()

            ' Specify the text range for the Comment. 
            ' Insert the new CommentRangeStart before the first run of paragraph.
            firstParagraph.InsertBefore(New CommentRangeStart() With {.Id = id}, firstParagraph.GetFirstChild(Of Run)())

            ' Insert the new CommentRangeEnd after last run of paragraph.
            Dim cmtEnd = firstParagraph.InsertAfter(New CommentRangeEnd() With {.Id = id}, firstParagraph.Elements(Of Run)().Last())

            ' Compose a run with CommentReference and insert it.
            firstParagraph.InsertAfter(New Run(New CommentReference() With {.Id = id}), cmtEnd)
        End Using
    End Sub
End Module