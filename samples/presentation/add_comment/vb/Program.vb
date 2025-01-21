Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Office2016.Presentation.Command
Imports DocumentFormat.OpenXml.Office2021.PowerPoint.Comment
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation

Module Program
    Sub Main(args As String())
        AddCommentToPresentation(args(0), args(1), args(2), String.Join(" ", args.Skip(3)))
    End Sub

    ' <Snippet0>
    Sub AddCommentToPresentation(file As String, initials As String, name As String, text As String)
        Using presentationDocument As PresentationDocument = PresentationDocument.Open(file, True)
            Dim presentationPart As PresentationPart = presentationDocument?.PresentationPart

            If (presentationPart Is Nothing) Then
                Throw New MissingFieldException("PresentationPart")
            End If

            ' <Snippet1>
            ' create missing PowerPointAuthorsPart if it is null
            If presentationDocument.PresentationPart.authorsPart Is Nothing Then
                presentationDocument.PresentationPart.AddNewPart(Of PowerPointAuthorsPart)()
            End If
            ' </Snippet1>

            ' <Snippet2>
            ' Add missing AuthorList if it is null
            If presentationDocument.PresentationPart.authorsPart Is Nothing Or presentationDocument.PresentationPart.authorsPart.AuthorList Is Nothing Then
                presentationDocument.PresentationPart.authorsPart.AuthorList = New AuthorList()
            End If

            ' Get the author or create a new one
            Dim author As Author = presentationDocument.PresentationPart.authorsPart.AuthorList _
                .ChildElements.OfType(Of Author)().Where(Function(a) a.Name?.Value = name).FirstOrDefault()

            If author Is Nothing Then
                Dim authorId As String = String.Concat("{", Guid.NewGuid(), "}")
                Dim userId As String = String.Concat(If(name.Split(" "c).FirstOrDefault(), "user"), "@example.com::", Guid.NewGuid())
                author = New Author() With {.Id = authorId, .Name = name, .Initials = initials, .UserId = userId, .ProviderId = String.Empty}

                presentationDocument.PresentationPart.authorsPart.AuthorList.AppendChild(author)
            End If
            ' </Snippet2>

            ' <Snippet3>
            ' Get the Id of the slide to add the comment to
            Dim slideId As SlideId = presentationDocument.PresentationPart.Presentation.SlideIdList?.Elements(Of SlideId)()?.FirstOrDefault()

            ' If slideId is null, there are no slides, so return
            If slideId Is Nothing Then Return
            ' </Snippet3>
            Dim ran As New Random()
            Dim cid As UInt32Value = Convert.ToUInt32(ran.Next(100000000, 999999999))

            ' <Snippet4>
            ' Get the relationship id of the slide if it exists
            Dim relId As String = slideId.RelationshipId

            ' Use the relId to get the slide if it exists, otherwise take the first slide in the sequence
            Dim slidePart As SlidePart = If(relId IsNot Nothing, CType(presentationPart.GetPartById(relId), SlidePart), presentationDocument.PresentationPart.SlideParts.First())

            ' If the slide part has comments parts take the first PowerPointCommentsPart
            ' otherwise add a new one
            Dim powerPointCommentPart As PowerPointCommentPart = slidePart.commentParts.FirstOrDefault()

            If (powerPointCommentPart Is Nothing) Then
                powerPointCommentPart = slidePart.AddNewPart(Of PowerPointCommentPart)()
            End If
            ' </Snippet4>

            ' <Snippet5>
            ' Create the comment using the new modern comment class DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment
            Dim comment As New DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment(
                New SlideMonikerList(
                    New DocumentMoniker(),
                    New SlideMoniker() With {
                        .CId = cid,
                        .SldId = slideId.Id
                    }),
                New TextBodyType(
                    New BodyProperties(),
                    New ListStyle(),
                    New Paragraph(New Run(New DocumentFormat.OpenXml.Drawing.Text(text))))) With {
                .Id = String.Concat("{", Guid.NewGuid(), "}"),
                .AuthorId = author.Id,
                .Created = DateTime.Now
            }

            ' If the comment list does not exist, add one.
            If (powerPointCommentPart.CommentList Is Nothing) Then
                powerPointCommentPart.CommentList = New DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList()
            End If
            ' Add the comment to the comment list
            powerPointCommentPart.CommentList.AppendChild(comment)
            ' </Snippet5>

            ' <Snippet6>
            ' Get the presentation extension list if it exists
            Dim presentationExtensionList As SlideExtensionList = slidePart.Slide.ChildElements.OfType(Of SlideExtensionList)().FirstOrDefault()
            ' Create a boolean that determines if this is the slide's first comment
            Dim isFirstComment As Boolean = False

            ' If the presentation extension list is null, add one and set this as the first comment for the slide
            If presentationExtensionList Is Nothing Then
                isFirstComment = True
                slidePart.Slide.AppendChild(New SlideExtensionList())
                presentationExtensionList = slidePart.Slide.ChildElements.OfType(Of SlideExtensionList)().First()
            End If

            ' Get the slide extension if it exists
            Dim presentationExtension As SlideExtension = presentationExtensionList.ChildElements.OfType(Of SlideExtension)().FirstOrDefault()

            ' If the slide extension is null, add it and set this as a new comment
            If presentationExtension Is Nothing Then
                isFirstComment = True
                presentationExtensionList.AddChild(New SlideExtension() With {.Uri = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"})
                presentationExtension = presentationExtensionList.ChildElements.OfType(Of SlideExtension)().First()
            End If

            ' If this is the first comment for the slide add the comment relationship
            If isFirstComment Then
                presentationExtension.AddChild(New CommentRelationship() With {.Id = slidePart.GetIdOfPart(powerPointCommentPart)})
            End If
            ' </Snippet6>
        End Using
    End Sub
    ' </Snippet0>
End Module
