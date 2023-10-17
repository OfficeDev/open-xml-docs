Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        Dim fileName = args(0)
        Dim authorName = args(1)

        'Public Sub AcceptRevisions(ByVal fileName As String, ByVal authorName As String)
        ' Given a document name and an author name, accept revisions. 
        Using wdDoc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            Dim body As Body = wdDoc.MainDocumentPart.Document.Body

            ' Handle the formatting changes.
            Dim changes As List(Of OpenXmlElement) =
                body.Descendants(Of ParagraphPropertiesChange)() _
                .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList()

            For Each change In changes
                change.Remove()
            Next

            ' Handle the deletions.
            Dim deletions As List(Of OpenXmlElement) =
                body.Descendants(Of Deleted)() _
                .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList()

            deletions.AddRange(body.Descendants(Of DeletedRun)() _
            .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList())

            deletions.AddRange(body.Descendants(Of DeletedMathControl)() _
            .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList())

            For Each deletion In deletions
                deletion.Remove()
            Next

            ' Handle the insertions.
            Dim insertions As List(Of OpenXmlElement) =
                body.Descendants(Of Inserted)() _
                .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList()

            insertions.AddRange(body.Descendants(Of InsertedRun)() _
            .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList())

            insertions.AddRange(body.Descendants(Of InsertedMathControl)() _
            .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList())

            For Each insertion In insertions
                ' Found new content. Promote them to the same level as node, and then
                ' delete the node.
                For Each run In insertion.Elements(Of Run)()
                    If run Is insertion.FirstChild Then
                        insertion.InsertAfterSelf(New Run(run.OuterXml))
                    Else
                        insertion.NextSibling().InsertAfterSelf(New Run(run.OuterXml))
                    End If
                Next
                insertion.RemoveAttribute("rsidR", "https://schemas.openxmlformats.org/wordprocessingml/2006/main")
                insertion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main")
                insertion.Remove()
            Next
        End Using
    End Sub
End Module
