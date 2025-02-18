Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports System
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main(args As String())
        AcceptAllRevisions(args(0), args(1))
    End Sub

    Sub AcceptAllRevisions(fileName As String, authorName As String)
        Using wdDoc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            If wdDoc.MainDocumentPart Is Nothing OrElse wdDoc.MainDocumentPart.Document.Body Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
            End If

            Dim body As Body = wdDoc.MainDocumentPart.Document.Body

            ' Handle the formatting changes.
            RemoveElements(body.Descendants(Of ParagraphPropertiesChange)().Where(Function(c) c.Author?.Value = authorName))

            ' Handle the deletions.
            RemoveElements(body.Descendants(Of Deleted)().Where(Function(c) c.Author?.Value = authorName))
            RemoveElements(body.Descendants(Of DeletedRun)().Where(Function(c) c.Author?.Value = authorName))
            RemoveElements(body.Descendants(Of DeletedMathControl)().Where(Function(c) c.Author?.Value = authorName))

            ' Handle the insertions.
            HandleInsertions(body, authorName)

            ' Handle move from elements.
            RemoveElements(body.Descendants(Of Paragraph)().Where(Function(p) p.Descendants(Of MoveFrom)().Any(Function(m) m.Author?.Value = authorName)))
            RemoveElements(body.Descendants(Of MoveFromRangeEnd)())

            ' Handle move to elements.
            HandleMoveToElements(body, authorName)
        End Using
    End Sub

    ' Method to remove elements from the document body
    Sub RemoveElements(elements As IEnumerable(Of OpenXmlElement))
        For Each element In elements.ToList()
            element.Remove()
        Next
    End Sub

    ' Method to handle insertions in the document body
    Sub HandleInsertions(body As Body, authorName As String)
        ' Collect all insertion elements by the specified author
        Dim insertions As List(Of OpenXmlElement) = body.Descendants(Of Inserted)().Cast(Of OpenXmlElement)().ToList()
        insertions.AddRange(body.Descendants(Of InsertedRun)().Where(Function(c) c.Author?.Value = authorName))
        insertions.AddRange(body.Descendants(Of InsertedMathControl)().Where(Function(c) c.Author?.Value = authorName))

        For Each insertion In insertions
            ' Promote new content to the same level as the node and then delete the node
            For Each run In insertion.Elements(Of Run)()
                If run Is insertion.FirstChild Then
                    insertion.InsertAfterSelf(New Run(run.OuterXml))
                Else
                    Dim nextSibling As OpenXmlElement = insertion.NextSibling()
                    nextSibling.InsertAfterSelf(New Run(run.OuterXml))
                End If
            Next

            ' Remove specific attributes and the insertion element itself
            insertion.RemoveAttribute("rsidR", "https://schemas.openxmlformats.org/wordprocessingml/2006/main")
            insertion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main")
            insertion.Remove()
        Next
    End Sub

    ' Method to handle move-to elements in the document body
    Sub HandleMoveToElements(body As Body, authorName As String)
        ' Collect all move-to elements by the specified author
        Dim moveToElements As List(Of OpenXmlElement) = body.Descendants(Of MoveToRun)().Cast(Of OpenXmlElement)().ToList()
        moveToElements.AddRange(body.Descendants(Of Paragraph)().Where(Function(p) p.Descendants(Of MoveFrom)().Any(Function(m) m.Author?.Value = authorName)))
        moveToElements.AddRange(body.Descendants(Of MoveToRangeEnd)())

        For Each toElement In moveToElements
            ' Promote new content to the same level as the node and then delete the node
            For Each run In toElement.Elements(Of Run)()
                toElement.InsertBeforeSelf(New Run(run.OuterXml))
            Next
            ' Remove the move-to element itself
            toElement.Remove()
        Next
    End Sub
End Module