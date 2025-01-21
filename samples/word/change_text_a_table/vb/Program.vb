Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        ' <Snippet4>
        ChangeTextInCell(args(0), args(1))
        ' </Snippet4>
    End Sub


    ' <Snippet0>
    ' Change the text in a table in a word processing document.
    Public Sub ChangeTextInCell(ByVal filepath As String, ByVal txt As String)
        ' Use the file name and path passed in as an argument to 
        ' Open an existing document. 
        ' <Snippet1>
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' </Snippet1>
            ' <Snippet2>
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().First()

            ' Find the second row in the table.
            Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(1)

            ' Find the third cell in the row.
            Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(2)
            ' </Snippet2>
            ' <Snippet3>
            ' Find the first paragraph in the table cell.
            Dim p As Paragraph = cell.Elements(Of Paragraph)().First()

            ' Find the first run in the paragraph.
            Dim r As Run = p.Elements(Of Run)().First()

            ' Set the text for the run.
            Dim t As Text = r.Elements(Of Text)().First()
            t.Text = txt
            ' </Snippet3>
        End Using
    End Sub
    ' </Snippet0>
End Module