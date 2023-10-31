Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub



    ' Change the text in a table in a word processing document.
    Public Sub ChangeTextInCell(ByVal filepath As String, ByVal txt As String)
        ' Use the file name and path passed in as an argument to 
        ' Open an existing document. 
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().First()

            ' Find the second row in the table.
            Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(1)

            ' Find the third cell in the row.
            Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(2)

            ' Find the first paragraph in the table cell.
            Dim p As Paragraph = cell.Elements(Of Paragraph)().First()

            ' Find the first run in the paragraph.
            Dim r As Run = p.Elements(Of Run)().First()

            ' Set the text for the run.
            Dim t As Text = r.Elements(Of Text)().First()
            t.Text = txt
        End Using
    End Sub
End Module