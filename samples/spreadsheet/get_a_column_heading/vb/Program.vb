Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet


Module MyModule

    Sub Main(args As String())
    End Sub

    ' Given a document name, a worksheet name, and a cell name, gets the column of the cell and returns
    ' the content of the first cell in that column.
    Public Function GetColumnHeading(ByVal docName As String, ByVal worksheetName As String, ByVal cellName As String) As String
        ' Open the document as read-only.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, False)

        Using (document)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)
            If (sheets.Count() = 0) Then
                ' The specified worksheet does not exist.
                Return Nothing
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(sheets.First.Id), WorksheetPart)

            ' Get the column name for the specified cell.
            Dim columnName As String = GetColumnName(cellName)

            ' Get the cells in the specified column and order them by row.
            Dim cells As IEnumerable(Of Cell) = worksheetPart.Worksheet.Descendants(Of Cell)().Where(Function(c) _
                String.Compare(GetColumnName(c.CellReference.Value), columnName, True) = 0).OrderBy(Function(r) GetRowIndex(r.CellReference))

            If (cells.Count() = 0) Then
                ' The specified column does not exist.
                Return Nothing
            End If

            ' Get the first cell in the column.
            Dim headCell As Cell = cells.First()

            ' If the content of the first cell is stored as a shared string, get the text of the first cell
            ' from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            If (((headCell.DataType) IsNot Nothing) AndAlso (headCell.DataType.Value = CellValues.SharedString)) Then
                Dim shareStringPart As SharedStringTablePart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
                Dim items() As SharedStringItem = shareStringPart.SharedStringTable.Elements(Of SharedStringItem)().ToArray()
                Return items(Integer.Parse(headCell.CellValue.Text)).InnerText
            Else
                Return headCell.CellValue.Text
            End If

        End Using
    End Function

    ' Given a cell name, parses the specified cell to get the column name.
    Private Function GetColumnName(ByVal cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As Regex = New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)
        Return match.Value
    End Function

    ' Given a cell name, parses the specified cell to get the row index.
    Private Function GetRowIndex(ByVal cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As Regex = New Regex("\d+")
        Dim match As Match = regex.Match(cellName)
        Return UInteger.Parse(match.Value)
    End Function
End Module