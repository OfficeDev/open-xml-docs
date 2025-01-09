Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        Console.WriteLine("Column heading: {0}", GetColumnHeading(args(0), args(1), args(2)))
    End Sub

    ' Given a document name, a worksheet name, and a cell name, gets the column of the cell and returns
    ' the content of the first cell in that column.
    ' <Snippet0>
    Function GetColumnHeading(docName As String, worksheetName As String, cellName As String) As String
        ' Open the document as read-only.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, False)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart?.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)

            If sheets Is Nothing OrElse sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return Nothing
            End If

            Dim id As String = sheets.First().Id

            If id Is Nothing Then
                ' The worksheet does not have an ID.
                Return Nothing
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(id), WorksheetPart)

            ' <Snippet3>
            ' Get the column name for the specified cell.
            Dim columnName As String = GetColumnName(cellName)

            ' Get the cells in the specified column and order them by row.
            Dim cells As IEnumerable(Of Cell) = worksheetPart.Worksheet.Descendants(Of Cell)().Where(Function(c) String.Compare(GetColumnName(c.CellReference?.Value), columnName, True) = 0) _
                .OrderBy(Function(r) If(GetRowIndex(r.CellReference), 0))
            ' </Snippet3>

            ' <Snippet4>
            If cells.Count() = 0 Then
                ' The specified column does not exist.
                Return Nothing
            End If

            ' Get the first cell in the column.
            Dim headCell As Cell = cells.First()
            ' </Snippet4>

            ' <Snippet5>
            ' If the content of the first cell is stored as a shared string, get the text of the first cell
            ' from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            Dim idx As Integer

            If headCell.DataType IsNot Nothing AndAlso headCell.DataType.Value = CellValues.SharedString AndAlso Integer.TryParse(headCell.CellValue?.Text, idx) Then
                Dim shareStringPart As SharedStringTablePart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
                Dim items As SharedStringItem() = shareStringPart.SharedStringTable.Elements(Of SharedStringItem)().ToArray()

                Return items(idx).InnerText
            Else
                Return headCell.CellValue?.Text
            End If
            ' </Snippet5>
        End Using
    End Function

    ' Given a cell name, parses the specified cell to get the column name.
    Function GetColumnName(cellName As String) As String
        If cellName Is Nothing Then
            Return String.Empty
        End If

        ' <Snippet1>
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)

        Return match.Value
        ' </Snippet1>
    End Function

    ' Given a cell name, parses the specified cell to get the row index.
    Function GetRowIndex(cellName As String) As UInteger?
        If cellName Is Nothing Then
            Return Nothing
        End If

        ' <Snippet2>
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As New Regex("\d+")
        Dim match As Match = regex.Match(cellName)

        Return UInteger.Parse(match.Value)
        ' </Snippet2>
    End Function
    ' </Snippet0>
End Module
