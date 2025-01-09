Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet


Module MyModule

    Sub Main(args As String())
        CalculateSumOfCellRange(args(0), args(1), args(2), args(3), args(4))
    End Sub

    ' Given a document name, a worksheet name, the name of the first cell in the contiguous range, 
    ' the name of the last cell in the contiguous range, and the name of the results cell, 
    ' calculates the sum of the cells in the contiguous range and inserts the result into the results cell.
    ' Note: All cells in the contiguous range must contain numbers.
    ' <Snippet0>
    ' <Snippet1>
    Private Sub CalculateSumOfCellRange(ByVal docName As String, ByVal worksheetName As String, ByVal firstCellName As String,
    ByVal lastCellName As String, ByVal resultCell As String)
        ' Open the document for editing.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (document)
            Dim sheets As IEnumerable(Of Sheet) =
                document.WorkbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)
            If (sheets.Count() = 0) Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(sheets.First().Id), WorksheetPart)
            Dim worksheet As Worksheet = worksheetPart.Worksheet

            ' Get the row number and column name for the first and last cells in the range.
            Dim firstRowNum As UInteger = GetRowIndex(firstCellName)
            Dim lastRowNum As UInteger = GetRowIndex(lastCellName)
            Dim firstColumn As String = GetColumnName(firstCellName)
            Dim lastColumn As String = GetColumnName(lastCellName)

            Dim sum As Double = 0

            ' Iterate through the cells within the range and add their values to the sum.
            For Each row As Row In worksheet.Descendants(Of Row)().Where(Function(r) r.RowIndex.Value >= firstRowNum _
                                                                             AndAlso r.RowIndex.Value <= lastRowNum)
                For Each cell As Cell In row
                    Dim columnName As String = GetColumnName(cell.CellReference.Value)
                    If ((CompareColumn(columnName, firstColumn) >= 0) AndAlso (CompareColumn(columnName, lastColumn) <= 0)) Then
                        sum = (sum + Double.Parse(cell.CellValue.Text))
                    End If
                Next
            Next

            ' Get the SharedStringTablePart and add the result to it.
            ' If the SharedStringPart does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart
            If (document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0) Then
                shareStringPart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
            Else
                shareStringPart = document.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the result into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(("Result:" + sum.ToString()), shareStringPart)

            Dim result As Cell = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart)

            ' Set the value of the cell.
            result.CellValue = New CellValue(index.ToString())
            result.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
            worksheetPart.Worksheet.Save()
        End Using
    End Sub
    ' </Snippet1>

    ' <Snippet2>
    ' Given a cell name, parses the specified cell to get the row index.
    Private Function GetRowIndex(ByVal cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As Regex = New Regex("\d+")
        Dim match As Match = regex.Match(cellName)
        Return UInteger.Parse(match.Value)
    End Function
    ' </Snippet2>

    ' <Snippet3>
    ' Given a cell name, parses the specified cell to get the column name.
    Private Function GetColumnName(ByVal cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As Regex = New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)
        Return match.Value
    End Function
    ' </Snippet3>

    ' <Snippet4>
    ' Given two columns, compares the columns.
    Private Function CompareColumn(ByVal column1 As String, ByVal column2 As String) As Integer
        If (column1.Length > column2.Length) Then
            Return 1
        ElseIf (column1.Length < column2.Length) Then
            Return -1
        Else
            Return String.Compare(column1, column2, True)
        End If
    End Function
    ' </Snippet4>

    ' <Snippet5>
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create it.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                ' The text already exists in the part. Return its index.
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function
    ' </Snippet5>

    ' <Snippet6>
    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, return it. 
    Private Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = (columnName + rowIndex.ToString())

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row
        If (sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).Count() <> 0) Then
            row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).First()
        Else
            row = New Row()
            row.RowIndex = rowIndex
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If (row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0) Then
            Return row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If (String.Compare(cell.CellReference.Value, cellReference, True) > 0) Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As Cell = New Cell
            newCell.CellReference = cellReference

            row.InsertBefore(newCell, refCell)
            worksheet.Save()

            Return newCell
        End If
    End Function
    ' </Snippet6>
    ' </Snippet0>

End Module