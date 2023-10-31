Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
    End Sub



    ' Given a document name, a worksheet name, and the names of two adjacent cells, merges the two cells.
    ' When two cells are merged, only the content from one cell is preserved:
    ' the upper-left cell for left-to-right languages or the upper-right cell for right-to-left languages.
    Private Sub MergeTwoCells(ByVal docName As String, ByVal sheetName As String, ByVal cell1Name As String, ByVal cell2Name As String)
        ' Open the document for editing.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (document)
            Dim worksheet As Worksheet = GetWorksheet(document, sheetName)
            If ((worksheet Is Nothing) OrElse (String.IsNullOrEmpty(cell1Name) OrElse String.IsNullOrEmpty(cell2Name))) Then
                Return
            End If

            ' Verify if the specified cells exist, and if they do not exist, create them.
            CreateSpreadsheetCellIfNotExist(worksheet, cell1Name)
            CreateSpreadsheetCellIfNotExist(worksheet, cell2Name)

            Dim mergeCells As MergeCells
            If (worksheet.Elements(Of MergeCells)().Count() > 0) Then
                mergeCells = worksheet.Elements(Of MergeCells).First()
            Else
                mergeCells = New MergeCells()

                ' Insert a MergeCells object into the specified position.
                If (worksheet.Elements(Of CustomSheetView)().Count() > 0) Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of CustomSheetView)().First())
                ElseIf (worksheet.Elements(Of DataConsolidate)().Count() > 0) Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of DataConsolidate)().First())
                ElseIf (worksheet.Elements(Of SortState)().Count() > 0) Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of SortState)().First())
                ElseIf (worksheet.Elements(Of AutoFilter)().Count() > 0) Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of AutoFilter)().First())
                ElseIf (worksheet.Elements(Of Scenarios)().Count() > 0) Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of Scenarios)().First())
                ElseIf (worksheet.Elements(Of ProtectedRanges)().Count() > 0) Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of ProtectedRanges)().First())
                ElseIf (worksheet.Elements(Of SheetProtection)().Count() > 0) Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of SheetProtection)().First())
                ElseIf (worksheet.Elements(Of SheetCalculationProperties)().Count() > 0) Then
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of SheetCalculationProperties)().First())
                Else
                    worksheet.InsertAfter(mergeCells, worksheet.Elements(Of SheetData)().First())
                End If
            End If

            ' Create the merged cell and append it to the MergeCells collection.
            Dim mergeCell As MergeCell = New MergeCell()
            mergeCell.Reference = New StringValue((cell1Name + (":" + cell2Name)))
            mergeCells.Append(mergeCell)

            worksheet.Save()
        End Using
    End Sub

    ' Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
    Private Function GetWorksheet(ByVal document As SpreadsheetDocument, ByVal worksheetName As String) As Worksheet
        Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)
        If (sheets.Count = 0) Then
            ' The specified worksheet does not exist.
            Return Nothing
        End If
        Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(sheets.First.Id), WorksheetPart)

        Return worksheetPart.Worksheet
    End Function

    ' Given a Worksheet and a cell name, verifies that the specified cell exists.
    ' If it does not exist, creates a new cell.
    Private Sub CreateSpreadsheetCellIfNotExist(ByVal worksheet As Worksheet, ByVal cellName As String)
        Dim columnName As String = GetColumnName(cellName)
        Dim rowIndex As UInteger = GetRowIndex(cellName)

        Dim rows As IEnumerable(Of Row) = worksheet.Descendants(Of Row)().Where(Function(r) r.RowIndex.Value = rowIndex.ToString())

        ' If the worksheet does not contain the specified row, create the specified row.
        ' Create the specified cell in that row, and insert the row into the worksheet.
        If (rows.Count = 0) Then
            Dim row As Row = New Row()
            row.RowIndex = New UInt32Value(rowIndex)

            Dim cell As Cell = New Cell()
            cell.CellReference = New StringValue(cellName)

            row.Append(cell)
            worksheet.Descendants(Of SheetData)().First().Append(row)
            worksheet.Save()
        Else
            Dim row As Row = rows.First()
            Dim cells As IEnumerable(Of Cell) = row.Elements(Of Cell)().Where(Function(c) c.CellReference.Value = cellName)

            ' If the row does not contain the specified cell, create the specified cell.
            If (cells.Count = 0) Then
                Dim cell As Cell = New Cell
                cell.CellReference = New StringValue(cellName)

                row.Append(cell)
                worksheet.Save()
            End If
        End If
    End Sub

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