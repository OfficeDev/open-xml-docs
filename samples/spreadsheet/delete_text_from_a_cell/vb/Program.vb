Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet


Module MyModule

    Sub Main(args As String())
    End Sub

    ' Given a document, a worksheet name, a column name, and a one-based row index,
    ' deletes the text from the cell at the specified column and row on the specified sheet.
    Public Sub DeleteTextFromCell(ByVal docName As String, ByVal sheetName As String, ByVal colName As String, ByVal rowIndex As UInteger)
        ' Open the document for editing.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (document)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart.Workbook.GetFirstChild(Of Sheets)().Elements(Of Sheet)().Where(Function(s) s.Name = sheetName.ToString())
            If (sheets.Count = 0) Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim relationshipId As String = sheets.First.Id.Value
            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(relationshipId), WorksheetPart)

            ' Get the cell at the specified column and row.
            Dim cell As Cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex)
            If (cell Is Nothing) Then
                ' The specified cell does not exist.
                Return
            End If

            cell.Remove()
            worksheetPart.Worksheet.Save()

        End Using
    End Sub

    ' Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
    Private Function GetSpreadsheetCell(ByVal worksheet As Worksheet, ByVal columnName As String, ByVal rowIndex As UInteger) As Cell
        Dim rows As IEnumerable(Of Row) = worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)().Where(Function(r) r.RowIndex = rowIndex.ToString())
        If (rows.Count = 0) Then
            ' A cell does not exist at the specified row.
            Return Nothing
        End If

        Dim cells As IEnumerable(Of Cell) = rows.First().Elements(Of Cell)().Where(Function(c) String.Compare(c.CellReference.Value, columnName + rowIndex.ToString(), True) = 0)
        If (cells.Count = 0) Then
            ' A cell does not exist at the specified column, in the specified row.
            Return Nothing
        End If

        Return cells.First
    End Function

    ' Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    ' reference the specified SharedStringItem and removes the item.
    Private Sub RemoveSharedStringItem(ByVal shareStringId As Integer, ByVal document As SpreadsheetDocument)
        Dim remove As Boolean = True

        For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
            Dim worksheet As Worksheet = part.Worksheet
            For Each cell In worksheet.GetFirstChild(Of SheetData)().Descendants(Of Cell)()
                ' Verify if other cells in the document reference the item.
                If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString AndAlso cell.CellValue.Text = shareStringId.ToString() Then
                    ' Other cells in the document still reference the item. Do not remove the item.
                    remove = False
                    Exit For
                End If
            Next

            If Not remove Then
                Exit For
            End If
        Next

        ' Other cells in the document do not reference the item. Remove the item.
        If remove Then
            Dim shareStringTablePart As SharedStringTablePart = document.WorkbookPart.SharedStringTablePart
            If shareStringTablePart Is Nothing Then
                Exit Sub
            End If

            Dim item As SharedStringItem = shareStringTablePart.SharedStringTable.Elements(Of SharedStringItem)().ElementAt(shareStringId)
            If item IsNot Nothing Then
                item.Remove()

                ' Refresh all the shared string references.
                For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
                    Dim worksheet As Worksheet = part.Worksheet
                    For Each cell In worksheet.GetFirstChild(Of SheetData)().Descendants(Of Cell)()
                        If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
                            Dim itemIndex As Integer = Integer.Parse(cell.CellValue.Text)
                            If itemIndex > shareStringId Then
                                cell.CellValue.Text = (itemIndex - 1).ToString()
                            End If
                        End If
                    Next
                    worksheet.Save()
                Next

                document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save()
            End If
        End If
    End Sub
End Module