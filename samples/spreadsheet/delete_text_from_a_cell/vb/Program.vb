Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        DeleteTextFromCell(args(0), args(1), args(2), UInteger.Parse(args(3)))
    End Sub

    ' <Snippet1>
    ' Given a document, a worksheet name, a column name, and a one-based row index,
    ' deletes the text from the cell at the specified column and row on the specified worksheet.
    ' <Snippet0>
    Sub DeleteTextFromCell(docName As String, sheetName As String, colName As String, rowIndex As UInteger)
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart?.Workbook?.GetFirstChild(Of Sheets)()?.Elements(Of Sheet)()?.Where(Function(s) s.Name IsNot Nothing AndAlso s.Name = sheetName)
            If sheets Is Nothing OrElse sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If
            Dim relationshipId As String = sheets.First()?.Id?.Value

            If relationshipId Is Nothing Then
                ' The worksheet does not have a relationship ID.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(relationshipId), WorksheetPart)

            ' Get the cell at the specified column and row.
            Dim cell As Cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex)
            If cell Is Nothing Then
                ' The specified cell does not exist.
                Return
            End If

            cell.Remove()
        End Using
    End Sub
    ' </Snippet1>

    ' <Snippet2>
    ' Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
    Function GetSpreadsheetCell(worksheet As Worksheet, columnName As String, rowIndex As UInteger) As Cell
        Dim rows As IEnumerable(Of Row) = worksheet.GetFirstChild(Of SheetData)()?.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex))
        If rows Is Nothing OrElse rows.Count() = 0 Then
            ' A cell does not exist at the specified row.
            Return Nothing
        End If

        Dim cells As IEnumerable(Of Cell) = rows.First().Elements(Of Cell)().Where(Function(c) String.Compare(c.CellReference?.Value, columnName & rowIndex, True) = 0)

        If cells.Count() = 0 Then
            ' A cell does not exist at the specified column, in the specified row.
            Return Nothing
        End If

        Return cells.FirstOrDefault()
    End Function
    ' </Snippet2>

    ' <Snippet3>
    ' Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    ' reference the specified SharedStringItem and removes the item.
    Sub RemoveSharedStringItem(shareStringId As Integer, document As SpreadsheetDocument)
        Dim remove As Boolean = True

        If document.WorkbookPart Is Nothing Then
            Return
        End If

        For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
            Dim cells = part.Worksheet.GetFirstChild(Of SheetData)()?.Descendants(Of Cell)()

            If cells Is Nothing Then
                Continue For
            End If

            For Each cell In cells
                ' Verify if other cells in the document reference the item.
                If cell.DataType IsNot Nothing AndAlso
                   cell.DataType.Value = CellValues.SharedString AndAlso
                   cell.CellValue?.Text = shareStringId.ToString() Then
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
                Return
            End If

            Dim item As SharedStringItem = shareStringTablePart.SharedStringTable.Elements(Of SharedStringItem)().ElementAt(shareStringId)
            If item IsNot Nothing Then
                item.Remove()

                ' Refresh all the shared string references.
                For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
                    Dim cells = part.Worksheet.GetFirstChild(Of SheetData)()?.Descendants(Of Cell)()

                    If cells Is Nothing Then
                        Continue For
                    End If

                    For Each cell In cells
                        Dim itemIndex As Integer

                        If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString AndAlso Integer.TryParse(cell.CellValue?.Text, itemIndex) Then
                            If itemIndex > shareStringId Then
                                cell.CellValue.Text = (itemIndex - 1).ToString()
                            End If
                        End If
                    Next
                    part.Worksheet.Save()
                Next

                document.WorkbookPart.SharedStringTablePart?.SharedStringTable.Save()
            End If
        End If
    End Sub
    ' </Snippet3>
    ' </Snippet0>
End Module
