Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        InsertText(args(0), args(1))
    End Sub

    ' <Snippet1>
    ' Given a document name and text, 
    ' inserts a new work sheet and writes the text to cell "A1" of the new worksheet.
    ' <Snippet0>
    Sub InsertText(docName As String, text As String)
        ' Open the document for editing.
        Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim workbookPart As WorkbookPart = If(spreadSheet.WorkbookPart, spreadSheet.AddWorkbookPart())

            ' Get the SharedStringTablePart. If it does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart
            If workbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0 Then
                shareStringPart = workbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
            Else
                shareStringPart = workbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the text into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(text, shareStringPart)

            ' Insert a new worksheet.
            Dim worksheetPart As WorksheetPart = InsertWorksheet(workbookPart)

            ' Insert cell A1 into the new worksheet.
            Dim cell As Cell = InsertCellInWorksheet("A", 1, worksheetPart)

            ' Set the value of cell A1.
            cell.CellValue = New CellValue(index.ToString())
            cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
        End Using
    End Sub
    ' </Snippet1>

    ' <Snippet2>
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Function InsertSharedStringItem(text As String, shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If shareStringPart.SharedStringTable Is Nothing Then
            shareStringPart.SharedStringTable = New SharedStringTable()
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If item.InnerText = text Then
                Return i
            End If

            i += 1
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))

        Return i
    End Function
    ' </Snippet2>

    ' <Snippet3>
    ' Given a WorkbookPart, inserts a new worksheet.
    Function InsertWorksheet(workbookPart As WorkbookPart) As WorksheetPart
        ' Add a new worksheet part to the workbook.
        Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData())

        Dim sheets As Sheets = If(workbookPart.Workbook.GetFirstChild(Of Sheets)(), workbookPart.Workbook.AppendChild(New Sheets()))
        Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)

        ' Get a unique ID for the new sheet.
        Dim sheetId As UInteger = 1
        If sheets.Elements(Of Sheet)().Count() > 0 Then
            sheetId = sheets.Elements(Of Sheet)().Select(Function(s)
                                                             If s.SheetId IsNot Nothing AndAlso s.SheetId.HasValue Then
                                                                 Return s.SheetId.Value
                                                             End If

                                                             Return 0
                                                         End Function).Max() + 1
        End If

        Dim sheetName As String = "Sheet" & sheetId

        ' Append the new worksheet and associate it with the workbook.
        Dim sheet As New Sheet() With {
            .Id = relationshipId,
            .SheetId = sheetId,
            .Name = sheetName
        }
        sheets.Append(sheet)

        Return newWorksheetPart
    End Function
    ' </Snippet3>


    ' <Snippet4>
    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, returns it. 
    Function InsertCellInWorksheet(columnName As String, rowIndex As UInteger, worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = columnName & rowIndex

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row

        If sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex)).Count() <> 0 Then
            row = sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex IsNot Nothing AndAlso r.RowIndex.Equals(rowIndex)).First()
        Else
            row = New Row() With {
                .RowIndex = rowIndex
            }
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If row.Elements(Of Cell)().Where(Function(c) c.CellReference IsNot Nothing AndAlso c.CellReference.Value = columnName & rowIndex).Count() > 0 Then
            Return row.Elements(Of Cell)().Where(Function(c) c.CellReference IsNot Nothing AndAlso c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing

            For Each cell As Cell In row.Elements(Of Cell)()
                If String.Compare(cell.CellReference?.Value, cellReference, True) > 0 Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As New Cell() With {
                .CellReference = cellReference
            }
            row.InsertBefore(newCell, refCell)

            Return newCell
        End If
    End Function
    ' </Snippet4>
    ' </Snippet0>
End Module
