Imports System.Linq
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet


Module MyModule
' Given a document name and text, 
    ' inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
    Public Function InsertText(ByVal docName As String, ByVal text As String)
        ' Open the document for editing.
        Dim spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (spreadSheet)
            ' Get the SharedStringTablePart. If it does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart

            If (spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).Count() > 0) Then
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).First()
            Else
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the text into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(text, shareStringPart)

            ' Insert a new worksheet.
            Dim worksheetPart As WorksheetPart = InsertWorksheet(spreadSheet.WorkbookPart)

            ' Insert cell A1 into the new worksheet.
            Dim cell As Cell = InsertCellInWorksheet("A", 1, worksheetPart)

            ' Set the value of cell A1.
            cell.CellValue = New CellValue(index.ToString)
            cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)

            ' Save the new worksheet.
            worksheetPart.Worksheet.Save()

            Return 0
        End Using
    End Function

    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function

    ' Given a WorkbookPart, inserts a new worksheet.
    Private Function InsertWorksheet(ByVal workbookPart As WorkbookPart) As WorksheetPart
        ' Add a new worksheet part to the workbook.
        Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData)
        newWorksheetPart.Worksheet.Save()
        Dim sheets As Sheets = workbookPart.Workbook.GetFirstChild(Of Sheets)()
        Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)

        ' Get a unique ID for the new sheet.
        Dim sheetId As UInteger = 1
        If (sheets.Elements(Of Sheet).Count() > 0) Then
            sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max() + 1
        End If

        Dim sheetName As String = ("Sheet" + sheetId.ToString())

        ' Add the new worksheet and associate it with the workbook.
        Dim sheet As Sheet = New Sheet
        sheet.Id = relationshipId
        sheet.SheetId = sheetId
        sheet.Name = sheetName
        sheets.Append(sheet)
        workbookPart.Workbook.Save()

        Return newWorksheetPart
    End Function

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
End Module
