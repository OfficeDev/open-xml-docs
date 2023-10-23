Module Program `
  Sub Main(args As String())`
  End Sub`

  
    Public Sub CreateSpreadsheetWorkbookWithNumValue(ByVal filepath As String)
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        Dim spreadsheetDocument As SpreadsheetDocument = spreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)

        ' Add a WorkbookPart to the document.
        Dim workbookpart As WorkbookPart = SpreadsheetDocument.AddWorkbookPart()
        workbookpart.Workbook = New Workbook()

        ' Add a WorksheetPart to the WorkbookPart.
        Dim worksheetPart As WorksheetPart = workbookpart.AddNewPart(Of WorksheetPart)()
        worksheetPart.Worksheet = New Worksheet(New SheetData())

        ' Add Sheets to the Workbook.
        Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

        ' Append a new worksheet and associate it with the workbook.
        Dim sheet As New Sheet() With {.Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), .SheetId = 1, .Name = "mySheet"}
        sheets.Append(sheet)

        ' Get the sheetData cell table.
        Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

        ' Add a row to the cell table.
        Dim row As Row
        row = New Row() With {.RowIndex = 1}
        sheetData.Append(row)

        ' In the new row, find the column location to insert a cell in A1.  
        Dim refCell As Cell = Nothing
        For Each cell As Cell In row.Elements(Of Cell)()
            If String.Compare(cell.CellReference.Value, "A1", True) > 0 Then
                refCell = cell
                Exit For
            End If
        Next

        ' Add the cell to the cell table at A1.
        Dim newCell As New Cell() With {.CellReference = "A1"}
        row.InsertBefore(newCell, refCell)

        ' Set the cell value to be a numeric value of 100.
        newCell.CellValue = New CellValue("100")
        newCell.DataType = New EnumValue(Of CellValues)(CellValues.Number)

        ' Close the document.
        SpreadsheetDocument.Close()
    End Sub

    Public Sub CreateSpreadsheetWorkbookWithNumValue(ByVal filepath As String)
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        Dim spreadsheetDocument As SpreadsheetDocument = spreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)

        ' Add a WorkbookPart to the document.
        Dim workbookpart As WorkbookPart = SpreadsheetDocument.AddWorkbookPart()
        workbookpart.Workbook = New Workbook()

        ' Add a WorksheetPart to the WorkbookPart.
        Dim worksheetPart As WorksheetPart = workbookpart.AddNewPart(Of WorksheetPart)()
        worksheetPart.Worksheet = New Worksheet(New SheetData())

        ' Add Sheets to the Workbook.
        Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

        ' Append a new worksheet and associate it with the workbook.
        Dim sheet As New Sheet() With {.Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), .SheetId = 1, .Name = "mySheet"}
        sheets.Append(sheet)

        ' Get the sheetData cell table.
        Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

        ' Add a row to the cell table.
        Dim row As Row
        row = New Row() With {.RowIndex = 1}
        sheetData.Append(row)

        ' In the new row, find the column location to insert a cell in A1.  
        Dim refCell As Cell = Nothing
        For Each cell As Cell In row.Elements(Of Cell)()
            If String.Compare(cell.CellReference.Value, "A1", True) > 0 Then
                refCell = cell
                Exit For
            End If
        Next

        ' Add the cell to the cell table at A1.
        Dim newCell As New Cell() With {.CellReference = "A1"}
        row.InsertBefore(newCell, refCell)

        ' Set the cell value to be a numeric value of 100.
        newCell.CellValue = New CellValue("100")
        newCell.DataType = New EnumValue(Of CellValues)(CellValues.Number)

        ' Close the document.
        SpreadsheetDocument.Close()
    End Sub
End Module