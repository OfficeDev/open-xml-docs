Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        CreateSpreadsheetWorkbook(args(0))
    End Sub

    ' <Snippet0>
    Sub CreateSpreadsheetWorkbook(filepath As String)
        ' Use 'Using' block to ensure proper disposal of the document
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)
            ' Add a WorkbookPart to the document.
            Dim workbookPart As WorkbookPart = spreadsheetDocument.AddWorkbookPart()
            workbookPart.Workbook = New Workbook()

            ' Add a WorksheetPart to the WorkbookPart.
            Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            worksheetPart.Worksheet = New Worksheet(New SheetData())

            ' Add Sheets to the Workbook.
            Dim sheets As Sheets = workbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

            ' Append a new worksheet and associate it with the workbook.
            Dim sheet As Sheet = New Sheet() With {
                .Id = workbookPart.GetIdOfPart(worksheetPart),
                .SheetId = 1,
                .Name = "mySheet"
            }
            sheets.Append(sheet)

            ' Get the sheetData cell table.
            Dim sheetData As SheetData = If(worksheetPart.Worksheet.GetFirstChild(Of SheetData)(), worksheetPart.Worksheet.AppendChild(Of SheetData)(New SheetData()))

            ' Add a row to the cell table.
            Dim row As Row = New Row() With {
                .RowIndex = 1
            }
            sheetData.Append(row)

            ' In the new row, find the column location to insert a cell in A1.  
            Dim refCell As Cell = Nothing

            For Each cell As Cell In row.Elements(Of Cell)()
                If String.Compare(cell.CellReference?.Value, "A1", True) > 0 Then
                    refCell = cell
                    Exit For
                End If
            Next

            ' Add the cell to the cell table at A1.
            Dim newCell As Cell = New Cell() With {
                .CellReference = "A1"
            }
            row.InsertBefore(newCell, refCell)

            ' Set the cell value to be a numeric value of 100.
            newCell.CellValue = New CellValue("100")
            newCell.DataType = New EnumValue(Of CellValues)(CellValues.Number)
        End Using
    End Sub
    ' </Snippet0>
End Module
