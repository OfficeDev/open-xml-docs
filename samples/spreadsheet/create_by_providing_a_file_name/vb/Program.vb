Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        CreateSpreadsheetWorkbook(args(0))
    End Sub

    ' <Snippet0>
    Sub CreateSpreadsheetWorkbook(filepath As String)
        ' <Snippet1>
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)
            ' </Snippet1>

            ' <Snippet2>
            ' Add a WorkbookPart to the document.
            Dim workbookPart As WorkbookPart = spreadsheetDocument.AddWorkbookPart()
            workbookPart.Workbook = New Workbook()
            ' </Snippet2>

            ' <Snippet3>
            ' Add a WorksheetPart to the WorkbookPart.
            Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            worksheetPart.Worksheet = New Worksheet(New SheetData())

            ' Add Sheets to the Workbook.
            Dim sheets As Sheets = workbookPart.Workbook.AppendChild(New Sheets())

            ' Append a new worksheet and associate it with the workbook.
            Dim sheet As New Sheet() With {
                .Id = workbookPart.GetIdOfPart(worksheetPart),
                .SheetId = 1,
                .Name = "mySheet"
            }
            sheets.Append(sheet)
            ' </Snippet3>

            workbookPart.Workbook.Save()
        End Using
    End Sub
    ' </Snippet0>
End Module
