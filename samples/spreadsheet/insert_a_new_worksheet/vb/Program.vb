Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.Linq

Module Program
    Sub Main(args As String())
        InsertWorksheet(args(0))
    End Sub

    ' Given a document name, inserts a new worksheet.
    ' <Snippet0>
    Sub InsertWorksheet(docName As String)
        ' <Snippet1>
        ' Open the document for editing.
        Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            ' </Snippet1>
            Dim workbookPart As WorkbookPart = If(spreadSheet.WorkbookPart, spreadSheet.AddWorkbookPart())
            ' Add a blank WorksheetPart.
            Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            newWorksheetPart.Worksheet = New Worksheet(New SheetData())

            Dim sheets As Sheets = If(workbookPart.Workbook.GetFirstChild(Of Sheets)(), workbookPart.Workbook.AppendChild(New Sheets()))
            Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)

            ' Get a unique ID for the new worksheet.
            Dim sheetId As UInteger = 1
            If sheets.Elements(Of Sheet)().Count() > 0 Then
                sheetId = sheets.Elements(Of Sheet)().Select(Function(s) s.SheetId?.Value).Max() + 1
            End If

            ' Give the new worksheet a name.
            Dim sheetName As String = "Sheet" & sheetId

            ' Append the new worksheet and associate it with the workbook.
            Dim sheet As New Sheet() With {
                .Id = relationshipId,
                .SheetId = sheetId,
                .Name = sheetName
            }
            sheets.Append(sheet)
        End Using
    End Sub
    ' </Snippet0>
End Module



