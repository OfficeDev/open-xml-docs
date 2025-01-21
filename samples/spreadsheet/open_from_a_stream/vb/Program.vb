Imports System.IO
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        Using fileStream As New FileStream(args(0), FileMode.Open, FileAccess.ReadWrite)
            OpenAndAddToSpreadsheetStream(fileStream)
        End Using
        ' </Snippet2>
    End Sub

    ' <Snippet0>
    Sub OpenAndAddToSpreadsheetStream(stream As Stream)
        ' Open a SpreadsheetDocument based on a stream.
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(stream, True)

            If spreadsheetDocument IsNot Nothing Then
                ' Get or create the WorkbookPart
                Dim workbookPart As WorkbookPart = If(spreadsheetDocument.WorkbookPart, spreadsheetDocument.AddWorkbookPart())

                ' <Snippet1>

                ' Add a new worksheet.
                Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                newWorksheetPart.Worksheet = New Worksheet(New SheetData())

                ' </Snippet1>

                Dim workbook As Workbook = If(workbookPart.Workbook, New Workbook())

                If workbookPart.Workbook Is Nothing Then
                    workbookPart.Workbook = workbook
                End If

                Dim sheets As Sheets = If(workbook.GetFirstChild(Of Sheets)(), workbook.AppendChild(New Sheets()))
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
            End If
        End Using
    End Sub
    ' </Snippet0>
End Module
