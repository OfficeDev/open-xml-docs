Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
    End Sub



    ' Given a document name, inserts a new worksheet.
    Public Sub InsertWorksheet(ByVal docName As String)
        ' Open the document for editing.
        Dim spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (spreadSheet)
            ' Add a blank WorksheetPart.
            Dim newWorksheetPart As WorksheetPart = spreadSheet.WorkbookPart.AddNewPart(Of WorksheetPart)()
            newWorksheetPart.Worksheet = New Worksheet(New SheetData())
            ' newWorksheetPart.Worksheet.Save()

            Dim sheets As Sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()
            Dim relationshipId As String = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart)

            ' Get a unique ID for the new worksheet.
            Dim sheetId As UInteger = 1
            If (sheets.Elements(Of Sheet).Count > 0) Then
                sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max + 1
            End If

            ' Give the new worksheet a name.
            Dim sheetName As String = ("Sheet" + sheetId.ToString())

            ' Append the new worksheet and associate it with the workbook.
            Dim sheet As Sheet = New Sheet
            sheet.Id = relationshipId
            sheet.SheetId = sheetId
            sheet.Name = sheetName
            sheets.Append(sheet)
        End Using
    End Sub
End Module