Imports System.IO
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
    End Sub



    Public Sub OpenAndAddToSpreadsheetStream(ByVal stream As Stream)
        ' Open a SpreadsheetDocument based on a stream.
        Dim mySpreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(stream, True)

        ' Add a new worksheet.
        Dim newWorksheetPart As WorksheetPart = mySpreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData())
        newWorksheetPart.Worksheet.Save()

        Dim sheets As Sheets = mySpreadsheetDocument.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()
        Dim relationshipId As String = mySpreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart)

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
        mySpreadsheetDocument.WorkbookPart.Workbook.Save()

        'Dispose the document handle.
        mySpreadsheetDocument.Dispose()

        'Caller must close the stream.
    End Sub
End Module