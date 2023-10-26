Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    Public Sub OpenSpreadsheetDocumentReadonly(ByVal filepath As String)
        ' Open a SpreadsheetDocument based on a filepath.
        Using spreadsheetDocument As SpreadsheetDocument = spreadsheetDocument.Open(filepath, False)
            ' Attempt to add a new WorksheetPart.
            ' The call to AddNewPart generates an exception because the file is read-only.
            Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

            ' The rest of the code will not be called.
        End Using
    End Sub
End Module