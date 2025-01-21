Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        Dim sheets = GetHiddenSheets(args(0))

        For Each sheet In sheets
            Console.WriteLine(sheet.Name)
        Next
    End Sub

    ' <Snippet0>
    Function GetHiddenSheets(fileName As String) As List(Of Sheet)
        Dim returnVal As New List(Of Sheet)()

        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' <Snippet1>
            Dim wbPart As WorkbookPart = document.WorkbookPart

            If wbPart IsNot Nothing Then
                Dim sheets = wbPart.Workbook.Descendants(Of Sheet)()
                ' </Snippet1>

                ' Look for sheets where there is a State attribute defined, 
                ' where the State has a value,
                ' and where the value is either Hidden or VeryHidden.

                ' <Snippet2>
                Dim hiddenSheets = sheets.Where(Function(item) item.State IsNot Nothing AndAlso
                    item.State.HasValue AndAlso
                    (item.State.Value = SheetStateValues.Hidden OrElse
                    item.State.Value = SheetStateValues.VeryHidden))
                ' </Snippet2>

                returnVal = hiddenSheets.ToList()
            End If
        End Using

        Return returnVal
    End Function
    ' </Snippet0>
End Module
